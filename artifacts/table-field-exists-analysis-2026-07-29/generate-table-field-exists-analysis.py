from __future__ import annotations

import csv
import hashlib
import re
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import List, Optional, Tuple, Dict


REPO_ROOT = Path(__file__).resolve().parents[2]
MODULE_DIR = REPO_ROOT / "src" / "access" / "exported" / "modules"
OUT_DIR = Path(__file__).resolve().parent

TARGET_NAMES = {"TableExists", "FieldExists"}


@dataclass
class ProcedureDef:
    module_name: str
    module_path: str
    procedure_name: str
    visibility: str
    signature: str
    start_line: int
    end_line: int
    body: str
    uses_currentdb: bool
    uses_explicit_db_param: bool
    has_error_handling: bool
    return_behavior: str
    empty_name_handling: str
    missing_behavior: str
    body_hash: str
    signature_key: str
    exact_group_key: str


@dataclass
class CallSite:
    procedure_name: str
    target_module: str
    target_signature: str
    caller_module: str
    line_number: int
    line_text: str
    call_style: str
    arguments: str
    needs_explicit_db: bool
    uses_currentdb_context: bool


def read_text(path: Path) -> List[str]:
    return path.read_text(encoding="utf-8").splitlines()


def strip_inline_comment(line: str) -> str:
    in_string = False
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            if in_string and i + 1 < len(line) and line[i + 1] == '"':
                i += 2
                continue
            in_string = not in_string
            i += 1
            continue
        if not in_string and ch == "'":
            return line[:i]
        i += 1
    return line


def collapse_continuations(lines: List[str]) -> List[Tuple[int, int, str]]:
    result: List[Tuple[int, int, str]] = []
    buffer = ""
    start_line = 0
    for idx, line in enumerate(lines, start=1):
        if not buffer:
            buffer = line
            start_line = idx
        else:
            buffer += "\n" + line
        trimmed = line.rstrip()
        if trimmed.endswith(" _"):
            buffer = buffer[:-2]
            continue
        result.append((start_line, idx, buffer))
        buffer = ""
        start_line = 0
    if buffer:
        result.append((start_line, len(lines), buffer))
    return result


def normalize_signature(signature: str) -> str:
    return re.sub(r"\s+", " ", signature.replace("\n", " ")).strip()


def normalize_body_for_hash(body: str) -> str:
    normalized_lines: List[str] = []
    for raw_line in body.splitlines():
        clean = strip_inline_comment(raw_line).strip()
        if clean:
            normalized_lines.append(re.sub(r"\s+", " ", clean.lower()))
    return "\n".join(normalized_lines)


def detect_return_behavior(body: str, proc_name: str) -> str:
    clean = normalize_body_for_hash(body)
    if re.search(rf"\b{re.escape(proc_name.lower())}\s*=\s*true\b", clean):
        if re.search(rf"\b{re.escape(proc_name.lower())}\s*=\s*false\b", clean):
            return "Assigns True/False explicitly"
        return "Assigns True explicitly"
    if re.search(rf"\b{re.escape(proc_name.lower())}\s*=\s*false\b", clean):
        return "Assigns False explicitly"
    if "exit function" in clean:
        return "Implicit default False with early Exit Function"
    return "Implicit default return"


def detect_empty_name_handling(body: str) -> str:
    clean = normalize_body_for_hash(body)
    if re.search(r"lenb?\s*\(\s*trim\$\(", clean) or "len(" in clean:
        if "exit function" in clean:
            return "Checks empty/trimmed input and exits early"
        return "Checks empty/trimmed input"
    return "No explicit empty-name guard"


def detect_missing_behavior(body: str) -> str:
    clean = normalize_body_for_hash(body)
    if "for each tdf in db.tabledefs" in clean or "for each fld in db.tabledefs" in clean:
        return "Iterates DAO collection; returns False when not found"
    if "tabledefs" in clean and "fields" in clean:
        return "Traverses TableDefs/Fields; returns False when not found"
    return "Manual review"


def parse_procedures(module_path: Path) -> List[ProcedureDef]:
    module_name = module_path.stem
    lines = read_text(module_path)
    logical_lines = collapse_continuations(lines)
    procedures: List[ProcedureDef] = []
    current = None
    body_lines: List[str] = []

    for start_line, end_line, text in logical_lines:
        if current is None:
            match = re.match(r"^\s*(Public|Private)\s+Function\s+(TableExists|FieldExists)\b", text, re.IGNORECASE)
            if match:
                current = {
                    "visibility": match.group(1),
                    "procedure_name": match.group(2),
                    "signature": normalize_signature(text),
                    "start_line": start_line,
                }
                body_lines = [text]
            continue

        body_lines.append(text)
        if re.match(rf"^\s*End\s+Function\s*$", text, re.IGNORECASE):
            body = "\n".join(body_lines)
            proc_name = current["procedure_name"]
            normalized_body = normalize_body_for_hash(body)
            signature_key = re.sub(r"\s+", " ", current["signature"].lower())
            body_hash = hashlib.sha1(normalized_body.encode("utf-8")).hexdigest()
            exact_group_key = f"{proc_name.lower()}|{signature_key}|{body_hash}"
            procedures.append(
                ProcedureDef(
                    module_name=module_name,
                    module_path=str(module_path),
                    procedure_name=proc_name,
                    visibility=current["visibility"],
                    signature=current["signature"],
                    start_line=current["start_line"],
                    end_line=end_line,
                    body=body,
                    uses_currentdb=bool(re.search(r"\bCurrentDb\b", body, re.IGNORECASE)),
                    uses_explicit_db_param=bool(re.search(r"\bByVal\s+db\s+As\s+DAO\.Database\b", current["signature"], re.IGNORECASE)),
                    has_error_handling=bool(re.search(r"^\s*On\s+Error\b", body, re.IGNORECASE | re.MULTILINE) or re.search(r"\bResume\b", body, re.IGNORECASE)),
                    return_behavior=detect_return_behavior(body, proc_name),
                    empty_name_handling=detect_empty_name_handling(body),
                    missing_behavior=detect_missing_behavior(body),
                    body_hash=body_hash,
                    signature_key=signature_key,
                    exact_group_key=exact_group_key,
                )
            )
            current = None
            body_lines = []

    return procedures


def definition_line_match(line: str) -> bool:
    return bool(re.match(r"^\s*(Public|Private)\s+Function\s+(TableExists|FieldExists)\b", line, re.IGNORECASE))


def declaration_like(line: str) -> bool:
    return bool(re.match(r"^\s*(Dim|Const|Private|Public|Static|Function|Sub|Property|Type|Enum)\b", line, re.IGNORECASE))


def assignment_to_function(line: str, name: str) -> bool:
    return bool(re.match(rf"^\s*{re.escape(name)}\s*=", line, re.IGNORECASE))


def extract_argument_slice(line: str, start_index: int) -> str:
    open_idx = line.find("(", start_index)
    if open_idx < 0:
        return ""
    depth = 0
    in_string = False
    for idx in range(open_idx, len(line)):
        ch = line[idx]
        if ch == '"':
            if in_string and idx + 1 < len(line) and line[idx + 1] == '"':
                continue
            in_string = not in_string
        elif not in_string:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    return line[open_idx + 1 : idx].strip()
    return ""


def find_callsites(procedures: List[ProcedureDef]) -> List[CallSite]:
    by_module: Dict[str, Dict[str, ProcedureDef]] = {}
    for proc in procedures:
        by_module.setdefault(proc.module_name, {})[proc.procedure_name.lower()] = proc

    callsites: List[CallSite] = []
    for module_path in sorted(MODULE_DIR.glob("*.bas")):
        module_name = module_path.stem
        lines = read_text(module_path)
        for line_no, raw_line in enumerate(lines, start=1):
            clean = strip_inline_comment(raw_line)
            if not clean.strip():
                continue
            if definition_line_match(clean) or declaration_like(clean):
                continue

            for proc_name in TARGET_NAMES:
                if assignment_to_function(clean, proc_name):
                    continue

                qualified = re.search(rf"\b([A-Za-z_][A-Za-z0-9_]*)\.{proc_name}\s*\(", clean, re.IGNORECASE)
                unqualified = re.search(rf"\b{proc_name}\s*\(", clean, re.IGNORECASE)

                if qualified:
                    target_module_name = qualified.group(1)
                    start_idx = qualified.start()
                    arguments = extract_argument_slice(clean, qualified.start())
                    target_proc = None
                    if target_module_name in by_module:
                        target_proc = by_module[target_module_name].get(proc_name.lower())
                    if target_proc is None:
                        continue
                    callsites.append(
                        CallSite(
                            procedure_name=proc_name,
                            target_module=target_proc.module_name,
                            target_signature=target_proc.signature,
                            caller_module=module_name,
                            line_number=line_no,
                            line_text=clean.strip(),
                            call_style="qualified",
                            arguments=arguments,
                            needs_explicit_db=target_proc.uses_explicit_db_param,
                            uses_currentdb_context=bool(re.search(r"\bCurrentDb\b", arguments, re.IGNORECASE)),
                        )
                    )
                    continue

                if unqualified:
                    if module_name not in by_module or proc_name.lower() not in by_module[module_name]:
                        continue
                    target_proc = by_module[module_name][proc_name.lower()]
                    arguments = extract_argument_slice(clean, unqualified.start())
                    callsites.append(
                        CallSite(
                            procedure_name=proc_name,
                            target_module=target_proc.module_name,
                            target_signature=target_proc.signature,
                            caller_module=module_name,
                            line_number=line_no,
                            line_text=clean.strip(),
                            call_style="unqualified-local",
                            arguments=arguments,
                            needs_explicit_db=target_proc.uses_explicit_db_param,
                            uses_currentdb_context=bool(re.search(r"\bCurrentDb\b", arguments, re.IGNORECASE)),
                        )
                    )

    return callsites


def write_csv(path: Path, rows: List[dict]) -> None:
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    procedures: List[ProcedureDef] = []
    for module_path in sorted(MODULE_DIR.glob("*.bas")):
        procedures.extend(parse_procedures(module_path))

    callsites = find_callsites(procedures)

    definitions_rows = []
    implementation_groups: Dict[str, List[ProcedureDef]] = {}
    for proc in procedures:
        implementation_groups.setdefault(proc.exact_group_key, []).append(proc)
        proc_calls = [c for c in callsites if c.target_module == proc.module_name and c.target_signature == proc.signature and c.procedure_name.lower() == proc.procedure_name.lower()]
        definitions_rows.append(
            {
                "procedure_name": proc.procedure_name,
                "module_name": proc.module_name,
                "start_line": proc.start_line,
                "end_line": proc.end_line,
                "visibility": proc.visibility,
                "signature": proc.signature,
                "uses_currentdb": proc.uses_currentdb,
                "uses_explicit_dao_database": proc.uses_explicit_db_param,
                "has_error_handling": proc.has_error_handling,
                "return_behavior": proc.return_behavior,
                "empty_name_handling": proc.empty_name_handling,
                "missing_behavior": proc.missing_behavior,
                "actual_call_count": len(proc_calls),
                "body": proc.body,
            }
        )

    callsite_rows = []
    for call in callsites:
        callsite_rows.append(asdict(call))

    diff_rows = []
    for proc_name in sorted(TARGET_NAMES):
        exact_buckets: Dict[str, List[ProcedureDef]] = {}
        for proc in [p for p in procedures if p.procedure_name == proc_name]:
            exact_buckets.setdefault(proc.exact_group_key, []).append(proc)

        idx = 1
        for key, bucket in sorted(exact_buckets.items(), key=lambda item: (item[1][0].signature, item[1][0].module_name)):
            group_id = f"{proc_name}-EXACT-{idx:03d}"
            idx += 1
            signatures = sorted({b.signature for b in bucket})
            modules = sorted({b.module_name for b in bucket})
            diff_rows.append(
                {
                    "group_id": group_id,
                    "procedure_name": proc_name,
                    "group_type": "EXACT",
                    "definition_count": len(bucket),
                    "modules": "; ".join(modules),
                    "signatures": " || ".join(signatures),
                    "uses_currentdb": "; ".join(sorted({str(b.uses_currentdb) for b in bucket})),
                    "uses_explicit_dao_database": "; ".join(sorted({str(b.uses_explicit_db_param) for b in bucket})),
                    "has_error_handling": "; ".join(sorted({str(b.has_error_handling) for b in bucket})),
                    "return_behavior": " || ".join(sorted({b.return_behavior for b in bucket})),
                    "empty_name_handling": " || ".join(sorted({b.empty_name_handling for b in bucket})),
                    "missing_behavior": " || ".join(sorted({b.missing_behavior for b in bucket})),
                }
            )

        sig_groups: Dict[str, List[ProcedureDef]] = {}
        for proc in [p for p in procedures if p.procedure_name == proc_name]:
            sig_groups.setdefault(proc.signature_key, []).append(proc)
        sig_idx = 1
        for _, bucket in sorted(sig_groups.items(), key=lambda item: item[1][0].signature):
            if len({b.body_hash for b in bucket}) > 1:
                diff_rows.append(
                    {
                        "group_id": f"{proc_name}-SAME-SIGNATURE-DIFF-{sig_idx:03d}",
                        "procedure_name": proc_name,
                        "group_type": "SAME_SIGNATURE_DIFFERENT_BODY",
                        "definition_count": len(bucket),
                        "modules": "; ".join(sorted({b.module_name for b in bucket})),
                        "signatures": bucket[0].signature,
                        "uses_currentdb": "; ".join(sorted({str(b.uses_currentdb) for b in bucket})),
                        "uses_explicit_dao_database": "; ".join(sorted({str(b.uses_explicit_db_param) for b in bucket})),
                        "has_error_handling": "; ".join(sorted({str(b.has_error_handling) for b in bucket})),
                        "return_behavior": " || ".join(sorted({b.return_behavior for b in bucket})),
                        "empty_name_handling": " || ".join(sorted({b.empty_name_handling for b in bucket})),
                        "missing_behavior": " || ".join(sorted({b.missing_behavior for b in bucket})),
                    }
                )
                sig_idx += 1

    table_defs = [p for p in procedures if p.procedure_name == "TableExists"]
    field_defs = [p for p in procedures if p.procedure_name == "FieldExists"]

    call_summary_lines = []
    for proc in procedures:
        proc_calls = [c for c in callsites if c.target_module == proc.module_name and c.target_signature == proc.signature and c.procedure_name == proc.procedure_name]
        call_summary_lines.append(
            f"- {proc.procedure_name} | {proc.module_name}:{proc.start_line}-{proc.end_line} | {len(proc_calls)} call(s)"
        )

    exact_group_lines = []
    for row in diff_rows:
        exact_group_lines.append(
            f"- {row['group_id']}: {row['procedure_name']} | {row['group_type']} | {row['definition_count']} definition(s) | {row['modules']}"
        )

    removal_candidates = []
    change_candidates = []
    for proc in procedures:
        if proc.signature.startswith("Private Function TableExists(ByVal db As DAO.Database"):
            removal_candidates.append(f"- {proc.module_name}:{proc.start_line}-{proc.end_line} {proc.signature}")
            change_candidates.append(f"- Module {proc.module_name}: unqualified local calls to {proc.signature} -> future central schema helper with explicit db")
        elif proc.signature.startswith("Private Function TableExists(ByVal tableName") or proc.signature.startswith("Private Function TableExists(ByVal table_name"):
            removal_candidates.append(f"- {proc.module_name}:{proc.start_line}-{proc.end_line} {proc.signature}")
            change_candidates.append(f"- Module {proc.module_name}: unqualified local calls to {proc.signature} -> future central schema helper for CurrentDb callers")
        elif proc.signature.startswith("Private Function FieldExists(ByVal db As DAO.Database"):
            removal_candidates.append(f"- {proc.module_name}:{proc.start_line}-{proc.end_line} {proc.signature}")
            change_candidates.append(f"- Module {proc.module_name}: unqualified local calls to {proc.signature} -> future central schema helper with explicit db")
        elif proc.signature.startswith("Private Function FieldExists(ByVal tableName") or proc.signature.startswith("Private Function FieldExists(ByVal table_name"):
            removal_candidates.append(f"- {proc.module_name}:{proc.start_line}-{proc.end_line} {proc.signature}")
            change_candidates.append(f"- Module {proc.module_name}: unqualified local calls to {proc.signature} -> future central schema helper for CurrentDb callers")

    md = [
        "# TableExists / FieldExists Refactoring Plan",
        "",
        "## Counts",
        f"- Actual TableExists definitions: {len(table_defs)}",
        f"- Actual FieldExists definitions: {len(field_defs)}",
        "",
        "## Actual call counts per definition",
        *call_summary_lines,
        "",
        "## Global identical implementation groups",
        *exact_group_lines,
        "",
        "## Functional differences",
        "- There are two main API families for each function: CurrentDb-based and explicit DAO.Database-based.",
        "- Name-only variants differ mostly in parameter naming (`tableName` vs `table_name`) and occasional error-handling phrasing.",
        "- Explicit `DAO.Database` variants are structurally closest to a future shared schema helper.",
        "- CurrentDb variants are convenient for repository-style modules but couple the helper to ambient database state.",
        "",
        "## Recommended public signatures",
        "- Preferred readability-first option: `Public Function TableExists(ByVal tableName As String, ByVal db As DAO.Database) As Boolean`",
        "- Preferred readability-first option: `Public Function FieldExists(ByVal tableName As String, ByVal fieldName As String, ByVal db As DAO.Database) As Boolean`",
        "- Assessment: an optional `DAO.Database` parameter is technically possible in VBA only as an optional `Variant/Object` pattern, but it is less explicit and less readable than a required database argument.",
        "- Recommended CurrentDb caller pattern: caller resolves `CurrentDb` explicitly and passes `db` into the canonical helper.",
        "",
        "## Recommended target module",
        "- Preferred target module: `modDbSchema`",
        "- Reason: both helpers are schema-/metadata-oriented and should live beside other table/index/field inspection utilities rather than in a broad generic db access module.",
        "",
        "## Later removal candidates",
        *sorted(removal_candidates),
        "",
        "## Later call-site changes",
        *sorted(set(change_candidates)),
        "",
        "## Risks",
        "- Private local helpers currently resolve unqualified calls inside their own module; migration must update those call sites deliberately.",
        "- CurrentDb-based and explicit-db-based call sites should not be merged blindly without deciding transaction and backend-routing expectations.",
        "- Schema modules that operate on linked/back-end databases may require explicit database objects to avoid regressions.",
        "- Any future public helper should preserve current False-on-missing behavior and existing error-handling conventions where relied upon.",
    ]

    write_csv(OUT_DIR / "table-field-exists-definitions.csv", definitions_rows)
    write_csv(OUT_DIR / "table-field-exists-call-sites.csv", callsite_rows)
    write_csv(OUT_DIR / "table-field-exists-implementation-differences.csv", diff_rows)
    (OUT_DIR / "table-field-exists-refactoring-plan.md").write_text("\n".join(md), encoding="utf-8")

    print(f"TABLE_DEFINITIONS={len(table_defs)}")
    print(f"FIELD_DEFINITIONS={len(field_defs)}")
    print(f"CALLS={len(callsites)}")
    print(f"OUTDIR={OUT_DIR}")


if __name__ == "__main__":
    main()
