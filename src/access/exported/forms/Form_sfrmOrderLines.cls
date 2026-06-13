VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmOrderLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ExportKind: CODE_BEHIND_ONLY
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_sfrmOrderLines
' Purpose   : Editable sales-order line subform with lightweight recalculation.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_sfrmOrderLines"
Private Const TABLE_ORD_ORDER_LINE As String = "ord_order_line"
Private Const FIELD_ORDER_LINE_ID As String = "order_line_id"
Private Const FIELD_ORDER_ID As String = "order_id"
Private Const FIELD_LINE_NO As String = "line_no"
Private Const FIELD_DISCOUNT_TYPE As String = "discount_type"
Private Const FIELD_SURCHARGE_TYPE As String = "surcharge_type"

Private m_pendingDeleteOrderId As Long

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
        "parent_name='" & ResolveParentNameForLogging() & _
        "'; parent_order_id=" & CStr(GetParentOrderIdForLogging()) & "."

    InitializeRecordSource
    ConfigureEmptyLinesBehavior
    ConfigureAdjustmentCombo "cboDiscountType", FIELD_DISCOUNT_TYPE
    ConfigureAdjustmentCombo "cboSurchargeType", FIELD_SURCHARGE_TYPE
    modFwTranslationRuntime.ApplyTranslations Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    m_pendingDeleteOrderId = GetCurrentOrderId()
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo ErrorHandler

    EnsureNewLineDefaults
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_BeforeInsert", Err
End Sub

Private Sub Form_AfterInsert()
    On Error GoTo ErrorHandler

    RecalculateParentOrder False
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_AfterInsert", Err
End Sub

Private Sub Form_AfterUpdate()
    On Error GoTo ErrorHandler

    RecalculateParentOrder False
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_AfterUpdate", Err
End Sub

Private Sub Form_Delete(Cancel As Integer)
    On Error GoTo ErrorHandler

    m_pendingDeleteOrderId = GetCurrentOrderId()
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Delete", Err
End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)
    On Error GoTo ErrorHandler

    If Status = acDeleteOK Then
        RecalculateParentOrder True
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_AfterDelConfirm", Err
End Sub

Private Sub InitializeRecordSource()
    On Error GoTo ErrorHandler

    If LenB(Trim$(modDaoHelper.NzString(Me.RecordSource))) = 0 Then
        Me.RecordSource = TABLE_ORD_ORDER_LINE
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "InitializeRecordSource", Err
End Sub

Private Sub ConfigureEmptyLinesBehavior()
    On Error GoTo ErrorHandler

    ' A sales order without lines is valid. Keep the subform bound and editable
    ' so the user can add the first line directly into an empty detail section.
    Me.AllowAdditions = True
    Me.DataEntry = False
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureEmptyLinesBehavior", Err
End Sub

Private Sub ConfigureAdjustmentCombo(ByVal controlName As String, ByVal controlSource As String)
    On Error GoTo SafeExit

    If Not HasControl(controlName) Then
        Exit Sub
    End If

    With Me.Controls(controlName)
        .ControlSource = controlSource
        .RowSourceType = "Value List"
        .RowSource = "NONE;PERCENT;AMOUNT"
        .BoundColumn = 1
        .ColumnCount = 1
        .LimitToList = True
    End With

SafeExit:
End Sub

Private Sub EnsureNewLineDefaults()
    On Error GoTo ErrorHandler

    Dim parentOrderId As Long

    parentOrderId = GetParentOrderId()
    If parentOrderId <= 0 Then
        Exit Sub
    End If

    If HasControl(FIELD_ORDER_ID) Then
        If modDaoHelper.NzLong(Me.Controls(FIELD_ORDER_ID).Value, 0) <= 0 Then
            Me.Controls(FIELD_ORDER_ID).Value = parentOrderId
        End If
    End If

    If HasControl(FIELD_LINE_NO) Then
        If modDaoHelper.NzLong(Me.Controls(FIELD_LINE_NO).Value, 0) <= 0 Then
            Me.Controls(FIELD_LINE_NO).Value = GetNextLineNo(parentOrderId)
        End If
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureNewLineDefaults", Err
End Sub

Private Function GetNextLineNo(ByVal OrderId As Long) As Long
    On Error GoTo ErrorHandler

    Dim nextValue As Variant

    If OrderId <= 0 Then
        GetNextLineNo = 1
        Exit Function
    End If

    nextValue = DMax(FIELD_LINE_NO, TABLE_ORD_ORDER_LINE, "[" & FIELD_ORDER_ID & "]=" & CStr(OrderId))
    GetNextLineNo = modDaoHelper.NzLong(nextValue, 0) + 1
    Exit Function

ErrorHandler:
    GetNextLineNo = 1
    modErrorHandler.HandleError MODULE_NAME, "GetNextLineNo", Err
End Function

Private Sub RecalculateParentOrder(ByVal preferPendingDeleteId As Boolean)
    On Error GoTo ErrorHandler

    Dim OrderId As Long
    Dim parentForm As Access.Form

    If preferPendingDeleteId And m_pendingDeleteOrderId > 0 Then
        OrderId = m_pendingDeleteOrderId
    Else
        OrderId = GetCurrentOrderId()
        If OrderId <= 0 Then
            OrderId = GetParentOrderId()
        End If
    End If

    If OrderId <= 0 Then
        Exit Sub
    End If

    Call modOrderCalculationService.RecalculateOrder(OrderId)

    Set parentForm = ResolveParentOrderForm()
    If Not parentForm Is Nothing Then
        CallByName parentForm, "RefreshCalculatedTotals", VbMethod, True
    End If

    m_pendingDeleteOrderId = OrderId
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RecalculateParentOrder", Err
End Sub

Private Function GetCurrentOrderId() As Long
    On Error GoTo SafeExit

    If Me.NewRecord Then
        Exit Function
    End If

    If HasControl(FIELD_ORDER_ID) Then
        GetCurrentOrderId = modDaoHelper.NzLong(Me.Controls(FIELD_ORDER_ID).Value, 0)
    End If

SafeExit:
End Function

Private Function GetParentOrderId() As Long
    On Error GoTo SafeExit

    Dim parentForm As Access.Form

    Set parentForm = ResolveParentOrderForm()
    If parentForm Is Nothing Then
        Exit Function
    End If

    If HasFormControl(parentForm, "txtOrderId") Then
        GetParentOrderId = modDaoHelper.NzLong(parentForm.Controls("txtOrderId").Value, 0)
    ElseIf HasFormControl(parentForm, FIELD_ORDER_ID) Then
        GetParentOrderId = modDaoHelper.NzLong(parentForm.Controls(FIELD_ORDER_ID).Value, 0)
    Else
        GetParentOrderId = modDaoHelper.NzLong(CallByName(parentForm, "GetCurrentOrderDetailId", VbMethod), 0)
    End If

SafeExit:
End Function

Private Function ResolveParentOrderForm() As Access.Form
    On Error GoTo SafeExit

    Dim parentForm As Access.Form

    Set parentForm = Me.Parent
    If parentForm Is Nothing Then
        Exit Function
    End If

    If StrComp(parentForm.Name, "frmOrderDetail", vbTextCompare) = 0 Then
        Set ResolveParentOrderForm = parentForm
    End If

SafeExit:
End Function

Private Function ResolveParentNameForLogging() As String
    On Error GoTo SafeExit

    If Not Me.Parent Is Nothing Then
        ResolveParentNameForLogging = Me.Parent.Name
    End If

SafeExit:
    If LenB(ResolveParentNameForLogging) = 0 Then
        ResolveParentNameForLogging = "<none>"
    End If
End Function

Private Function GetParentOrderIdForLogging() As Long
    On Error GoTo SafeExit

    GetParentOrderIdForLogging = GetParentOrderId()

SafeExit:
End Function

Private Function HasControl(ByVal ControlName As String) As Boolean
    On Error GoTo SafeExit

    Dim currentControl As Control

    For Each currentControl In Me.Controls
        If StrComp(currentControl.Name, ControlName, vbTextCompare) = 0 Then
            HasControl = True
            Exit Function
        End If
    Next currentControl

SafeExit:
End Function

Private Function HasFormControl(ByVal formInstance As Access.Form, ByVal controlName As String) As Boolean
    On Error GoTo SafeExit

    Dim currentControl As Control

    For Each currentControl In formInstance.Controls
        If StrComp(currentControl.Name, controlName, vbTextCompare) = 0 Then
            HasFormControl = True
            Exit Function
        End If
    Next currentControl

SafeExit:
End Function
