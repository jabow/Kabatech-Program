Attribute VB_Name = "Individual"
Option Explicit
Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Integer) As Integer



'Callback for btnStart onAction
Public Sub Ind_Start(control As IRibbonControl)
    If Standard.start() = False Then GoTo error
error:
End Sub

'Callback for btnHoodOpen onAction
Public Sub Ind_Hood(control As IRibbonControl)
    If Standard.hood_open() = False Then GoTo error
error:
End Sub

'Callback for btnAdapter onAction
Public Sub Ind_Adapter(control As IRibbonControl)
    formAdapter.Show
End Sub

'Callback for btnRollers onAction
Public Sub Ind_Rollers(control As IRibbonControl)
    formRollers.Show
End Sub

'Callback for btnWWF onAction
Public Sub Ind_WWF(control As IRibbonControl)

formWWF.Show

End Sub

'Callback for btnWWOF onAction
Public Sub IndWWOF(control As IRibbonControl)

formWWOF.Show

End Sub

'Callback for btnFWOW onAction
Public Sub Ind_FWOW(control As IRibbonControl)
formFWW.Show
End Sub

'Callback for btnLOM onAction
Public Sub Ind_LOM(control As IRibbonControl)
    If Standard.line_off_marker() = False Then GoTo error
error:
End Sub

'Callback for btnPositiom onAction
Public Sub Ind_Position(control As IRibbonControl)

    formPosition.Show

End Sub

'Callback for btnCutting onAction
Public Sub Ind_Cutting(control As IRibbonControl)
    Standard.cutting (1)
End Sub

'Callback for btnClamping onAction
Public Sub Ind_Clamping(control As IRibbonControl)

formClamping.Show

End Sub

