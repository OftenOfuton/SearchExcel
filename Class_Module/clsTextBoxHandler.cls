Option Explicit

' clsTextBoxHandler クラスモジュール
Public WithEvents TextBox As MSForms.TextBox

Private Sub TextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Enter キーで全文検索を実行
    If KeyCode = vbKeyReturn Then
        Call ApplyFullTextSearch
        ' ビープ音や改行入力を抑止
        KeyCode = 0
    End If
End Sub
