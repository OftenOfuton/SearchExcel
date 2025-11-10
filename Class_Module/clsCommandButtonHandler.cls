Option Explicit

' clsCommandButtonHandler クラスモジュール
Public WithEvents CommandButton As MSForms.CommandButton
Public ButtonIndex As Integer

Private Sub CommandButton_Click()
    ' 全文検索ボタンかどうか判定
    On Error Resume Next
    Dim btnName As String
    btnName = CommandButton.Name
    On Error GoTo 0

    If btnName = "ButtonFullSearch" Or ButtonIndex = -1 Then
        Call ApplyFullTextSearch
    Else
        Call ResetTagFilterAndListBox(ButtonIndex)
    End If
End Sub
