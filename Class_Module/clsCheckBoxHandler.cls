Option Explicit

' clsListBoxHandler クラスモジュール
Public WithEvents CheckBox As MSForms.CheckBox

Private Sub CheckBox_Change()
    Call startController
End Sub
