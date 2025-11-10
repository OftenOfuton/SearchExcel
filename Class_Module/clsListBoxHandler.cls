Option Explicit

' clsListBoxHandler クラスモジュール
Public WithEvents ListBox As MSForms.ListBox
Public ListBoxIndex As Integer

Private Sub ListBox_Change()
    Call ApplyTagFiltersAndUpdateListBoxes(ListBoxIndex)
End Sub
