Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
#End If

Dim cmdHandlers_CoB() As clsCommandButtonHandler
Dim cmdHandlers_LB() As clsListBoxHandler
Dim cmdHandler_ChB As clsCheckBoxHandler ' コンボボックスと区別するためCBではなくChB. あと1個だけなので配列ではなく単体
Dim cmdHandler_FullSearchTxt As clsTextBoxHandler
Dim cmdHandler_FullSearch As clsCommandButtonHandler

Sub startController()
    Dim i As Integer, k As Integer
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ActiveSheet
    Set tbl = UniqueTable(ws)

    Dim uf As UserForm
    Set uf = MyController_UF

    ' コントローラ起動時は全文検索ボックスをクリアしておく
    On Error Resume Next
    uf.Controls("TextBoxFullSearch").Value = ""
    On Error GoTo 0

    ' ===== 前提チェック =====
    Dim numOfTagsCol As Integer
    Dim tagColNames As Collection
    Dim tagsMissing As Boolean
    Set tagColNames = New Collection

    ' Tags_列の連番チェック
    i = 1
    tagsMissing = False
    Do While True
        On Error Resume Next
        Dim colname As String
        colname = "Tags_" & i
        Dim col As ListColumn
        Set col = tbl.ListColumns(colname)
        If Err.Number <> 0 Then
            ' 連番が途切れたら終了
            Exit Do
        End If
        tagColNames.Add colname
        i = i + 1
        On Error GoTo 0
    Loop
    numOfTagsCol = tagColNames.Count
    ' 1からnumOfTagsColまでスキップなく存在するか再確認
    For k = 1 To numOfTagsCol
        On Error Resume Next
        Dim col2 As ListColumn
        Set col2 = tbl.ListColumns("Tags_" & k)
        If Err.Number <> 0 Then
            tagsMissing = True
            Exit For
        End If
        On Error GoTo 0
    Next k
    If tagsMissing Or numOfTagsCol = 0 Then
        MsgBox "Tags_列が1から連番でスキップなく存在しません。テーブル構造を確認してください。" & vbCrLf & "終了します。"
        Exit Sub
    End If

    ' FF_列の調整（全削除→必要分だけ追加）
    Application.ScreenUpdating = False
    ' まずFF_で始まる列をすべて削除
    For i = tbl.ListColumns.Count To 1 Step -1
        Dim col3 As ListColumn
        Set col3 = tbl.ListColumns(i)
        Dim colname3 As String
        colname3 = col3.Name
        If Left(colname3, 3) = "FF_" Then
            Dim suffix As String ' ← ループ内でのみ有効
            suffix = Mid(colname3, 4)
            col3.Delete
        End If
    Next i
    ' FF_1～FF_Kを新規追加
    For i = 1 To numOfTagsCol
        tbl.ListColumns.Add.Name = "FF_" & i
    Next i
    Application.ScreenUpdating = True

    If tbl Is Nothing Then
        MsgBox "シート上のテーブルが1つでないとこのツールは動かせません．" & vbCrLf & "終了します．"
        Exit Sub
    End If

    ' ===== テーブルのフィルター全解除 =====
    Call ClearAllFilters(tbl)
    ' ===== フィルターフラグ列・A_FF列を非表示にする =====
    Dim col4 As ListColumn
    For Each col4 In tbl.ListColumns
        Dim colname4 As String
        colname4 = col4.Name
        ' 列名が「FF_」で始まる、またはA_FF列なら非表示
        If Left(colname4, 3) = "FF_" Or colname4 = "A_FF" Then
            If Not col4.Range.EntireColumn.Hidden Then
                col4.Range.EntireColumn.Hidden = True
            End If
        End If
    Next col4

    ' ===== MyController_UF にコントロールを入れる =====
    ' ===== 既にリストボックスがある場合はスキップする =====
    ' ※リストボックスのコレクションを取得する関数があって便利なのでこれを利用している．判定条件はなんでもよい

    Dim listBoxes As Collection
    Set listBoxes = getListBoxCollection(uf)

    If listBoxes.Count = 0 Then
        ' ===== チェックボックス→テキストボックス→全文検索ボタンの順で生成 =====
        Dim chkBox As MSForms.CheckBox
        Set chkBox = uf.Controls.Add("Forms.CheckBox.1", "CheckBox1", True)
        chkBox.Value = True
        Set cmdHandler_ChB = New clsCheckBoxHandler
        Set cmdHandler_ChB.CheckBox = chkBox
        Set chkBox = Nothing

        Dim txtFullSearch As MSForms.TextBox
        Set txtFullSearch = uf.Controls.Add("Forms.TextBox.1", "TextBoxFullSearch", True)
        ' 位置・サイズは後で一括設定
    ' TextBox の Enter キーを捕まえるハンドラを割り当て
    Set cmdHandler_FullSearchTxt = New clsTextBoxHandler
    Set cmdHandler_FullSearchTxt.TextBox = txtFullSearch
        Dim btnFullSearch As MSForms.CommandButton
        Set btnFullSearch = uf.Controls.Add("Forms.CommandButton.1", "ButtonFullSearch", True)
        btnFullSearch.Caption = "全文検索"
        ' 位置・サイズは後で一括設定
    ' ButtonFullSearch に既存の clsCommandButtonHandler を割り当てる（新規クラスモジュールは不要）
    Set cmdHandler_FullSearch = New clsCommandButtonHandler
    Set cmdHandler_FullSearch.CommandButton = btnFullSearch
    cmdHandler_FullSearch.ButtonIndex = -1 ' 特殊値: 全文検索

        ' ===== リストボックスを生成 =====
        ReDim cmdHandlers_LB(1 To numOfTagsCol)
        Dim iLB As Integer
        For iLB = 1 To numOfTagsCol
            Dim lstBox As MSForms.ListBox
            Set lstBox = uf.Controls.Add("Forms.ListBox.1", "ListBox" & iLB, True)
            ' リストボックスのプロパティを設定
            With lstBox
                .Left = 50 + (iLB - 1) * 110
                .Top = 100
                .Width = 100
                .Height = 300
            End With
            ' クラスモジュールのインスタンスを作成し、ListBoxを設定
            Set cmdHandlers_LB(iLB) = New clsListBoxHandler
            Set cmdHandlers_LB(iLB).ListBox = lstBox
            cmdHandlers_LB(iLB).ListBoxIndex = iLB
        Next iLB

        ' ===== コマンドボタンを生成 =====
        ReDim cmdHandlers_CoB(1 To numOfTagsCol)
        Dim iCB As Integer
        For iCB = 1 To numOfTagsCol
            Dim cmdButton As MSForms.CommandButton
            Set cmdButton = uf.Controls.Add("Forms.CommandButton.1", "CommandButton" & iCB, True)
            ' コマンドボタンのプロパティを設定
            With cmdButton
                .Caption = "Reset " & iCB
                .Left = 50
                .Top = 50 + (iCB - 1) * 40
                .Width = 100
                .Height = 30
            End With
            ' クラスモジュールのインスタンスを作成し、CommandButtonを設定
            Set cmdHandlers_CoB(iCB) = New clsCommandButtonHandler
            Set cmdHandlers_CoB(iCB).CommandButton = cmdButton
            cmdHandlers_CoB(iCB).ButtonIndex = iCB
        Next iCB
    End If

    ' ===== ユーザーフォームとコントロールの整形・配置 =====
    ' カーソルをA1セルに持ってくる
    ' Range("A1").Activate

    ' 最下行の番号を取得
    Dim lastVisibleRow As Integer
    lastVisibleRow = ActiveWindow.visibleRange.Rows.Count


    ' --- UserForm配置仕様に基づく幅計算 ---
    Dim num_TopColumns As Integer
    num_TopColumns = getNumTopColumns(tbl)
    Dim ufWidth As Double, ufHeight As Double
    Dim listBoxWidths() As Double
    Dim iCol As Integer

    ufHeight = ws.Range("A" & lastVisibleRow).Top - ws.Range("A2").Top

    If num_TopColumns = 0 Then
        MsgBox "テーブルの左端がA列の場合は使用できません。シート左側に空き列を設けてください。"
        Exit Sub
    ElseIf num_TopColumns >= numOfTagsCol Then
        ' A列からnum_TopColumns列の右端まで
        ufWidth = 0
        ReDim listBoxWidths(1 To numOfTagsCol)
        Dim colWidths() As Double
        colWidths = getColumnWidths(ws, 1, num_TopColumns)
        For iCol = 1 To numOfTagsCol
            listBoxWidths(iCol) = colWidths(iCol)
            ufWidth = ufWidth + colWidths(iCol)
        Next iCol
    Else
        ' A列の幅でUserFormを作り、等分
        ufWidth = ws.Columns(1).Width
        ReDim listBoxWidths(1 To numOfTagsCol)
        For iCol = 1 To numOfTagsCol
            listBoxWidths(iCol) = ufWidth / numOfTagsCol
        Next iCol
    End If

    ' パラメータ ボタン・リストボックス・チェックボックス
    Dim offsetX As Integer: offsetX = 10
    Dim offsetY As Integer: offsetY = 10
    Dim fSize As Integer: fSize = Application.StandardFontSize
    Dim leftPos As Double
    Dim checkBoxHeight As Integer: checkBoxHeight = 24
    ' --- 最上部パーツ配置 ---
    Dim topMargin As Integer: topMargin = 10
    Dim partHeight As Integer: partHeight = 24
    Dim partGap As Integer: partGap = 8
    Dim checkBoxWidth As Integer: checkBoxWidth = 100 ' チェックボックス+ラベルの幅
    Dim btnFullSearchWidth As Integer: btnFullSearchWidth = 80
    Dim partGap1 As Integer: partGap1 = 8 ' テキストボックスとボタンの間
    Dim partGap2 As Integer: partGap2 = 8 ' ボタンとチェックボックスの間
    Dim txtFullSearchLeft As Double, txtFullSearchWidth As Double
    Dim btnFullSearchLeft As Double, checkBoxLeft As Double
    txtFullSearchLeft = offsetX
    txtFullSearchWidth = ufWidth - btnFullSearchWidth - checkBoxWidth
    If txtFullSearchWidth < 120 Then txtFullSearchWidth = 120
    btnFullSearchLeft = txtFullSearchLeft + txtFullSearchWidth + partGap1
    checkBoxLeft = btnFullSearchLeft + btnFullSearchWidth + partGap2
    On Error Resume Next
    ' テキストボックス
    With uf.Controls("TextBoxFullSearch")
        .Top = topMargin
        .Left = txtFullSearchLeft
        .Width = txtFullSearchWidth
        .Height = partHeight
        .FontSize = fSize
        .Font.Name = Application.StandardFont
    End With
    ' コマンドボタン
    With uf.Controls("ButtonFullSearch")
        .Top = topMargin
        .Left = btnFullSearchLeft
        .Width = btnFullSearchWidth
        .Height = partHeight
        .FontSize = fSize
        .Font.Name = Application.StandardFont
    End With
    ' チェックボックス
    With uf.Controls("CheckBox1")
        .Top = topMargin
        .Left = checkBoxLeft
        .Width = checkBoxWidth
        .Height = partHeight
        .Caption = "絞り込み"
        .FontSize = fSize
        .Font.Name = Application.StandardFont
    End With
    On Error GoTo 0
    ' --- Resetボタン配置 ---
    Dim commandButtonHeight As Integer: commandButtonHeight = 24
    Dim resetTop As Integer: resetTop = topMargin + partHeight + partGap
    ' --- リストボックス配置 ---
    Dim checkBottom As Integer: checkBottom = resetTop + commandButtonHeight + partGap
    Dim listBoxHeight As Integer
    ' UserForm下端に余白を持たせて見栄えを良くする（例: 30pt余白）
    Dim bottomMargin As Integer: bottomMargin = 30
    listBoxHeight = ufHeight - checkBottom - bottomMargin


    ' === ウィンドウ枠の固定一時解除対応 ===
    Dim wasFreezePanes As Boolean: wasFreezePanes = False
    Dim splitRow As Long, splitCol As Long
    If ActiveWindow.FreezePanes = True Then
        wasFreezePanes = True
        splitRow = ActiveWindow.SplitRow
        splitCol = ActiveWindow.SplitColumn
        ActiveWindow.FreezePanes = False
    End If


    ' === UserForm表示前にA1セルを選択し、ウィンドウの左端がA列になるようにする ===
    ws.Activate
    ws.Range("A1").Select

    ' === ユーザーフォーム ===
    ' DPIスケーリング対応: OSの拡大率を考慮してA3セル上付近に表示
    Dim dpi As Long
    Dim dpiScale As Double
    dpi = GetSystemDPI()
    dpiScale = dpi / 96

    ' Excelウィンドウ左上のスクリーン座標（ピクセル）
    Dim winLeftPx As Long, winTopPx As Long
    winLeftPx = ActiveWindow.PointsToScreenPixelsX(0)
    winTopPx = ActiveWindow.PointsToScreenPixelsY(0)

    ' A3セルのワークシート上座標（ポイント）
    Dim a3LeftPt As Double, a3TopPt As Double
    a3LeftPt = ws.Range("A3").Left
    a3TopPt = ws.Range("A3").Top

    ' ピクセル→ポイント変換（DPI倍率で補正）
    Dim winLeftPt As Double, winTopPt As Double
    winLeftPt = winLeftPx * 72 / dpi
    winTopPt = winTopPx * 72 / dpi

    With MyController_UF
        .StartUpPosition = 0
        .Top = winTopPt + a3TopPt
        .Left = winLeftPt + a3LeftPt
        .Height = ufHeight
        .Width = ufWidth
    End With


    ' === リストボックス ===
    Set listBoxes = getListBoxCollection(uf)
    Dim iLB2 As Integer
    Dim nLB As Integer: nLB = listBoxes.Count
    leftPos = 0
    Dim shrinkW As Double: shrinkW = 4 ' 4ptだけ幅を縮小 見栄え調整のため
    For iLB2 = 1 To nLB
        With listBoxes(iLB2)
            .Top = checkBottom
            .Left = leftPos
            .Width = listBoxWidths(iLB2) - shrinkW
            .Height = listBoxHeight
            .FontSize = fSize
            .Font.Name = Application.StandardFont
        End With
        leftPos = leftPos + (listBoxWidths(iLB2) - shrinkW)
    Next iLB2

    ' === コマンドボタン ===
    Dim iCmd As Integer
    For iCmd = 1 To listBoxes.Count
        On Error Resume Next
        Dim btn As Object
        Set btn = uf.Controls("CommandButton" & iCmd)
        If Not btn Is Nothing Then
            With btn
                .Top = resetTop
                .Left = listBoxes(iCmd).Left
                .Width = listBoxes(iCmd).Width
                .Height = commandButtonHeight
                .FontSize = fSize
                .Font.Name = Application.StandardFont
            End With
        End If
        On Error GoTo 0
    Next iCmd

    ' （重複配置ブロック削除済み）

    ' ===== リストボックスに値を入れる =====
    ' 各リストボックスのマルチセレクトを有効にする
    ' リストボックスをクリア（ユーザフォームが出ている状態で再度実行されたときのため）して，再度値を取得
    Dim iLB3 As Integer
    For iLB3 = 1 To listBoxes.Count
        listBoxes(iLB3).Clear
        listBoxes(iLB3).MultiSelect = fmMultiSelectMulti
        Call setTagsIntoListBox(tbl, listBoxes, iLB3, True)
    Next iLB3


    ' ユーザーフォームを表示
    MyController_UF.Show vbModeless

    ' === ウィンドウ枠の固定を元に戻す ===
    If wasFreezePanes Then
        ' wsがNothingにならないよう、UserForm表示後にNothing化する
        ws.Cells(splitRow + 1, splitCol + 1).Select
        ActiveWindow.FreezePanes = True
    End If

    Set ws = Nothing
    Set tbl = Nothing
    Set listBoxes = Nothing
    Set uf = Nothing

End Sub



Function NumOfColumnsStartWith(tbl As ListObject, searchText As String) As Integer
    Dim num As Integer
    Dim col As ListColumn
    num = 0
    
    For Each col In tbl.ListColumns
        
        If Left(col.Name, Len(searchText)) = searchText Then
            num = num + 1
        End If
    Next col
    
    NumOfColumnsStartWith = num
    
End Function

Function UniqueTable(ws As Worksheet) As ListObject
    ' テーブルがユニークでない場合は Nothing を返し，呼び出し元のプロシージャでExit Subする
    If ws.ListObjects.Count = 1 Then
        Set UniqueTable = ws.ListObjects(1)
    Else
        Set UniqueTable = Nothing
    End If
End Function

'--- UserForm配置仕様用 補助関数 ---
Function getNumTopColumns(tbl As ListObject) As Integer
    getNumTopColumns = tbl.Range.Column - 1
End Function

' ws: Worksheet, startCol: 1, endCol: N
Function getColumnWidths(ws As Worksheet, startCol As Integer, endCol As Integer) As Double()
    Dim arr() As Double
    Dim i As Integer
    ReDim arr(1 To endCol - startCol + 1)
    For i = startCol To endCol
        arr(i - startCol + 1) = ws.Columns(i).Width
    Next i
    getColumnWidths = arr
End Function

Function getListBoxCollection(uf As UserForm) As Collection
    Dim ret As New Collection
    Dim ctrl As Control
    Dim i As Integer
    Dim num As Integer
    num = 0
    For Each ctrl In uf.Controls
        If TypeName(ctrl) = "ListBox" Then
            num = num + 1
        End If
    Next ctrl
    
    For i = 1 To num
        On Error Resume Next
        Dim ctrlObj As Object
        Set ctrlObj = uf.Controls("ListBox" & i)
        If Not ctrlObj Is Nothing Then
            ret.Add ctrlObj
        End If
        On Error GoTo 0
    Next i
    
    Set getListBoxCollection = ret
End Function
Function getCommandButtonCollection(uf As UserForm) As Collection
    Dim ret As New Collection
    Dim ctrl As Control
    For Each ctrl In uf.Controls
        If TypeName(ctrl) = "CommandButton" Then
            If Left(ctrl.Name, 12) = "CommandButton" Then
                ret.Add ctrl
            End If
        End If
    Next ctrl
    Set getCommandButtonCollection = ret
End Function


Sub setTagsIntoListBox(tbl As ListObject, listBoxes As Collection, num As Integer, flagOfVisibleOnly As Boolean)

    Dim rng As Range
    Set rng = tbl.ListColumns("Tags_" & num).DataBodyRange
    ' フラグがTrueなら対象のTags列の表示セルだけ取得する
    If flagOfVisibleOnly Then
        On Error Resume Next
        Set rng = rng.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
    End If

    Dim allTags As String
    allTags = ""

    ' 取得したセルを","で結合して1つのStringにする
    If rng Is Nothing Then
        allTags = ""
    Else
        Dim cell As Range
        For Each cell In rng
            allTags = allTags & "," & cell.Value
        Next cell
    End If

    ' カンマで区切って配列に入れる
    Dim tagsArray() As String
    tagsArray = Split(allTags, ",")

    ' カンマの前後の余分なスペースをトリム
    Dim i As Integer
    For i = LBound(tagsArray) To UBound(tagsArray)
        tagsArray(i) = Trim(tagsArray(i))
    Next i

    ' 空文字と重複を削除
    Dim uniqueTags As New Collection
    On Error Resume Next
    Dim tag As String
    For i = LBound(tagsArray) To UBound(tagsArray)
        tag = Trim(tagsArray(i))
        If tag <> "" Then
            uniqueTags.Add tag, CStr(tag) ' CStrを使ってキーとして追加
        End If
    Next i
    On Error GoTo 0

    ' Collectionから配列に変換
    ' タグが0個ならListBoxに何も入れなくていいのでExit Subする
    If uniqueTags.Count = 0 Then
        Exit Sub
    Else
        Dim resultArray() As String
        ReDim resultArray(1 To uniqueTags.Count)
        For i = 1 To uniqueTags.Count
            resultArray(i) = uniqueTags(i)
        Next i

        ' 配列resultArrayをバブルソートで整列する
        Dim j As Integer
        Dim temp As String
        For i = LBound(resultArray) To UBound(resultArray) - 1
            For j = i + 1 To UBound(resultArray)
                If resultArray(i) > resultArray(j) Then
                    temp = resultArray(i)
                    resultArray(i) = resultArray(j)
                    resultArray(j) = temp
                End If
            Next j
        Next i

        ' 対象のリストボックスに値を入れる
        If num >= 1 And num <= listBoxes.Count Then
            For i = 1 To UBound(resultArray)
                listBoxes(num).AddItem resultArray(i)
            Next i
        End If
    End If
End Sub

Public Sub ApplyTagFiltersAndUpdateListBoxes(num As Integer)
    Application.ScreenUpdating = False ' 画面描画を停止
    Dim i As Integer, j As Integer
    Dim selectedTagsGroup As Collection
    Set selectedTagsGroup = New Collection
    
    Dim subCollection As Collection
    
    
    Dim uf As UserForm
    Set uf = MyController_UF
    
    Dim row As ListRow
    Dim containsTag As Boolean
    
    Dim listBoxes As Collection
    Set listBoxes = getListBoxCollection(uf)
    
    Dim tbl As ListObject
    Set tbl = UniqueTable(ActiveSheet)
    
    Dim tag As Variant
    
    ' Copilotが追加した変数
    Dim tagsCollection As Collection
'    Dim filterCriteria(1 To 4) As String
    ' 既存のFF_フィルターのみをクリア（A_FFや他の列のフィルターは保持する）
    Dim fb As Integer
    For fb = 1 To listBoxes.Count
        Call removeFilter(fb, tbl)
    Next fb

    For i = 1 To listBoxes.Count
        Set tagsCollection = New Collection
        
        ' リストボックスの選択項目をコレクションに追加
        For j = 0 To listBoxes(i).ListCount - 1
            If listBoxes(i).Selected(j) Then
                tagsCollection.Add listBoxes(i).List(j)
            End If
        Next j
        
        ' 各行のTags_n列を確認し、FilterFlag_n列にフラグを設定
        For Each row In tbl.ListRows
            containsTag = False
            For Each tag In tagsCollection
                If InStr(1, row.Range(tbl.ListColumns("Tags_" & i).index).Value, tag, vbTextCompare) > 0 Then
                    containsTag = True
                    Exit For
                End If
            Next tag
            If containsTag Then
                row.Range(tbl.ListColumns("FF_" & i).index).Value = 1
            Else
                row.Range(tbl.ListColumns("FF_" & i).index).Value = 0
            End If
        Next row
        
        ' タグ選択が一つもないならフィルターをかけない
        If tagsCollection.Count <> 0 Then
            With tbl.Range
                .AutoFilter Field:=tbl.ListColumns("FF_" & i).index, Criteria1:=1
            End With
        End If
        
    Next i
    
    Set selectedTagsGroup = Nothing
    Application.ScreenUpdating = True ' 画面描画を開始
    
    ' ===== 絞り込みが無効ならここで終わり =====
    If uf.CheckBox1.Value = False Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False ' 画面描画を停止
    ' num番目以降の列をフィルター解除
    For i = num + 1 To listBoxes.Count
        Call removeFilter(i, tbl)
    Next i
    ' num番目より後のリストボックスを再設定（絞り込まれたタグだけを再設定）
    For i = num + 1 To listBoxes.Count
        listBoxes(i).Clear
        Call setTagsIntoListBox(tbl, listBoxes, i, True)
    Next i
    
    

    
    Application.ScreenUpdating = True ' 画面描画を開始
    
    

End Sub

Sub removeFilter(num As Integer, tbl As ListObject)
    If tbl.AutoFilter.FilterMode Then
    tbl.Range.AutoFilter Field:=tbl.ListColumns("FF_" & num).index
    End If
    
End Sub

Sub ResetTagFilterAndListBox(num As Integer)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ActiveSheet
    Set tbl = UniqueTable(ws)
    Dim uf As UserForm
    Set uf = MyController_UF
    Dim i As Integer
    
    Dim listBoxes As Collection
    Set listBoxes = getListBoxCollection(uf)
    
    If uf.Controls("CheckBox1").Value = True Then
        If tbl.AutoFilter.FilterMode Then
            For i = num To listBoxes.Count
                tbl.Range.AutoFilter Field:=tbl.ListColumns("FF_" & i).index
            Next i
            For i = num To listBoxes.Count
                listBoxes(i).Clear
                Call setTagsIntoListBox(tbl, listBoxes, i, True)
            Next i
        End If
    Else
        If tbl.AutoFilter.FilterMode Then
            tbl.Range.AutoFilter Field:=tbl.ListColumns("FF_" & num).index
        End If
        
        listBoxes(num).Clear
        Call setTagsIntoListBox(tbl, listBoxes, num, False)

    End If
    

    
End Sub

Sub ClearAllFilters(tbl As ListObject)
    If tbl.AutoFilter.FilterMode Then
        tbl.AutoFilter.ShowAllData
    End If
End Sub

Sub ClearAllFiltersOfUniqueTableOnActiveSheet()
    Dim tbl As ListObject
    Set tbl = UniqueTable(ActiveSheet)
        If tbl.AutoFilter.FilterMode Then
            tbl.AutoFilter.ShowAllData
        End If

End Sub


' ユーザーフォームがロードされているかどうかを返すヘルパー
Public Function IsUserFormLoaded(formName As String) As Boolean
    Dim uf As Object
    For Each uf In VBA.UserForms
        If uf.Name = formName Then
            IsUserFormLoaded = True
            Exit Function
        End If
    Next uf
    IsUserFormLoaded = False
End Function


' ワークシート上の「フィルター全解除ボタン」から呼び出すための統合ハンドラ
' MyController_UF が開いている場合はコントローラを再起動して UI を更新し、
' 開いていない場合は従来どおりテーブルのフィルターを全解除する。
Public Sub ClearAllFiltersButtonHandler()
    On Error Resume Next
    If IsUserFormLoaded("MyController_UF") Then
        ' テーブルのフィルターを解除
        Call ClearAllFiltersOfUniqueTableOnActiveSheet

        ' 再描画のため startController を呼ぶ（ただし startController は
        ' 既に存在するリストボックスがある場合は新規作成をスキップする）
        Call startController
    Else
        Call ClearAllFiltersOfUniqueTableOnActiveSheet
    End If
    On Error GoTo 0
End Sub


Public Sub ApplyFullTextSearch()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim uf As UserForm
    Dim searchText As String
    Dim i As Integer, j As Integer
    Dim row As ListRow
    Dim matches As Boolean

    Set uf = MyController_UF
    Set ws = ActiveSheet
    Set tbl = UniqueTable(ws)
    If tbl Is Nothing Then Exit Sub

    ' --- A_FF列があるか確認、なければ追加 ---
    On Error Resume Next
    Dim aFFCol As ListColumn
    Set aFFCol = tbl.ListColumns("A_FF")
    If Err.Number <> 0 Or aFFCol Is Nothing Then
        Err.Clear
        tbl.ListColumns.Add.Name = "A_FF"
        Set aFFCol = tbl.ListColumns("A_FF")
    End If
    On Error GoTo 0

    ' 非表示にする
    On Error Resume Next
    aFFCol.Range.EntireColumn.Hidden = True
    On Error GoTo 0

    searchText = Trim(CStr(uf.Controls("TextBoxFullSearch").Value))

    Application.ScreenUpdating = False

    ' まず既存のA_FFフィルターを解除
    If tbl.AutoFilter.FilterMode Then
        tbl.Range.AutoFilter Field:=aFFCol.Index
    End If

    ' 空文字ならA_FFを全て1にして終了（フィルターはかけない）
    If searchText = "" Then
        For Each row In tbl.ListRows
            row.Range(tbl.ListColumns("A_FF").Index).Value = 1
        Next row
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' 各行を走査して検索文字列が含まれるかを判定
    ' ただし、タグフィルタ (FF_*) が既にかかっている場合は可視行のみを対象とする
    Dim listBoxes As Collection
    Set listBoxes = getListBoxCollection(uf)

    Dim anyFFFiltered As Boolean: anyFFFiltered = False
    If tbl.AutoFilter.FilterMode Then
        Dim f As Integer
        For f = 1 To listBoxes.Count
            On Error Resume Next
            Dim fieldPos As Integer
            fieldPos = tbl.ListColumns("FF_" & f).Index - tbl.Range.Column + 1
            If fieldPos >= 1 Then
                If tbl.AutoFilter.Filters(fieldPos).On Then
                    anyFFFiltered = True
                    Exit For
                End If
            End If
            On Error GoTo 0
        Next f
    End If

    For Each row In tbl.ListRows
        ' 可視行のみ処理する（タグフィルタ適用時）
        If anyFFFiltered Then
            If row.Range.Rows(1).EntireRow.Hidden Then
                ' フィルタで非表示の行はスキップ
                GoTo ContinueRow
            End If
        End If
        matches = False
        For j = 1 To tbl.ListColumns.Count
            Dim colName As String
            colName = tbl.ListColumns(j).Name
            ' FF_ と A_FF 列は検索対象外にする
            If Left(colName, 3) <> "FF_" And colName <> "A_FF" Then
                Dim cellValue As String
                cellValue = CStr(row.Range(tbl.ListColumns(j).Index).Value)
                If InStr(1, cellValue, searchText, vbTextCompare) > 0 Then
                    matches = True
                    Exit For
                End If
            End If
        Next j

        If matches Then
            row.Range(tbl.ListColumns("A_FF").Index).Value = 1
        Else
            row.Range(tbl.ListColumns("A_FF").Index).Value = 0
        End If
ContinueRow:
    Next row

    ' A_FF列でフィルターをかける
    With tbl.Range
        .AutoFilter Field:=tbl.ListColumns("A_FF").Index, Criteria1:=1
    End With

    ' フィルターに合わせてリストボックスの項目を再構築
    Set listBoxes = getListBoxCollection(uf)
    For i = 1 To listBoxes.Count
        listBoxes(i).Clear
        Call setTagsIntoListBox(tbl, listBoxes, i, True)
    Next i

    Application.ScreenUpdating = True
End Sub

    ' --- DPI取得用ヘルパー関数 ---
    Private Function GetSystemDPI() As Long
    #If VBA7 Then
        Dim hDC As LongPtr
        hDC = GetDC(0)
        If hDC <> 0 Then
            GetSystemDPI = GetDeviceCaps(hDC, 88) ' LOGPIXELSX
            ReleaseDC 0, hDC
        Else
            GetSystemDPI = 96
        End If
    #Else
        Dim hDC As Long
        hDC = GetDC(0)
        If hDC <> 0 Then
            GetSystemDPI = GetDeviceCaps(hDC, 88)
            ReleaseDC 0, hDC
        Else
            GetSystemDPI = 96
        End If
    #End If
    End Function
