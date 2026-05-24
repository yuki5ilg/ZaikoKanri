Attribute VB_Name = "mSetup"
Option Explicit

' ============================================================
' mSetup - VBAプロジェクトに必要なモジュール/フォームを自動生成
'
' 【事前準備】
'   Excel オプション > トラスト センター > マクロの設定
'   「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」にチェック
'
' 【使い方】
'   1. このファイルをインポート
'   2. Alt+F8 → Setup を実行
'   3. Alt+F8 → Main を実行
' ============================================================

Private Const VBEXT_CT_STDMODULE As Long = 1
Private Const VBEXT_CT_MSFORM    As Long = 3

Public Sub Setup()
    On Error GoTo ErrHandler

    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    RemoveComponent vbProj, "mMain"
    RemoveComponent vbProj, "frmSettings"
    RemoveComponent vbProj, "frmInput"

    Dim mMain As Object
    Set mMain = vbProj.VBComponents.Add(VBEXT_CT_STDMODULE)
    mMain.Name = "mMain"
    mMain.CodeModule.AddFromString GetMMainCode()

    BuildFrmSettings vbProj
    BuildFrmInput vbProj

    MsgBox "セットアップ完了！" & vbCrLf & "Alt+F8 → Main を実行してください。", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "セットアップ失敗:" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Sub RemoveComponent(vbProj As Object, compName As String)
    Dim comp As Object
    On Error Resume Next
    Set comp = vbProj.VBComponents(compName)
    If Not comp Is Nothing Then vbProj.VBComponents.Remove comp
    On Error GoTo 0
End Sub

Private Function GetDesigner(comp As Object) As Object
    Dim i As Integer
    Application.VBE.MainWindow.Visible = True
    On Error Resume Next
    comp.Activate
    On Error GoTo 0
    For i = 1 To 50
        DoEvents
        If Not comp.Designer Is Nothing Then
            Set GetDesigner = comp.Designer
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 1000, , _
              comp.Name & " の Designer 取得失敗。" & vbCrLf & _
              "「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」が有効か確認してください。"
End Function

Private Function AddCtl(d As Object, progID As String, name As String, _
                        L As Single, T As Single, W As Single, H As Single, _
                        Optional caption As String = "") As Object
    Dim c As Object
    Set c = d.Controls.Add(progID, name)
    c.Left = L : c.Top = T : c.Width = W : c.Height = H
    On Error Resume Next
    If caption <> "" Then c.Caption = caption
    On Error GoTo 0
    Set AddCtl = c
End Function

Private Sub BuildFrmSettings(vbProj As Object)
    Dim comp As Object
    Set comp = vbProj.VBComponents.Add(VBEXT_CT_MSFORM)
    comp.Name = "frmSettings"
    comp.Properties("Caption") = "設定"
    comp.Properties("Width") = 300
    comp.Properties("Height") = 165

    Dim d As Object
    Set d = GetDesigner(comp)

    AddCtl d, "Forms.Label.1",         "lblSheetName", 12, 12,  100, 18, "対象シート名"
    AddCtl d, "Forms.TextBox.1",       "txtSheetName", 120, 10, 150, 20
    AddCtl d, "Forms.Label.1",         "lblFolder",    12, 42,  100, 18, "画像フォルダ"
    AddCtl d, "Forms.TextBox.1",       "txtFolder",    120, 40, 120, 20
    AddCtl d, "Forms.CommandButton.1", "btnBrowse",    245, 40, 30,  20, "..."
    AddCtl d, "Forms.CommandButton.1", "btnOK",        120, 95, 70,  25, "OK"
    AddCtl d, "Forms.CommandButton.1", "btnCancel",    200, 95, 70,  25, "キャンセル"

    comp.CodeModule.AddFromString GetFrmSettingsCode()
End Sub

Private Sub BuildFrmInput(vbProj As Object)
    Dim comp As Object
    Set comp = vbProj.VBComponents.Add(VBEXT_CT_MSFORM)
    comp.Name = "frmInput"
    comp.Properties("Caption") = "車両データ入力"
    comp.Properties("Width") = 600
    comp.Properties("Height") = 480

    Dim d As Object
    Set d = GetDesigner(comp)

    ' 左ペイン: 画像ナビ
    AddCtl d, "Forms.Label.1",         "lblProgress",  6,   6,   280, 18, "1 / 1"
    Dim img As Object
    Set img = AddCtl(d, "Forms.Image.1", "imgPreview",  6,   28,  280, 360)
    On Error Resume Next
    img.PictureSizeMode = 3
    img.BorderStyle = 1
    On Error GoTo 0
    AddCtl d, "Forms.Label.1",         "lblFileName",  6,   392, 280, 18, ""
    AddCtl d, "Forms.CommandButton.1", "btnPrev",      6,   414, 40,  22, "←"
    AddCtl d, "Forms.CommandButton.1", "btnNext",      50,  414, 40,  22, "→"

    ' 右ペイン: 入力フォーム（グレード・月は車名/年式に統合）
    Dim labels As Variant, names As Variant, i As Integer
    labels = Array("車名(グレード含む)", "年式/月", "色", "車台番号", "評価点", _
                   "走行距離(km)", "落札価格(円)", "消費税(円)", "自動車税(円)", _
                   "リサイクル(円)", "落札手数料(円)", "オークション会場", "出品番号")
    names  = Array("CarName", "Year", "Color", "Chassis", "Score", _
                   "Mileage", "Price", "Tax", "CarTax", "Recycle", "Fee", "Venue", "LotNum")

    Dim bL As Single, bT As Single, rH As Single
    bL = 300 : bT = 28 : rH = 26
    For i = 0 To UBound(labels)
        AddCtl d, "Forms.Label.1",   "lbl" & names(i), bL,      bT + i * rH, 95, 18, CStr(labels(i))
        AddCtl d, "Forms.TextBox.1", "txt" & names(i), bL + 98, bT + i * rH, 110, 20
    Next i

    Dim btnImp As Object
    Set btnImp = AddCtl(d, "Forms.CommandButton.1", "btnImport", 300, 420, 100, 28, "取込")
    On Error Resume Next
    btnImp.Default = True
    On Error GoTo 0
    AddCtl d, "Forms.CommandButton.1", "btnCancel", 410, 420, 100, 28, "中止"

    comp.CodeModule.AddFromString GetFrmInputCode()
End Sub

Private Function GetMMainCode() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' グローバル変数" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Public g_SheetName  As String" & vbCrLf
    s = s & "Public g_FolderPath As String" & vbCrLf
    s = s & "Public g_Cancelled  As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 1台分のデータ" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Public Type CarData" & vbCrLf
    s = s & "    CarNumber  As String   ' 管理番号（例: 27-001）" & vbCrLf
    s = s & "    YearMonth  As String   ' 年式/月（例: 2021/7）" & vbCrLf
    s = s & "    CarName    As String   ' 車名+グレード" & vbCrLf
    s = s & "    Score      As Variant  ' 評価点" & vbCrLf
    s = s & "    Price      As Variant  ' 落札価格" & vbCrLf
    s = s & "    Tax        As Variant  ' 消費税" & vbCrLf
    s = s & "    CarTax     As Variant  ' 自動車税" & vbCrLf
    s = s & "    Recycle    As Variant  ' リサイクル料" & vbCrLf
    s = s & "    AuctionFee As Variant  ' 落札手数料" & vbCrLf
    s = s & "    Venue      As String   ' オークション会場" & vbCrLf
    s = s & "    LotNumber  As String   ' 出品番号" & vbCrLf
    s = s & "    Color      As String   ' 色" & vbCrLf
    s = s & "    Chassis    As String   ' 車台番号" & vbCrLf
    s = s & "    Mileage    As Variant  ' 走行距離" & vbCrLf
    s = s & "End Type" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' エントリーポイント" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Sub Main()" & vbCrLf
    s = s & "    g_SheetName  = ""27期""" & vbCrLf
    s = s & "    g_FolderPath = """"" & vbCrLf
    s = s & "    g_Cancelled  = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    frmSettings.Show" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' A列を下から検索して連番の最大値+1を返す" & vbCrLf
    s = s & "' パターン: \d+-\d+ （例: 27-005）" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Function GetNextNumber(ws As Worksheet) As Long" & vbCrLf
    s = s & "    Dim i       As Long" & vbCrLf
    s = s & "    Dim maxNum  As Long" & vbCrLf
    s = s & "    Dim cellVal As String" & vbCrLf
    s = s & "    Dim parts() As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    maxNum = 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row To 1 Step -1" & vbCrLf
    s = s & "        On Error Resume Next" & vbCrLf
    s = s & "        cellVal = Trim(CStr(ws.Cells(i, 1).Value2))" & vbCrLf
    s = s & "        On Error GoTo 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        ' ハイフンが1つだけあるセルを対象" & vbCrLf
    s = s & "        If InStr(cellVal, ""-"") > 0 And _" & vbCrLf
    s = s & "           InStr(cellVal, ""-"") = InStrRev(cellVal, ""-"") Then" & vbCrLf
    s = s & "            parts = Split(cellVal, ""-"")" & vbCrLf
    s = s & "            If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then" & vbCrLf
    s = s & "                Dim num As Long" & vbCrLf
    s = s & "                num = CLng(parts(1))" & vbCrLf
    s = s & "                If num > maxNum Then maxNum = num" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    GetNextNumber = maxNum + 1" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 次の書き込み先行番号を返す" & vbCrLf
    s = s & "' Q列（17列）の「着」を末尾として次行を算出" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Function GetNextWriteRow(ws As Worksheet) As Long" & vbCrLf
    s = s & "    Dim lastRow As Long" & vbCrLf
    s = s & "    Dim qVal    As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lastRow = ws.Cells(ws.Rows.Count, 17).End(xlUp).Row" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If lastRow < 10 Then" & vbCrLf
    s = s & "        GetNextWriteRow = 10" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    qVal = Trim(CStr(ws.Cells(lastRow, 17).Value2))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Select Case qVal" & vbCrLf
    s = s & "        Case ""着""" & vbCrLf
    s = s & "            GetNextWriteRow = lastRow + 1   ' 行2の次 = 新しい行1" & vbCrLf
    s = s & "        Case ""予定""" & vbCrLf
    s = s & "            GetNextWriteRow = lastRow + 2   ' 行1の次の次 = 新しい行1" & vbCrLf
    s = s & "        Case Else" & vbCrLf
    s = s & "            GetNextWriteRow = lastRow + 1" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' シート名から期番号を取得（""27期"" → ""27""）" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Function GetPeriodNumber(ws As Worksheet) As String" & vbCrLf
    s = s & "    Dim s      As String" & vbCrLf
    s = s & "    Dim i      As Integer" & vbCrLf
    s = s & "    Dim result As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    s      = ws.Name" & vbCrLf
    s = s & "    result = """"" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 1 To Len(s)" & vbCrLf
    s = s & "        If IsNumeric(Mid(s, i, 1)) Then" & vbCrLf
    s = s & "            result = result & Mid(s, i, 1)" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            Exit For" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If result = """" Then result = ""27""" & vbCrLf
    s = s & "    GetPeriodNumber = result" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' シートへ2行書き込み（Value2のみ使用・書式変更なし）" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Sub WriteToSheet(ws As Worksheet, data As CarData)" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    r = GetNextWriteRow(ws)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' --- 行1 ---" & vbCrLf
    s = s & "    ws.Cells(r, 1).Value2  = data.CarNumber              ' A: 管理番号" & vbCrLf
    s = s & "    ws.Cells(r, 4).Value2  = data.YearMonth              ' D: 年式/月" & vbCrLf
    s = s & "    ws.Cells(r, 6).Value2  = data.CarName                ' F: 車名・グレード" & vbCrLf
    s = s & "    ws.Cells(r, 17).Value2 = ""予定""                      ' Q: 固定" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If data.Score      <> """" Then ws.Cells(r,  8).Value2 = data.Score" & vbCrLf
    s = s & "    If data.Price      <> """" Then ws.Cells(r,  9).Value2 = CLng(data.Price)" & vbCrLf
    s = s & "    If data.Tax        <> """" Then ws.Cells(r, 10).Value2 = CLng(data.Tax)" & vbCrLf
    s = s & "    If data.CarTax     <> """" Then ws.Cells(r, 11).Value2 = CLng(data.CarTax)" & vbCrLf
    s = s & "    If data.Recycle    <> """" Then ws.Cells(r, 12).Value2 = CLng(data.Recycle)" & vbCrLf
    s = s & "    If data.AuctionFee <> """" Then ws.Cells(r, 13).Value2 = CLng(data.AuctionFee)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' --- 行2 ---" & vbCrLf
    s = s & "    ws.Cells(r + 1, 2).Value2  = data.Venue              ' B: 会場" & vbCrLf
    s = s & "    ws.Cells(r + 1, 3).Value2  = data.LotNumber          ' C: 出品番号" & vbCrLf
    s = s & "    ws.Cells(r + 1, 4).Value2  = data.Color              ' D: 色" & vbCrLf
    s = s & "    ws.Cells(r + 1, 6).Value2  = data.Chassis            ' F: 車台番号" & vbCrLf
    s = s & "    ws.Cells(r + 1, 17).Value2 = ""着""                    ' Q: 固定" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If data.Mileage <> """" Then ws.Cells(r + 1, 7).Value2 = CLng(data.Mileage)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 画像を「処理済み」サブフォルダへ移動" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Sub MoveToProcessed(filePath As String)" & vbCrLf
    s = s & "    On Error GoTo ErrHandler" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim folder    As String" & vbCrLf
    s = s & "    Dim fileName  As String" & vbCrLf
    s = s & "    Dim processed As String" & vbCrLf
    s = s & "    Dim dest      As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    folder    = Left(filePath, InStrRev(filePath, ""\""))" & vbCrLf
    s = s & "    fileName  = Mid(filePath, InStrRev(filePath, ""\"") + 1)" & vbCrLf
    s = s & "    processed = folder & ""処理済み\""" & vbCrLf
    s = s & "    dest      = processed & fileName" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Dir(processed, vbDirectory) = """" Then MkDir processed" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' 同名ファイルがあれば連番を付ける" & vbCrLf
    s = s & "    If Dir(dest) <> """" Then" & vbCrLf
    s = s & "        Dim dot  As Integer" & vbCrLf
    s = s & "        Dim base As String" & vbCrLf
    s = s & "        Dim ext  As String" & vbCrLf
    s = s & "        Dim cnt  As Integer" & vbCrLf
    s = s & "        dot  = InStrRev(fileName, ""."")" & vbCrLf
    s = s & "        base = Left(fileName, dot - 1)" & vbCrLf
    s = s & "        ext  = Mid(fileName, dot)" & vbCrLf
    s = s & "        cnt  = 1" & vbCrLf
    s = s & "        Do While Dir(processed & base & ""_"" & cnt & ext) <> """"" & vbCrLf
    s = s & "            cnt = cnt + 1" & vbCrLf
    s = s & "        Loop" & vbCrLf
    s = s & "        dest = processed & base & ""_"" & cnt & ext" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    FileCopy filePath, dest" & vbCrLf
    s = s & "    Kill filePath" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "ErrHandler:" & vbCrLf
    s = s & "    MsgBox ""ファイル移動に失敗しました:"" & vbCrLf & filePath & vbCrLf & _" & vbCrLf
    s = s & "           Err.Description, vbExclamation" & vbCrLf
    s = s & "End Sub" & vbCrLf
    GetMMainCode = s
End Function

Private Function GetFrmSettingsCode() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    txtSheetName.Text = g_SheetName" & vbCrLf
    s = s & "    txtFolder.Text    = g_FolderPath" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnBrowse_Click()" & vbCrLf
    s = s & "    With Application.FileDialog(msoFileDialogFolderPicker)" & vbCrLf
    s = s & "        .Title = ""画像フォルダを選択してください""" & vbCrLf
    s = s & "        .AllowMultiSelect = False" & vbCrLf
    s = s & "        If .Show = True Then" & vbCrLf
    s = s & "            txtFolder.Text = .SelectedItems(1)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnOK_Click()" & vbCrLf
    s = s & "    ' --- バリデーション ---" & vbCrLf
    s = s & "    If Trim(txtSheetName.Text) = """" Then" & vbCrLf
    s = s & "        MsgBox ""シート名を入力してください。"", vbExclamation" & vbCrLf
    s = s & "        txtSheetName.SetFocus" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Trim(txtFolder.Text) = """" Then" & vbCrLf
    s = s & "        MsgBox ""画像フォルダを選択してください。"", vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Dir(Trim(txtFolder.Text), vbDirectory) = """" Then" & vbCrLf
    s = s & "        MsgBox ""フォルダが見つかりません:"" & vbCrLf & txtFolder.Text, vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Sheets(Trim(txtSheetName.Text))" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If ws Is Nothing Then" & vbCrLf
    s = s & "        MsgBox ""シート「"" & txtSheetName.Text & ""」が見つかりません。"", vbExclamation" & vbCrLf
    s = s & "        txtSheetName.SetFocus" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' --- 設定確定 ---" & vbCrLf
    s = s & "    g_SheetName  = Trim(txtSheetName.Text)" & vbCrLf
    s = s & "    g_FolderPath = Trim(txtFolder.Text)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' --- 画像ファイル収集 ---" & vbCrLf
    s = s & "    Dim exts(3) As String" & vbCrLf
    s = s & "    exts(0) = ""*.jpg"" : exts(1) = ""*.jpeg""" & vbCrLf
    s = s & "    exts(2) = ""*.png"" : exts(3) = ""*.bmp""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim fileList As New Collection" & vbCrLf
    s = s & "    Dim i As Integer, f As String" & vbCrLf
    s = s & "    For i = 0 To 3" & vbCrLf
    s = s & "        f = Dir(g_FolderPath & ""\"" & exts(i))" & vbCrLf
    s = s & "        Do While f <> """"" & vbCrLf
    s = s & "            fileList.Add g_FolderPath & ""\"" & f" & vbCrLf
    s = s & "            f = Dir()" & vbCrLf
    s = s & "        Loop" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If fileList.Count = 0 Then" & vbCrLf
    s = s & "        MsgBox ""画像ファイル（JPG/PNG/BMP）が見つかりませんでした。"", vbInformation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' --- Collection → 配列変換 ---" & vbCrLf
    s = s & "    Dim files() As String" & vbCrLf
    s = s & "    ReDim files(0 To fileList.Count - 1)" & vbCrLf
    s = s & "    Dim j As Integer" & vbCrLf
    s = s & "    j = 0" & vbCrLf
    s = s & "    Dim item As Variant" & vbCrLf
    s = s & "    For Each item In fileList" & vbCrLf
    s = s & "        files(j) = CStr(item)" & vbCrLf
    s = s & "        j = j + 1" & vbCrLf
    s = s & "    Next item" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' --- 1枚ずつ処理（1取込 = 1件追記） ---" & vbCrLf
    s = s & "    Dim processed As Integer" & vbCrLf
    s = s & "    processed = 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 0 To UBound(files)" & vbCrLf
    s = s & "        frmInput.SetupForm files, i" & vbCrLf
    s = s & "        frmInput.Show vbModal" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        If g_Cancelled Then" & vbCrLf
    s = s & "            MsgBox ""処理を中止しました。"" & vbCrLf & _" & vbCrLf
    s = s & "                   processed & "" / "" & fileList.Count & "" 件処理済み"", vbInformation" & vbCrLf
    s = s & "            Unload frmInput" & vbCrLf
    s = s & "            Unload Me" & vbCrLf
    s = s & "            Exit Sub" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        Dim data As CarData" & vbCrLf
    s = s & "        data           = frmInput.GetData()" & vbCrLf
    s = s & "        data.CarNumber = GetPeriodNumber(ws) & ""-"" & _" & vbCrLf
    s = s & "                         Format(GetNextNumber(ws), ""000"")" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        WriteToSheet ws, data" & vbCrLf
    s = s & "        MoveToProcessed files(i)" & vbCrLf
    s = s & "        processed = processed + 1" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    MsgBox processed & "" 件を処理しました。"", vbInformation" & vbCrLf
    s = s & "    Unload frmInput" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnCancel_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    GetFrmSettingsCode = s
End Function

Private Function GetFrmInputCode() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private m_Files() As String" & vbCrLf
    s = s & "Private m_Count   As Integer" & vbCrLf
    s = s & "Private m_Idx     As Integer" & vbCrLf
    s = s & "Private m_Main    As Integer   ' 取込対象のインデックス（ナビで変わらない）" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 初期化: 全画像配列 + 今回の取込対象インデックス" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Public Sub SetupForm(files() As String, mainIdx As Integer)" & vbCrLf
    s = s & "    m_Files = files" & vbCrLf
    s = s & "    m_Count = UBound(files) - LBound(files) + 1" & vbCrLf
    s = s & "    m_Idx   = mainIdx" & vbCrLf
    s = s & "    m_Main  = mainIdx" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    RefreshImage" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' === AI-OCR ここから（API課金のため無効化中）===" & vbCrLf
    s = s & "    ' Dim d As CarData" & vbCrLf
    s = s & "    ' d = CallClaudeOCR(m_Files(m_Main))" & vbCrLf
    s = s & "    ' txtCarName.Text  = d.CarName   ' 車名（グレード含む）" & vbCrLf
    s = s & "    ' txtYear.Text     = d.YearMonth ' 年式/月（例: 2021/7）" & vbCrLf
    s = s & "    ' txtColor.Text    = d.Color" & vbCrLf
    s = s & "    ' txtChassis.Text  = d.Chassis" & vbCrLf
    s = s & "    ' txtScore.Text    = d.Score" & vbCrLf
    s = s & "    ' txtMileage.Text  = d.Mileage" & vbCrLf
    s = s & "    ' txtPrice.Text    = d.Price" & vbCrLf
    s = s & "    ' txtTax.Text      = d.Tax" & vbCrLf
    s = s & "    ' txtCarTax.Text   = d.CarTax" & vbCrLf
    s = s & "    ' txtRecycle.Text  = d.Recycle" & vbCrLf
    s = s & "    ' txtFee.Text      = d.AuctionFee" & vbCrLf
    s = s & "    ' txtVenue.Text    = d.Venue" & vbCrLf
    s = s & "    ' txtLotNum.Text   = d.LotNumber" & vbCrLf
    s = s & "    ' === AI-OCR ここまで ===" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' --- デフォルト値（AIが有効になったら削除） ---" & vbCrLf
    s = s & "    txtCarName.Text  = ""スズキ クロスビー MZ""" & vbCrLf
    s = s & "    txtYear.Text     = ""2021/7""" & vbCrLf
    s = s & "    txtColor.Text    = ""ホワイトパール""" & vbCrLf
    s = s & "    txtChassis.Text  = ""MYN15S-100001""" & vbCrLf
    s = s & "    txtScore.Text    = ""4.5""" & vbCrLf
    s = s & "    txtMileage.Text  = ""35000""" & vbCrLf
    s = s & "    txtPrice.Text    = ""980000""" & vbCrLf
    s = s & "    txtTax.Text      = ""98000""" & vbCrLf
    s = s & "    txtCarTax.Text   = ""35400""" & vbCrLf
    s = s & "    txtRecycle.Text  = ""12000""" & vbCrLf
    s = s & "    txtFee.Text      = ""32000""" & vbCrLf
    s = s & "    txtVenue.Text    = ""USS大阪""" & vbCrLf
    s = s & "    txtLotNum.Text   = ""12345""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 画像表示更新（ナビ用。m_Idxが変わるだけ）" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Private Sub RefreshImage()" & vbCrLf
    s = s & "    If m_Count = 0 Then Exit Sub" & vbCrLf
    s = s & "    Dim prefix As String" & vbCrLf
    s = s & "    If m_Idx = m_Main Then" & vbCrLf
    s = s & "        prefix = ""★ ""   ' 取込対象" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        prefix = ""　""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    lblProgress.Caption = prefix & (m_Idx + 1) & "" / "" & m_Count" & vbCrLf
    s = s & "    lblFileName.Caption = Mid(m_Files(m_Idx), InStrRev(m_Files(m_Idx), ""\"") + 1)" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    imgPreview.Picture = LoadPicture(m_Files(m_Idx))" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnPrev_Click()" & vbCrLf
    s = s & "    If m_Idx > 0 Then m_Idx = m_Idx - 1 : RefreshImage" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnNext_Click()" & vbCrLf
    s = s & "    If m_Idx < m_Count - 1 Then m_Idx = m_Idx + 1 : RefreshImage" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 入力値を CarData で返す（グレード・月は統合済みのため不要）" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Public Function GetData() As CarData" & vbCrLf
    s = s & "    Dim data As CarData" & vbCrLf
    s = s & "    data.CarName     = Trim(txtCarName.Text)   ' 車名（グレード含む）" & vbCrLf
    s = s & "    data.YearMonth   = Trim(txtYear.Text)      ' 年式/月（例: 2021/7）" & vbCrLf
    s = s & "    data.Score       = Trim(txtScore.Text)" & vbCrLf
    s = s & "    data.Price       = Trim(txtPrice.Text)" & vbCrLf
    s = s & "    data.Tax         = Trim(txtTax.Text)" & vbCrLf
    s = s & "    data.CarTax      = Trim(txtCarTax.Text)" & vbCrLf
    s = s & "    data.Recycle     = Trim(txtRecycle.Text)" & vbCrLf
    s = s & "    data.AuctionFee  = Trim(txtFee.Text)" & vbCrLf
    s = s & "    data.Venue       = Trim(txtVenue.Text)" & vbCrLf
    s = s & "    data.LotNumber   = Trim(txtLotNum.Text)" & vbCrLf
    s = s & "    data.Color       = Trim(txtColor.Text)" & vbCrLf
    s = s & "    data.Chassis     = Trim(txtChassis.Text)" & vbCrLf
    s = s & "    data.Mileage     = Trim(txtMileage.Text)" & vbCrLf
    s = s & "    GetData = data" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnImport_Click()" & vbCrLf
    s = s & "    If Trim(txtCarName.Text) = """" Then" & vbCrLf
    s = s & "        MsgBox ""車名を入力してください。"", vbExclamation" & vbCrLf
    s = s & "        txtCarName.SetFocus" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnCancel_Click()" & vbCrLf
    s = s & "    If MsgBox(""処理を中止しますか？"", vbQuestion + vbYesNo) = vbYes Then" & vbCrLf
    s = s & "        g_Cancelled = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    GetFrmInputCode = s
End Function

