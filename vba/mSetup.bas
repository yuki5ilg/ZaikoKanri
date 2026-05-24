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
'   1. mSetup.bas だけインポート
'   2. Alt+F8 → Setup 実行
'   3. Alt+F8 → Main 実行
' ============================================================

Private Const VBEXT_CT_STDMODULE As Long = 1
Private Const VBEXT_CT_MSFORM    As Long = 3

Public Sub Setup()
    Dim step As String
    On Error GoTo ErrHandler

    step = "0: VBE初期化"
    Application.VBE.MainWindow.Visible = True
    DoEvents

    step = "1/5: VBProject取得"
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    step = "2/5: 既存コンポーネント削除"
    RemoveComponent vbProj, "mMain"
    RemoveComponent vbProj, "frmSettings"
    RemoveComponent vbProj, "frmSales"
    RemoveComponent vbProj, "frmInput"

    step = "3/5: mMainモジュール追加"
    Dim m As Object
    Set m = vbProj.VBComponents.Add(VBEXT_CT_STDMODULE)
    m.Name = "mMain"

    step = "4/5: mMainコード埋め込み"
    m.CodeModule.AddFromString GetMMainCode()

    step = "5/5: frmInput追加"
    BuildFrmInput vbProj

    MsgBox "セットアップ完了！" & vbCrLf & "Alt+F8 → Main を実行してください。", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "セットアップ失敗 [" & step & "]:" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Sub RemoveComponent(vbProj As Object, compName As String)
    Dim comp As Object
    On Error Resume Next
    Set comp = vbProj.VBComponents(compName)
    If Not comp Is Nothing Then vbProj.VBComponents.Remove comp
    On Error GoTo 0
End Sub

Private Sub BuildFrmInput(vbProj As Object)
    Dim comp As Object, n As Long, i As Long
    Set comp = vbProj.VBComponents.Add(VBEXT_CT_MSFORM)
    comp.Name = "frmInput"
    For i = 1 To 30 : DoEvents : Next i
    n = comp.CodeModule.CountOfLines
    If n > 0 Then comp.CodeModule.DeleteLines 1, n
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

Private Function GetFrmInputCode() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ── モジュールレベル変数 ─────────────────────────────" & vbCrLf
    s = s & "Private m_Files()    As String" & vbCrLf
    s = s & "Private m_Count      As Integer" & vbCrLf
    s = s & "Private m_Idx        As Integer" & vbCrLf
    s = s & "Private m_Processing As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ── ボタンイベント用 WithEvents ──────────────────────" & vbCrLf
    s = s & "Private WithEvents m_btnBrowse  As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents m_btnImport  As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents m_btnCancel  As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents m_btnPrev    As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents m_btnNext    As MSForms.CommandButton" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 初期化: コントロール生成 → 初期値セット" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    Me.Caption = ""在庫管理""" & vbCrLf
    s = s & "    Me.Width   = 650" & vbCrLf
    s = s & "    Me.Height  = 720" & vbCrLf
    s = s & "    Me.ScrollBars   = 2" & vbCrLf
    s = s & "    Me.ScrollHeight = 1000" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    CreateControls" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    m_Processing = False" & vbCrLf
    s = s & "    CtlText(""txtSheetName"") = g_SheetName" & vbCrLf
    s = s & "    m_btnImport.Caption = ""開始""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 全コントロールをランタイムで生成（Designer 不要）" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Private Sub CreateControls()" & vbCrLf
    s = s & "    ' ── 設定行 ──────────────────────────────────────────" & vbCrLf
    s = s & "    Lbl ""lblSheetName"",  8,  8, 70, 18, ""シート名""" & vbCrLf
    s = s & "    Txt ""txtSheetName"", 80,  6, 80, 20" & vbCrLf
    s = s & "    Lbl ""lblFolder"",   170,  8, 60, 18, ""フォルダ""" & vbCrLf
    s = s & "    Txt ""txtFolder"",   235,  6, 255, 20" & vbCrLf
    s = s & "    Set m_btnBrowse = Btn(""btnBrowse"", 495, 4, 48, 22, ""参照..."")" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' ── 左ペイン: 画像 ──────────────────────────────────" & vbCrLf
    s = s & "    Lbl ""lblProgress"",  8, 34, 290, 18, ""画像なし""" & vbCrLf
    s = s & "    Dim img As Object" & vbCrLf
    s = s & "    Set img = Me.Controls.Add(""Forms.Image.1"", ""imgPreview"", True)" & vbCrLf
    s = s & "    img.Left = 8 : img.Top = 55 : img.Width = 290 : img.Height = 360" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    img.PictureSizeMode = 3" & vbCrLf
    s = s & "    img.BorderStyle = 1" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "    Lbl ""lblFileName"",  8, 420, 290, 18, """"" & vbCrLf
    s = s & "    Set m_btnPrev = Btn(""btnPrev"",  8, 442, 42, 22, ""←"")" & vbCrLf
    s = s & "    Set m_btnNext = Btn(""btnNext"", 55, 442, 42, 22, ""→"")" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' ── 右ペイン: 全フィールド ──────────────────────────" & vbCrLf
    s = s & "    Dim bL As Single, bT As Single, rH As Single" & vbCrLf
    s = s & "    bL = 312 : bT = 34 : rH = 24" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim bLbl As Variant, bName As Variant, i As Integer" & vbCrLf
    s = s & "    bLbl  = Array(""仕入れ日(B1)"", ""回次(C1)"", ""車名/グレード(F1)"", ""年式/月(D1)"", _" & vbCrLf
    s = s & "                  ""車検(G1)"", ""評価点(H)"", ""車輌代(I)"", ""消費税(J)"", _" & vbCrLf
    s = s & "                  ""自税(K)"", ""リサイクル(L)"", ""落札料(M)"", ""合計(N)"", ""評価損(O)"", _" & vbCrLf
    s = s & "                  ""車輌番号(P1)"", ""仕入れ先(B2)"", ""出品番号(C2)"", ""色(D2)"", _" & vbCrLf
    s = s & "                  ""車台番号(F2)"", ""走行距離km(G2)"", ""所有者(P2)"", ""補足(T)"")" & vbCrLf
    s = s & "    bName = Array(""Date"", ""Session"", ""CarName"", ""Year"", _" & vbCrLf
    s = s & "                  ""Shaken"", ""Score"", ""Price"", ""Tax"", _" & vbCrLf
    s = s & "                  ""CarTax"", ""Recycle"", ""Fee"", ""Total"", ""Loss"", _" & vbCrLf
    s = s & "                  ""Plate"", ""Supplier"", ""LotNum"", ""Color"", _" & vbCrLf
    s = s & "                  ""Chassis"", ""Mileage"", ""Owner"", ""Memo"")" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    For i = 0 To UBound(bLbl)" & vbCrLf
    s = s & "        Lbl ""lbl"" & bName(i), bL,       bT + i * rH, 105, 18, CStr(bLbl(i))" & vbCrLf
    s = s & "        Txt ""txt"" & bName(i), bL + 108, bT + i * rH, 112, 20" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim sepT As Single : sepT = bT + (UBound(bLbl) + 1) * rH" & vbCrLf
    s = s & "    Lbl ""lblSaleSep"", bL, sepT, 220, 18, ""── 売上 ──────────""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim sLbl As Variant, sName As Variant" & vbCrLf
    s = s & "    sLbl  = Array(""名義変更(U)"", ""売上日(V1)"", ""売上先(V2)"", ""売上回次(W1)"", _" & vbCrLf
    s = s & "                  ""売上出品番号(W2)"", ""売上車輌代(X)"", ""売上消費税(Y)"", _" & vbCrLf
    s = s & "                  ""売上リサイクル(Z)"", ""売上合計(AA)"", ""入金日(AB)"")" & vbCrLf
    s = s & "    sName = Array(""Meigi"", ""SaleDate"", ""Buyer"", ""SaleSession"", _" & vbCrLf
    s = s & "                  ""SaleLot"", ""SalePrice"", ""SaleTax"", _" & vbCrLf
    s = s & "                  ""SaleRecycle"", ""SaleTotal"", ""PayDate"")" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim salT As Single : salT = sepT + rH" & vbCrLf
    s = s & "    For i = 0 To UBound(sLbl)" & vbCrLf
    s = s & "        Lbl ""lbl"" & sName(i), bL,       salT + i * rH, 105, 18, CStr(sLbl(i))" & vbCrLf
    s = s & "        Txt ""txt"" & sName(i), bL + 108, salT + i * rH, 112, 20" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim btnT As Single : btnT = salT + (UBound(sLbl) + 1) * rH + 8" & vbCrLf
    s = s & "    Set m_btnImport = Btn(""btnImport"", bL,       btnT, 100, 28, ""開始"")" & vbCrLf
    s = s & "    Set m_btnCancel = Btn(""btnCancel"", bL + 110, btnT, 100, 28, ""閉じる"")" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ── コントロール生成ヘルパー ─────────────────────────" & vbCrLf
    s = s & "Private Sub Lbl(n As String, L As Single, T As Single, W As Single, H As Single, cap As String)" & vbCrLf
    s = s & "    Dim c As Object" & vbCrLf
    s = s & "    Set c = Me.Controls.Add(""Forms.Label.1"", n, True)" & vbCrLf
    s = s & "    c.Left = L : c.Top = T : c.Width = W : c.Height = H : c.Caption = cap" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Txt(n As String, L As Single, T As Single, W As Single, H As Single)" & vbCrLf
    s = s & "    Dim c As Object" & vbCrLf
    s = s & "    Set c = Me.Controls.Add(""Forms.TextBox.1"", n, True)" & vbCrLf
    s = s & "    c.Left = L : c.Top = T : c.Width = W : c.Height = H" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function Btn(n As String, L As Single, T As Single, W As Single, H As Single, cap As String) As MSForms.CommandButton" & vbCrLf
    s = s & "    Dim c As MSForms.CommandButton" & vbCrLf
    s = s & "    Set c = Me.Controls.Add(""Forms.CommandButton.1"", n, True)" & vbCrLf
    s = s & "    c.Left = L : c.Top = T : c.Width = W : c.Height = H : c.Caption = cap" & vbCrLf
    s = s & "    Set Btn = c" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Property Get CtlText(n As String) As String" & vbCrLf
    s = s & "    CtlText = Me.Controls(n).Text" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Property Let CtlText(n As String, val As String)" & vbCrLf
    s = s & "    Me.Controls(n).Text = val" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' ボタンイベント" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Private Sub m_btnBrowse_Click()" & vbCrLf
    s = s & "    With Application.FileDialog(msoFileDialogFolderPicker)" & vbCrLf
    s = s & "        .Title = ""画像フォルダを選択してください""" & vbCrLf
    s = s & "        .AllowMultiSelect = False" & vbCrLf
    s = s & "        If .Show Then CtlText(""txtFolder"") = .SelectedItems(1)" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub m_btnImport_Click()" & vbCrLf
    s = s & "    If Not m_Processing Then StartProcessing Else ProcessCurrentImage" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub m_btnCancel_Click()" & vbCrLf
    s = s & "    If MsgBox(""処理を中止しますか？"", vbQuestion + vbYesNo) = vbYes Then Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub m_btnPrev_Click()" & vbCrLf
    s = s & "    If m_Idx > 0 Then m_Idx = m_Idx - 1 : RefreshImage" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub m_btnNext_Click()" & vbCrLf
    s = s & "    If m_Idx < m_Count - 1 Then m_Idx = m_Idx + 1 : RefreshImage" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 開始処理" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Private Sub StartProcessing()" & vbCrLf
    s = s & "    If Trim(CtlText(""txtSheetName"")) = """" Then" & vbCrLf
    s = s & "        MsgBox ""シート名を入力してください。"", vbExclamation : Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Trim(CtlText(""txtFolder"")) = """" Then" & vbCrLf
    s = s & "        MsgBox ""フォルダを選択してください。"", vbExclamation : Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Dir(Trim(CtlText(""txtFolder"")), vbDirectory) = """" Then" & vbCrLf
    s = s & "        MsgBox ""フォルダが見つかりません。"", vbExclamation : Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Sheets(Trim(CtlText(""txtSheetName"")))" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "    If ws Is Nothing Then" & vbCrLf
    s = s & "        MsgBox ""シート「"" & CtlText(""txtSheetName"") & ""」が見つかりません。"", vbExclamation : Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    g_SheetName  = Trim(CtlText(""txtSheetName""))" & vbCrLf
    s = s & "    g_FolderPath = Trim(CtlText(""txtFolder""))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim exts(3) As String" & vbCrLf
    s = s & "    exts(0) = ""*.jpg"" : exts(1) = ""*.jpeg"" : exts(2) = ""*.png"" : exts(3) = ""*.bmp""" & vbCrLf
    s = s & "    Dim col As New Collection, i As Integer, f As String" & vbCrLf
    s = s & "    For i = 0 To 3" & vbCrLf
    s = s & "        f = Dir(g_FolderPath & ""\"" & exts(i))" & vbCrLf
    s = s & "        Do While f <> """"" & vbCrLf
    s = s & "            col.Add g_FolderPath & ""\"" & f" & vbCrLf
    s = s & "            f = Dir()" & vbCrLf
    s = s & "        Loop" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If col.Count = 0 Then" & vbCrLf
    s = s & "        MsgBox ""画像ファイル（JPG/PNG/BMP）が見つかりません。"", vbInformation : Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ReDim m_Files(0 To col.Count - 1)" & vbCrLf
    s = s & "    For i = 0 To col.Count - 1" & vbCrLf
    s = s & "        m_Files(i) = col(i + 1)" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    m_Count = col.Count : m_Idx = 0 : m_Processing = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    LoadDefaults" & vbCrLf
    s = s & "    RefreshImage" & vbCrLf
    s = s & "    m_btnImport.Caption = ""取込""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 取込処理" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Private Sub ProcessCurrentImage()" & vbCrLf
    s = s & "    If Trim(CtlText(""txtCarName"")) = """" Then" & vbCrLf
    s = s & "        MsgBox ""車名を入力してください。"", vbExclamation : Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Sheets(g_SheetName)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim data As CarData" & vbCrLf
    s = s & "    data = CollectCarData()" & vbCrLf
    s = s & "    data.CarNumber = GetPeriodNumber(ws) & ""-"" & Format(GetNextNumber(ws), ""000"")" & vbCrLf
    s = s & "    WriteToSheet ws, data" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Trim(CtlText(""txtSaleDate"")) <> """" Then" & vbCrLf
    s = s & "        Dim sd As SaleData" & vbCrLf
    s = s & "        sd = CollectSaleData()" & vbCrLf
    s = s & "        sd.CarNumber = data.CarNumber" & vbCrLf
    s = s & "        WriteToSheetSales ws, sd" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    MoveToProcessed m_Files(m_Idx)" & vbCrLf
    s = s & "    m_Idx = m_Idx + 1" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If m_Idx >= m_Count Then" & vbCrLf
    s = s & "        MsgBox m_Count & "" 件の処理が完了しました。"", vbInformation" & vbCrLf
    s = s & "        Unload Me" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        LoadDefaults" & vbCrLf
    s = s & "        RefreshImage" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' デフォルト値（AI-OCR有効化後はここを置き換える）" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Private Sub LoadDefaults()" & vbCrLf
    s = s & "    ' === AI-OCR ここから（API課金のため無効化中）===" & vbCrLf
    s = s & "    ' Dim d As CarData : d = CallClaudeOCR(m_Files(m_Idx))" & vbCrLf
    s = s & "    ' CtlText(""txtCarName"") = d.CarName ... 以下同様" & vbCrLf
    s = s & "    ' === AI-OCR ここまで ===" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    CtlText(""txtDate"")        = Format(Now, ""yyyy/m/d"")" & vbCrLf
    s = s & "    CtlText(""txtSession"")     = """"" & vbCrLf
    s = s & "    CtlText(""txtCarName"")     = ""スズキ クロスビー MZ""" & vbCrLf
    s = s & "    CtlText(""txtYear"")        = ""2021/7""" & vbCrLf
    s = s & "    CtlText(""txtShaken"")      = """"" & vbCrLf
    s = s & "    CtlText(""txtScore"")       = ""4.5""" & vbCrLf
    s = s & "    CtlText(""txtPrice"")       = ""980000""" & vbCrLf
    s = s & "    CtlText(""txtTax"")         = ""98000""" & vbCrLf
    s = s & "    CtlText(""txtCarTax"")      = ""35400""" & vbCrLf
    s = s & "    CtlText(""txtRecycle"")     = ""12000""" & vbCrLf
    s = s & "    CtlText(""txtFee"")         = ""32000""" & vbCrLf
    s = s & "    CtlText(""txtTotal"")       = """"" & vbCrLf
    s = s & "    CtlText(""txtLoss"")        = """"" & vbCrLf
    s = s & "    CtlText(""txtPlate"")       = """"" & vbCrLf
    s = s & "    CtlText(""txtSupplier"")    = ""USS大阪""" & vbCrLf
    s = s & "    CtlText(""txtLotNum"")      = ""12345""" & vbCrLf
    s = s & "    CtlText(""txtColor"")       = ""ホワイトパール""" & vbCrLf
    s = s & "    CtlText(""txtChassis"")     = ""MYN15S-100001""" & vbCrLf
    s = s & "    CtlText(""txtMileage"")     = ""35000""" & vbCrLf
    s = s & "    CtlText(""txtOwner"")       = """"" & vbCrLf
    s = s & "    CtlText(""txtMemo"")        = """"" & vbCrLf
    s = s & "    CtlText(""txtMeigi"")       = """"" & vbCrLf
    s = s & "    CtlText(""txtSaleDate"")    = """"" & vbCrLf
    s = s & "    CtlText(""txtBuyer"")       = """"" & vbCrLf
    s = s & "    CtlText(""txtSaleSession"") = """"" & vbCrLf
    s = s & "    CtlText(""txtSaleLot"")     = """"" & vbCrLf
    s = s & "    CtlText(""txtSalePrice"")   = """"" & vbCrLf
    s = s & "    CtlText(""txtSaleTax"")     = """"" & vbCrLf
    s = s & "    CtlText(""txtSaleRecycle"") = """"" & vbCrLf
    s = s & "    CtlText(""txtSaleTotal"")   = """"" & vbCrLf
    s = s & "    CtlText(""txtPayDate"")     = """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub RefreshImage()" & vbCrLf
    s = s & "    If m_Count = 0 Then Exit Sub" & vbCrLf
    s = s & "    Me.Controls(""lblProgress"").Caption = (m_Idx + 1) & "" / "" & m_Count" & vbCrLf
    s = s & "    Me.Controls(""lblFileName"").Caption = Mid(m_Files(m_Idx), InStrRev(m_Files(m_Idx), ""\"") + 1)" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Me.Controls(""imgPreview"").Picture = LoadPicture(m_Files(m_Idx))" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' データ収集" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Private Function CollectCarData() As CarData" & vbCrLf
    s = s & "    Dim d As CarData" & vbCrLf
    s = s & "    d.PurchaseDate = Trim(CtlText(""txtDate""))" & vbCrLf
    s = s & "    d.Session      = Trim(CtlText(""txtSession""))" & vbCrLf
    s = s & "    d.CarName      = Trim(CtlText(""txtCarName""))" & vbCrLf
    s = s & "    d.YearMonth    = Trim(CtlText(""txtYear""))" & vbCrLf
    s = s & "    d.Shaken       = Trim(CtlText(""txtShaken""))" & vbCrLf
    s = s & "    d.Score        = Trim(CtlText(""txtScore""))" & vbCrLf
    s = s & "    d.Price        = Trim(CtlText(""txtPrice""))" & vbCrLf
    s = s & "    d.Tax          = Trim(CtlText(""txtTax""))" & vbCrLf
    s = s & "    d.CarTax       = Trim(CtlText(""txtCarTax""))" & vbCrLf
    s = s & "    d.Recycle      = Trim(CtlText(""txtRecycle""))" & vbCrLf
    s = s & "    d.AuctionFee   = Trim(CtlText(""txtFee""))" & vbCrLf
    s = s & "    d.Total        = Trim(CtlText(""txtTotal""))" & vbCrLf
    s = s & "    d.Loss         = Trim(CtlText(""txtLoss""))" & vbCrLf
    s = s & "    d.Plate        = Trim(CtlText(""txtPlate""))" & vbCrLf
    s = s & "    d.Supplier     = Trim(CtlText(""txtSupplier""))" & vbCrLf
    s = s & "    d.LotNumber    = Trim(CtlText(""txtLotNum""))" & vbCrLf
    s = s & "    d.Color        = Trim(CtlText(""txtColor""))" & vbCrLf
    s = s & "    d.Chassis      = Trim(CtlText(""txtChassis""))" & vbCrLf
    s = s & "    d.Mileage      = Trim(CtlText(""txtMileage""))" & vbCrLf
    s = s & "    d.Owner        = Trim(CtlText(""txtOwner""))" & vbCrLf
    s = s & "    d.Memo         = Trim(CtlText(""txtMemo""))" & vbCrLf
    s = s & "    CollectCarData = d" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function CollectSaleData() As SaleData" & vbCrLf
    s = s & "    Dim d As SaleData" & vbCrLf
    s = s & "    d.Meigi       = Trim(CtlText(""txtMeigi""))" & vbCrLf
    s = s & "    d.SaleDate    = Trim(CtlText(""txtSaleDate""))" & vbCrLf
    s = s & "    d.Buyer       = Trim(CtlText(""txtBuyer""))" & vbCrLf
    s = s & "    d.SaleSession = Trim(CtlText(""txtSaleSession""))" & vbCrLf
    s = s & "    d.SaleLot     = Trim(CtlText(""txtSaleLot""))" & vbCrLf
    s = s & "    d.SalePrice   = Trim(CtlText(""txtSalePrice""))" & vbCrLf
    s = s & "    d.SaleTax     = Trim(CtlText(""txtSaleTax""))" & vbCrLf
    s = s & "    d.SaleRecycle = Trim(CtlText(""txtSaleRecycle""))" & vbCrLf
    s = s & "    d.SaleTotal   = Trim(CtlText(""txtSaleTotal""))" & vbCrLf
    s = s & "    d.PaymentDate = Trim(CtlText(""txtPayDate""))" & vbCrLf
    s = s & "    CollectSaleData = d" & vbCrLf
    s = s & "End Function" & vbCrLf
    GetFrmInputCode = s
End Function

