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
    On Error GoTo ErrHandler

    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    RemoveComponent vbProj, "mMain"
    RemoveComponent vbProj, "frmSettings"
    RemoveComponent vbProj, "frmSales"
    RemoveComponent vbProj, "frmInput"

    Dim m As Object
    Set m = vbProj.VBComponents.Add(VBEXT_CT_STDMODULE)
    m.Name = "mMain"
    m.CodeModule.AddFromString GetMMainCode()

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
              "「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」を確認してください。"
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

Private Sub BuildFrmInput(vbProj As Object)
    Dim comp As Object
    Set comp = vbProj.VBComponents.Add(VBEXT_CT_MSFORM)
    comp.Name = "frmInput"

    ' Properties へのアクセスは全て OERN で保護
    On Error Resume Next
    comp.Properties("Caption")      = "在庫管理"
    comp.Properties("Width")        = 640
    comp.Properties("Height")       = 660
    comp.Properties("ScrollBars")   = 2
    comp.Properties("ScrollHeight") = 980
    On Error GoTo 0

    Dim d As Object
    Set d = GetDesigner(comp)

    ' Designer 経由でもサイズ設定（より確実）
    On Error Resume Next
    d.Width  = 640
    d.Height = 660
    On Error GoTo 0

    ' ── 上段: 設定 ──────────────────────────────────
    AddCtl d, "Forms.Label.1",         "lblSheetName", 8,  8, 70, 18, "シート名"
    AddCtl d, "Forms.TextBox.1",       "txtSheetName", 80, 6, 80, 20
    AddCtl d, "Forms.Label.1",         "lblFolder",   170, 8, 60, 18, "フォルダ"
    AddCtl d, "Forms.TextBox.1",       "txtFolder",   235, 6, 260, 20
    AddCtl d, "Forms.CommandButton.1", "btnBrowse",   500, 4,  40, 22, "参照..."

    ' ── 左ペイン: 画像 ─────────────────────────────
    AddCtl d, "Forms.Label.1", "lblProgress", 8, 34, 290, 18, "画像なし"
    Dim img As Object
    Set img = AddCtl(d, "Forms.Image.1", "imgPreview", 8, 54, 290, 370)
    On Error Resume Next
    img.PictureSizeMode = 3
    img.BorderStyle = 1
    On Error GoTo 0
    AddCtl d, "Forms.Label.1",         "lblFileName", 8, 428, 290, 18, ""
    AddCtl d, "Forms.CommandButton.1", "btnPrev",     8, 450,  40, 22, "←"
    AddCtl d, "Forms.CommandButton.1", "btnNext",    52, 450,  40, 22, "→"

    ' ── 右ペイン: 全フィールド ──────────────────────
    Dim bL As Single, bT As Single, rH As Single
    bL = 310 : bT = 34 : rH = 24

    ' 仕入れ
    Dim buyLabels As Variant, buyNames As Variant, i As Integer
    buyLabels = Array("仕入れ日(B1)", "回次(C1)", "車名/グレード(F1)", "年式/月(D1)", _
                      "車検(G1)", "評価点(H)", "車輌代(I)", "消費税(J)", _
                      "自税(K)", "リサイクル(L)", "落札料(M)", "合計(N)", "評価損(O)", _
                      "車輌番号(P1)", "仕入れ先(B2)", "出品番号(C2)", "色(D2)", _
                      "車台番号(F2)", "走行距離km(G2)", "所有者(P2)", "補足(T)")
    buyNames  = Array("Date", "Session", "CarName", "Year", _
                      "Shaken", "Score", "Price", "Tax", _
                      "CarTax", "Recycle", "Fee", "Total", "Loss", _
                      "Plate", "Supplier", "LotNum", "Color", _
                      "Chassis", "Mileage", "Owner", "Memo")

    For i = 0 To UBound(buyLabels)
        AddCtl d, "Forms.Label.1",   "lbl" & buyNames(i), bL,      bT + i * rH, 105, 18, CStr(buyLabels(i))
        AddCtl d, "Forms.TextBox.1", "txt" & buyNames(i), bL + 108, bT + i * rH, 112, 20
    Next i

    ' 売上セパレータ
    Dim sepT As Single
    sepT = bT + (UBound(buyLabels) + 1) * rH
    AddCtl d, "Forms.Label.1", "lblSaleSep", bL, sepT, 220, 18, "── 売上 ──────────────"

    ' 売上
    Dim salLabels As Variant, salNames As Variant
    salLabels = Array("名義変更(U)", "売上日(V1)", "売上先(V2)", "売上回次(W1)", _
                      "売上出品番号(W2)", "売上車輌代(X)", "売上消費税(Y)", _
                      "売上リサイクル(Z)", "売上合計(AA)", "入金日(AB)")
    salNames  = Array("Meigi", "SaleDate", "Buyer", "SaleSession", _
                      "SaleLot", "SalePrice", "SaleTax", _
                      "SaleRecycle", "SaleTotal", "PayDate")

    Dim salT As Single
    salT = sepT + rH
    For i = 0 To UBound(salLabels)
        AddCtl d, "Forms.Label.1",   "lbl" & salNames(i), bL,      salT + i * rH, 105, 18, CStr(salLabels(i))
        AddCtl d, "Forms.TextBox.1", "txt" & salNames(i), bL + 108, salT + i * rH, 112, 20
    Next i

    ' ── ボタン ────────────────────────────────────
    Dim btnT As Single
    btnT = salT + (UBound(salLabels) + 1) * rH + 8
    Dim btnImp As Object
    Set btnImp = AddCtl(d, "Forms.CommandButton.1", "btnImport", bL,      btnT, 100, 28, "開始")
    On Error Resume Next
    btnImp.Default = True
    On Error GoTo 0
    AddCtl d, "Forms.CommandButton.1", "btnCancel", bL + 110, btnT, 100, 28, "閉じる"

    comp.CodeModule.AddFromString GetFrmInputCode()
End Sub

Private Function GetMMainCode() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public g_SheetName  As String" & vbCrLf
    s = s & "Public g_FolderPath As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 仕入れデータ（列 A-T）" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Public Type CarData" & vbCrLf
    s = s & "    CarNumber    As String   ' A:  仕入番号（自動）" & vbCrLf
    s = s & "    PurchaseDate As String   ' B1: 仕入れ日" & vbCrLf
    s = s & "    Supplier     As String   ' B2: 仕入れ先" & vbCrLf
    s = s & "    Session      As String   ' C1: 回次" & vbCrLf
    s = s & "    LotNumber    As String   ' C2: 出品番号" & vbCrLf
    s = s & "    YearMonth    As String   ' D1: 年式/月" & vbCrLf
    s = s & "    Color        As String   ' D2: 色" & vbCrLf
    s = s & "    CarName      As String   ' F1: 車名（グレード含む）" & vbCrLf
    s = s & "    Chassis      As String   ' F2: 車台番号" & vbCrLf
    s = s & "    Shaken       As String   ' G1: 車検" & vbCrLf
    s = s & "    Mileage      As Variant  ' G2: 走行距離" & vbCrLf
    s = s & "    Score        As Variant  ' H:  評価点" & vbCrLf
    s = s & "    Price        As Variant  ' I:  車輌代" & vbCrLf
    s = s & "    Tax          As Variant  ' J:  消費税" & vbCrLf
    s = s & "    CarTax       As Variant  ' K:  自税" & vbCrLf
    s = s & "    Recycle      As Variant  ' L:  リサイクル" & vbCrLf
    s = s & "    AuctionFee   As Variant  ' M:  落札料" & vbCrLf
    s = s & "    Total        As Variant  ' N:  合計" & vbCrLf
    s = s & "    Loss         As Variant  ' O:  評価損" & vbCrLf
    s = s & "    Plate        As String   ' P1: 車輌番号" & vbCrLf
    s = s & "    Owner        As String   ' P2: 所有者" & vbCrLf
    s = s & "    Memo         As String   ' T:  補足" & vbCrLf
    s = s & "End Type" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 売上データ（列 U-AB）" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Public Type SaleData" & vbCrLf
    s = s & "    CarNumber   As String   ' 仕入番号（検索キー）" & vbCrLf
    s = s & "    Meigi       As String   ' U:  名義変更" & vbCrLf
    s = s & "    SaleDate    As String   ' V1: 売上日" & vbCrLf
    s = s & "    Buyer       As String   ' V2: 売上先" & vbCrLf
    s = s & "    SaleSession As String   ' W1: 売上回次" & vbCrLf
    s = s & "    SaleLot     As String   ' W2: 売上出品番号" & vbCrLf
    s = s & "    SalePrice   As Variant  ' X:  売上車輌代" & vbCrLf
    s = s & "    SaleTax     As Variant  ' Y:  売上消費税" & vbCrLf
    s = s & "    SaleRecycle As Variant  ' Z:  売上リサイクル" & vbCrLf
    s = s & "    SaleTotal   As Variant  ' AA: 売上合計" & vbCrLf
    s = s & "    PaymentDate As String   ' AB: 入金日" & vbCrLf
    s = s & "End Type" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' エントリポイント" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Sub Main()" & vbCrLf
    s = s & "    g_SheetName  = ""27期""" & vbCrLf
    s = s & "    g_FolderPath = """"" & vbCrLf
    s = s & "    frmInput.Show" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' A列連番の最大値+1を返す" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Function GetNextNumber(ws As Worksheet) As Long" & vbCrLf
    s = s & "    Dim i As Long, maxNum As Long, cellVal As String, parts() As String" & vbCrLf
    s = s & "    maxNum = 0" & vbCrLf
    s = s & "    For i = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row To 1 Step -1" & vbCrLf
    s = s & "        On Error Resume Next" & vbCrLf
    s = s & "        cellVal = Trim(CStr(ws.Cells(i, 1).Value2))" & vbCrLf
    s = s & "        On Error GoTo 0" & vbCrLf
    s = s & "        If InStr(cellVal, ""-"") > 0 And InStr(cellVal, ""-"") = InStrRev(cellVal, ""-"") Then" & vbCrLf
    s = s & "            parts = Split(cellVal, ""-"")" & vbCrLf
    s = s & "            If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then" & vbCrLf
    s = s & "                Dim n As Long : n = CLng(parts(1))" & vbCrLf
    s = s & "                If n > maxNum Then maxNum = n" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    GetNextNumber = maxNum + 1" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 次の書き込み行を返す（Q列""着""を末尾として算出）" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Function GetNextWriteRow(ws As Worksheet) As Long" & vbCrLf
    s = s & "    Dim lastRow As Long, qVal As String" & vbCrLf
    s = s & "    lastRow = ws.Cells(ws.Rows.Count, 17).End(xlUp).Row" & vbCrLf
    s = s & "    If lastRow < 10 Then GetNextWriteRow = 10 : Exit Function" & vbCrLf
    s = s & "    qVal = Trim(CStr(ws.Cells(lastRow, 17).Value2))" & vbCrLf
    s = s & "    Select Case qVal" & vbCrLf
    s = s & "        Case ""着""   : GetNextWriteRow = lastRow + 1" & vbCrLf
    s = s & "        Case ""予定"" : GetNextWriteRow = lastRow + 2" & vbCrLf
    s = s & "        Case Else   : GetNextWriteRow = lastRow + 1" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' シート名から期番号取得（""27期"" → ""27""）" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Function GetPeriodNumber(ws As Worksheet) As String" & vbCrLf
    s = s & "    Dim s As String, i As Integer, result As String" & vbCrLf
    s = s & "    s = ws.Name : result = """"" & vbCrLf
    s = s & "    For i = 1 To Len(s)" & vbCrLf
    s = s & "        If IsNumeric(Mid(s, i, 1)) Then result = result & Mid(s, i, 1) Else Exit For" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    If result = """" Then result = ""27""" & vbCrLf
    s = s & "    GetPeriodNumber = result" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 仕入番号からシート行番号を返す（見つからなければ 0）" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Function FindCarRow(ws As Worksheet, carNum As String) As Long" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    For i = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    s = s & "        If Trim(CStr(ws.Cells(i, 1).Value2)) = Trim(carNum) Then" & vbCrLf
    s = s & "            FindCarRow = i : Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    FindCarRow = 0" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 仕入れデータをシートに2行書き込み" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Sub WriteToSheet(ws As Worksheet, data As CarData)" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    r = GetNextWriteRow(ws)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ws.Cells(r,  1).Value2 = data.CarNumber" & vbCrLf
    s = s & "    ws.Cells(r,  4).Value2 = data.YearMonth" & vbCrLf
    s = s & "    ws.Cells(r,  6).Value2 = data.CarName" & vbCrLf
    s = s & "    ws.Cells(r, 17).Value2 = ""予定""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If data.PurchaseDate <> """" Then ws.Cells(r,  2).Value2 = data.PurchaseDate" & vbCrLf
    s = s & "    If data.Session      <> """" Then ws.Cells(r,  3).Value2 = data.Session" & vbCrLf
    s = s & "    If data.Shaken       <> """" Then ws.Cells(r,  7).Value2 = data.Shaken" & vbCrLf
    s = s & "    If data.Score        <> """" Then ws.Cells(r,  8).Value2 = data.Score" & vbCrLf
    s = s & "    If data.Price        <> """" Then ws.Cells(r,  9).Value2 = CLng(data.Price)" & vbCrLf
    s = s & "    If data.Tax          <> """" Then ws.Cells(r, 10).Value2 = CLng(data.Tax)" & vbCrLf
    s = s & "    If data.CarTax       <> """" Then ws.Cells(r, 11).Value2 = CLng(data.CarTax)" & vbCrLf
    s = s & "    If data.Recycle      <> """" Then ws.Cells(r, 12).Value2 = CLng(data.Recycle)" & vbCrLf
    s = s & "    If data.AuctionFee   <> """" Then ws.Cells(r, 13).Value2 = CLng(data.AuctionFee)" & vbCrLf
    s = s & "    If data.Total        <> """" Then ws.Cells(r, 14).Value2 = CLng(data.Total)" & vbCrLf
    s = s & "    If data.Loss         <> """" Then ws.Cells(r, 15).Value2 = CLng(data.Loss)" & vbCrLf
    s = s & "    If data.Plate        <> """" Then ws.Cells(r, 16).Value2 = data.Plate" & vbCrLf
    s = s & "    If data.Memo         <> """" Then ws.Cells(r, 20).Value2 = data.Memo" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ws.Cells(r+1, 17).Value2 = ""着""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If data.Supplier  <> """" Then ws.Cells(r+1,  2).Value2 = data.Supplier" & vbCrLf
    s = s & "    If data.LotNumber <> """" Then ws.Cells(r+1,  3).Value2 = data.LotNumber" & vbCrLf
    s = s & "    If data.Color     <> """" Then ws.Cells(r+1,  4).Value2 = data.Color" & vbCrLf
    s = s & "    If data.Chassis   <> """" Then ws.Cells(r+1,  6).Value2 = data.Chassis" & vbCrLf
    s = s & "    If data.Mileage   <> """" Then ws.Cells(r+1,  7).Value2 = CLng(data.Mileage)" & vbCrLf
    s = s & "    If data.Owner     <> """" Then ws.Cells(r+1, 16).Value2 = data.Owner" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 売上データを既存行に書き込み" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Sub WriteToSheetSales(ws As Worksheet, data As SaleData)" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    r = FindCarRow(ws, data.CarNumber)" & vbCrLf
    s = s & "    If r = 0 Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If data.Meigi       <> """" Then ws.Cells(r,   21).Value2 = data.Meigi" & vbCrLf
    s = s & "    If data.SaleDate    <> """" Then ws.Cells(r,   22).Value2 = data.SaleDate" & vbCrLf
    s = s & "    If data.Buyer       <> """" Then ws.Cells(r+1, 22).Value2 = data.Buyer" & vbCrLf
    s = s & "    If data.SaleSession <> """" Then ws.Cells(r,   23).Value2 = data.SaleSession" & vbCrLf
    s = s & "    If data.SaleLot     <> """" Then ws.Cells(r+1, 23).Value2 = data.SaleLot" & vbCrLf
    s = s & "    If data.SalePrice   <> """" Then ws.Cells(r,   24).Value2 = CLng(data.SalePrice)" & vbCrLf
    s = s & "    If data.SaleTax     <> """" Then ws.Cells(r,   25).Value2 = CLng(data.SaleTax)" & vbCrLf
    s = s & "    If data.SaleRecycle <> """" Then ws.Cells(r,   26).Value2 = CLng(data.SaleRecycle)" & vbCrLf
    s = s & "    If data.SaleTotal   <> """" Then ws.Cells(r,   27).Value2 = CLng(data.SaleTotal)" & vbCrLf
    s = s & "    If data.PaymentDate <> """" Then ws.Cells(r,   28).Value2 = data.PaymentDate" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 画像を「処理済み」サブフォルダへ移動" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Sub MoveToProcessed(filePath As String)" & vbCrLf
    s = s & "    On Error GoTo ErrHandler" & vbCrLf
    s = s & "    Dim folder As String, fileName As String, processed As String, dest As String" & vbCrLf
    s = s & "    folder    = Left(filePath, InStrRev(filePath, ""\""))" & vbCrLf
    s = s & "    fileName  = Mid(filePath, InStrRev(filePath, ""\"") + 1)" & vbCrLf
    s = s & "    processed = folder & ""処理済み\""" & vbCrLf
    s = s & "    dest      = processed & fileName" & vbCrLf
    s = s & "    If Dir(processed, vbDirectory) = """" Then MkDir processed" & vbCrLf
    s = s & "    If Dir(dest) <> """" Then" & vbCrLf
    s = s & "        Dim dot As Integer, base As String, ext As String, cnt As Integer" & vbCrLf
    s = s & "        dot = InStrRev(fileName, ""."")" & vbCrLf
    s = s & "        base = Left(fileName, dot - 1) : ext = Mid(fileName, dot) : cnt = 1" & vbCrLf
    s = s & "        Do While Dir(processed & base & ""_"" & cnt & ext) <> """"" & vbCrLf
    s = s & "            cnt = cnt + 1" & vbCrLf
    s = s & "        Loop" & vbCrLf
    s = s & "        dest = processed & base & ""_"" & cnt & ext" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    FileCopy filePath, dest" & vbCrLf
    s = s & "    Kill filePath" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrHandler:" & vbCrLf
    s = s & "    MsgBox ""ファイル移動に失敗しました:"" & vbCrLf & filePath & vbCrLf & Err.Description, vbExclamation" & vbCrLf
    s = s & "End Sub" & vbCrLf
    GetMMainCode = s
End Function

Private Function GetFrmInputCode() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private m_Files()    As String" & vbCrLf
    s = s & "Private m_Count      As Integer" & vbCrLf
    s = s & "Private m_Idx        As Integer" & vbCrLf
    s = s & "Private m_Processing As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 初期化" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    txtSheetName.Text = g_SheetName" & vbCrLf
    s = s & "    m_Processing      = False" & vbCrLf
    s = s & "    btnImport.Caption = ""開始""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' フォルダ参照" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Private Sub btnBrowse_Click()" & vbCrLf
    s = s & "    With Application.FileDialog(msoFileDialogFolderPicker)" & vbCrLf
    s = s & "        .Title = ""画像フォルダを選択してください""" & vbCrLf
    s = s & "        .AllowMultiSelect = False" & vbCrLf
    s = s & "        If .Show Then txtFolder.Text = .SelectedItems(1)" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 取込ボタン（開始 / 取込 で挙動切り替え）" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Private Sub btnImport_Click()" & vbCrLf
    s = s & "    If Not m_Processing Then" & vbCrLf
    s = s & "        StartProcessing" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        ProcessCurrentImage" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 開始処理: バリデーション → 画像収集 → 1枚目表示" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Private Sub StartProcessing()" & vbCrLf
    s = s & "    If Trim(txtSheetName.Text) = """" Then" & vbCrLf
    s = s & "        MsgBox ""シート名を入力してください。"", vbExclamation" & vbCrLf
    s = s & "        txtSheetName.SetFocus : Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Trim(txtFolder.Text) = """" Then" & vbCrLf
    s = s & "        MsgBox ""フォルダを選択してください。"", vbExclamation : Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Dir(Trim(txtFolder.Text), vbDirectory) = """" Then" & vbCrLf
    s = s & "        MsgBox ""フォルダが見つかりません。"", vbExclamation : Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Sheets(Trim(txtSheetName.Text))" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "    If ws Is Nothing Then" & vbCrLf
    s = s & "        MsgBox ""シート「"" & txtSheetName.Text & ""」が見つかりません。"", vbExclamation : Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    g_SheetName  = Trim(txtSheetName.Text)" & vbCrLf
    s = s & "    g_FolderPath = Trim(txtFolder.Text)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim exts(3) As String" & vbCrLf
    s = s & "    exts(0) = ""*.jpg"" : exts(1) = ""*.jpeg"" : exts(2) = ""*.png"" : exts(3) = ""*.bmp""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim col As New Collection" & vbCrLf
    s = s & "    Dim i As Integer, f As String" & vbCrLf
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
    s = s & "    m_Count      = col.Count" & vbCrLf
    s = s & "    m_Idx        = 0" & vbCrLf
    s = s & "    m_Processing = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    LoadDefaults" & vbCrLf
    s = s & "    RefreshImage" & vbCrLf
    s = s & "    btnImport.Caption = ""取込""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 取込処理: 書き込み → 次の画像 or 完了" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Private Sub ProcessCurrentImage()" & vbCrLf
    s = s & "    If Trim(txtCarName.Text) = """" Then" & vbCrLf
    s = s & "        MsgBox ""車名を入力してください。"", vbExclamation" & vbCrLf
    s = s & "        txtCarName.SetFocus : Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Sheets(g_SheetName)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Dim data As CarData" & vbCrLf
    s = s & "    data           = CollectCarData()" & vbCrLf
    s = s & "    data.CarNumber = GetPeriodNumber(ws) & ""-"" & Format(GetNextNumber(ws), ""000"")" & vbCrLf
    s = s & "    WriteToSheet ws, data" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Trim(txtSaleDate.Text) <> """" Then" & vbCrLf
    s = s & "        Dim sdata As SaleData" & vbCrLf
    s = s & "        sdata           = CollectSaleData()" & vbCrLf
    s = s & "        sdata.CarNumber = data.CarNumber" & vbCrLf
    s = s & "        WriteToSheetSales ws, sdata" & vbCrLf
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
    s = s & "' デフォルト値セット（AI-OCR有効化後はここを置き換える）" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Private Sub LoadDefaults()" & vbCrLf
    s = s & "    ' === AI-OCR ここから（API課金のため無効化中）===" & vbCrLf
    s = s & "    ' Dim d As CarData" & vbCrLf
    s = s & "    ' d = CallClaudeOCR(m_Files(m_Idx))" & vbCrLf
    s = s & "    ' txtDate.Text        = d.PurchaseDate" & vbCrLf
    s = s & "    ' txtSession.Text     = d.Session" & vbCrLf
    s = s & "    ' txtCarName.Text     = d.CarName" & vbCrLf
    s = s & "    ' ..." & vbCrLf
    s = s & "    ' === AI-OCR ここまで ===" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtDate.Text     = Format(Now, ""yyyy/m/d"")" & vbCrLf
    s = s & "    txtSession.Text  = """"" & vbCrLf
    s = s & "    txtCarName.Text  = ""スズキ クロスビー MZ""" & vbCrLf
    s = s & "    txtYear.Text     = ""2021/7""" & vbCrLf
    s = s & "    txtShaken.Text   = """"" & vbCrLf
    s = s & "    txtScore.Text    = ""4.5""" & vbCrLf
    s = s & "    txtPrice.Text    = ""980000""" & vbCrLf
    s = s & "    txtTax.Text      = ""98000""" & vbCrLf
    s = s & "    txtCarTax.Text   = ""35400""" & vbCrLf
    s = s & "    txtRecycle.Text  = ""12000""" & vbCrLf
    s = s & "    txtFee.Text      = ""32000""" & vbCrLf
    s = s & "    txtTotal.Text    = """"" & vbCrLf
    s = s & "    txtLoss.Text     = """"" & vbCrLf
    s = s & "    txtPlate.Text    = """"" & vbCrLf
    s = s & "    txtSupplier.Text = ""USS大阪""" & vbCrLf
    s = s & "    txtLotNum.Text   = ""12345""" & vbCrLf
    s = s & "    txtColor.Text    = ""ホワイトパール""" & vbCrLf
    s = s & "    txtChassis.Text  = ""MYN15S-100001""" & vbCrLf
    s = s & "    txtMileage.Text  = ""35000""" & vbCrLf
    s = s & "    txtOwner.Text    = """"" & vbCrLf
    s = s & "    txtMemo.Text     = """"" & vbCrLf
    s = s & "    txtMeigi.Text    = """"" & vbCrLf
    s = s & "    txtSaleDate.Text = """"" & vbCrLf
    s = s & "    txtBuyer.Text    = """"" & vbCrLf
    s = s & "    txtSaleSession.Text  = """"" & vbCrLf
    s = s & "    txtSaleLot.Text      = """"" & vbCrLf
    s = s & "    txtSalePrice.Text    = """"" & vbCrLf
    s = s & "    txtSaleTax.Text      = """"" & vbCrLf
    s = s & "    txtSaleRecycle.Text  = """"" & vbCrLf
    s = s & "    txtSaleTotal.Text    = """"" & vbCrLf
    s = s & "    txtPayDate.Text      = """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 画像表示更新" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Private Sub RefreshImage()" & vbCrLf
    s = s & "    If m_Count = 0 Then Exit Sub" & vbCrLf
    s = s & "    lblProgress.Caption = (m_Idx + 1) & "" / "" & m_Count" & vbCrLf
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
    s = s & "Private Sub btnCancel_Click()" & vbCrLf
    s = s & "    If MsgBox(""処理を中止しますか？"", vbQuestion + vbYesNo) = vbYes Then Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' フォームから CarData を収集" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Private Function CollectCarData() As CarData" & vbCrLf
    s = s & "    Dim d As CarData" & vbCrLf
    s = s & "    d.PurchaseDate = Trim(txtDate.Text)" & vbCrLf
    s = s & "    d.Session      = Trim(txtSession.Text)" & vbCrLf
    s = s & "    d.CarName      = Trim(txtCarName.Text)" & vbCrLf
    s = s & "    d.YearMonth    = Trim(txtYear.Text)" & vbCrLf
    s = s & "    d.Shaken       = Trim(txtShaken.Text)" & vbCrLf
    s = s & "    d.Score        = Trim(txtScore.Text)" & vbCrLf
    s = s & "    d.Price        = Trim(txtPrice.Text)" & vbCrLf
    s = s & "    d.Tax          = Trim(txtTax.Text)" & vbCrLf
    s = s & "    d.CarTax       = Trim(txtCarTax.Text)" & vbCrLf
    s = s & "    d.Recycle      = Trim(txtRecycle.Text)" & vbCrLf
    s = s & "    d.AuctionFee   = Trim(txtFee.Text)" & vbCrLf
    s = s & "    d.Total        = Trim(txtTotal.Text)" & vbCrLf
    s = s & "    d.Loss         = Trim(txtLoss.Text)" & vbCrLf
    s = s & "    d.Plate        = Trim(txtPlate.Text)" & vbCrLf
    s = s & "    d.Supplier     = Trim(txtSupplier.Text)" & vbCrLf
    s = s & "    d.LotNumber    = Trim(txtLotNum.Text)" & vbCrLf
    s = s & "    d.Color        = Trim(txtColor.Text)" & vbCrLf
    s = s & "    d.Chassis      = Trim(txtChassis.Text)" & vbCrLf
    s = s & "    d.Mileage      = Trim(txtMileage.Text)" & vbCrLf
    s = s & "    d.Owner        = Trim(txtOwner.Text)" & vbCrLf
    s = s & "    d.Memo         = Trim(txtMemo.Text)" & vbCrLf
    s = s & "    CollectCarData = d" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' フォームから SaleData を収集" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Private Function CollectSaleData() As SaleData" & vbCrLf
    s = s & "    Dim d As SaleData" & vbCrLf
    s = s & "    d.Meigi       = Trim(txtMeigi.Text)" & vbCrLf
    s = s & "    d.SaleDate    = Trim(txtSaleDate.Text)" & vbCrLf
    s = s & "    d.Buyer       = Trim(txtBuyer.Text)" & vbCrLf
    s = s & "    d.SaleSession = Trim(txtSaleSession.Text)" & vbCrLf
    s = s & "    d.SaleLot     = Trim(txtSaleLot.Text)" & vbCrLf
    s = s & "    d.SalePrice   = Trim(txtSalePrice.Text)" & vbCrLf
    s = s & "    d.SaleTax     = Trim(txtSaleTax.Text)" & vbCrLf
    s = s & "    d.SaleRecycle = Trim(txtSaleRecycle.Text)" & vbCrLf
    s = s & "    d.SaleTotal   = Trim(txtSaleTotal.Text)" & vbCrLf
    s = s & "    d.PaymentDate = Trim(txtPayDate.Text)" & vbCrLf
    s = s & "    CollectSaleData = d" & vbCrLf
    s = s & "End Function" & vbCrLf
    GetFrmInputCode = s
End Function

