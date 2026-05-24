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
    RemoveComponent vbProj, "frmSales"

    Dim mMain As Object
    Set mMain = vbProj.VBComponents.Add(VBEXT_CT_STDMODULE)
    mMain.Name = "mMain"
    mMain.CodeModule.AddFromString GetMMainCode()

    BuildFrmSettings vbProj
    BuildFrmInput vbProj
    BuildFrmSales vbProj

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

    ' 右ペイン: 入力フォーム（全列対応）
    Dim labels As Variant, names As Variant, i As Integer
    labels = Array("仕入れ日(B1)", "回次(C1)", "車名/グレード(F1)", "年式/月(D1)", "車検(G1)", _
                   "評価点(H)", "車輌代(I)", "消費税(J)", "自税(K)", "リサイクル(L)", _
                   "落札料(M)", "合計(N)", "評価損(O)", "車輌番号(P1)", _
                   "仕入れ先(B2)", "出品番号(C2)", "色(D2)", "車台番号(F2)", _
                   "走行距離km(G2)", "所有者(P2)", "補足(T)")
    names  = Array("Date", "Session", "CarName", "Year", "Shaken", _
                   "Score", "Price", "Tax", "CarTax", "Recycle", _
                   "Fee", "Total", "Loss", "Plate", _
                   "Supplier", "LotNum", "Color", "Chassis", _
                   "Mileage", "Owner", "Memo")

    Dim bL As Single, bT As Single, rH As Single
    bL = 300 : bT = 28 : rH = 26
    For i = 0 To UBound(labels)
        AddCtl d, "Forms.Label.1",   "lbl" & names(i), bL,      bT + i * rH, 95, 18, CStr(labels(i))
        AddCtl d, "Forms.TextBox.1", "txt" & names(i), bL + 98, bT + i * rH, 110, 20
    Next i

    Dim frmH As Single
    frmH = bT + (UBound(labels) + 1) * rH + 50
    comp.Properties("Height") = frmH

    Dim btnTop As Single
    btnTop = bT + (UBound(labels) + 1) * rH + 6
    Dim btnImp As Object
    Set btnImp = AddCtl(d, "Forms.CommandButton.1", "btnImport", 300, btnTop, 100, 28, "取込")
    On Error Resume Next
    btnImp.Default = True
    On Error GoTo 0
    AddCtl d, "Forms.CommandButton.1", "btnCancel", 410, btnTop, 100, 28, "中止"

    comp.CodeModule.AddFromString GetFrmInputCode()
End Sub

Private Sub BuildFrmSales(vbProj As Object)
    Dim comp As Object
    Set comp = vbProj.VBComponents.Add(VBEXT_CT_MSFORM)
    comp.Name = "frmSales"
    comp.Properties("Caption") = "売上登録"
    comp.Properties("Width") = 320
    comp.Properties("Height") = 380

    Dim d As Object
    Set d = GetDesigner(comp)

    Dim labels As Variant, names As Variant, i As Integer
    labels = Array("シート名", "仕入番号(検索)", "名義変更(U)", "売上日(V1)", "売上先(V2)", _
                   "売上回次(W1)", "売上出品番号(W2)", "売上車輌代(X)", "売上消費税(Y)", _
                   "売上リサイクル(Z)", "売上合計(AA)", "入金日(AB)")
    names  = Array("SheetName", "CarNum", "Meigi", "SaleDate", "Buyer", _
                   "SaleSession", "SaleLot", "SalePrice", "SaleTax", _
                   "SaleRecycle", "SaleTotal", "PayDate")

    Dim bL As Single, bT As Single, rH As Single
    bL = 12 : bT = 12 : rH = 26
    For i = 0 To UBound(labels)
        AddCtl d, "Forms.Label.1",   "lbl" & names(i), bL,      bT + i * rH, 110, 18, CStr(labels(i))
        AddCtl d, "Forms.TextBox.1", "txt" & names(i), bL + 115, bT + i * rH, 150, 20
    Next i

    Dim btnTop As Single
    btnTop = bT + (UBound(labels) + 1) * rH + 6
    comp.Properties("Height") = btnTop + 50
    AddCtl d, "Forms.CommandButton.1", "btnOK",     bL + 60,  btnTop, 80, 26, "登録"
    AddCtl d, "Forms.CommandButton.1", "btnCancel", bL + 150, btnTop, 80, 26, "閉じる"

    comp.CodeModule.AddFromString GetFrmSalesCode()
End Sub

Private Function GetMMainCode() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public g_SheetName  As String" & vbCrLf
    s = s & "Public g_FolderPath As String" & vbCrLf
    s = s & "Public g_Cancelled  As Boolean" & vbCrLf
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
    s = s & "    CarNumber   As String   ' A:  仕入番号（検索キー）" & vbCrLf
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
    s = s & "' 仕入れ入力エントリ（画像フォルダから）" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Sub Main()" & vbCrLf
    s = s & "    g_SheetName  = ""27期""" & vbCrLf
    s = s & "    g_FolderPath = """"" & vbCrLf
    s = s & "    g_Cancelled  = False" & vbCrLf
    s = s & "    frmSettings.Show" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 売上入力エントリ" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Sub Main_Sales()" & vbCrLf
    s = s & "    g_SheetName = ""27期""" & vbCrLf
    s = s & "    frmSales.Show" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' A列を検索して連番の最大値+1を返す（例: 27-005 → 6）" & vbCrLf
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
    s = s & "                Dim num As Long" & vbCrLf
    s = s & "                num = CLng(parts(1))" & vbCrLf
    s = s & "                If num > maxNum Then maxNum = num" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    GetNextNumber = maxNum + 1" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' 次の書き込み先行番号を返す（Q列""着""を末尾として算出）" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Function GetNextWriteRow(ws As Worksheet) As Long" & vbCrLf
    s = s & "    Dim lastRow As Long, qVal As String" & vbCrLf
    s = s & "    lastRow = ws.Cells(ws.Rows.Count, 17).End(xlUp).Row" & vbCrLf
    s = s & "    If lastRow < 10 Then GetNextWriteRow = 10 : Exit Function" & vbCrLf
    s = s & "    qVal = Trim(CStr(ws.Cells(lastRow, 17).Value2))" & vbCrLf
    s = s & "    Select Case qVal" & vbCrLf
    s = s & "        Case ""着""  : GetNextWriteRow = lastRow + 1" & vbCrLf
    s = s & "        Case ""予定"" : GetNextWriteRow = lastRow + 2" & vbCrLf
    s = s & "        Case Else  : GetNextWriteRow = lastRow + 1" & vbCrLf
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
    s = s & "    ' 行1" & vbCrLf
    s = s & "    ws.Cells(r,  1).Value2 = data.CarNumber   ' A" & vbCrLf
    s = s & "    ws.Cells(r,  4).Value2 = data.YearMonth   ' D" & vbCrLf
    s = s & "    ws.Cells(r,  6).Value2 = data.CarName     ' F" & vbCrLf
    s = s & "    ws.Cells(r, 17).Value2 = ""予定""           ' Q" & vbCrLf
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
    s = s & "    ' 行2" & vbCrLf
    s = s & "    ws.Cells(r+1, 17).Value2 = ""着""           ' Q" & vbCrLf
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
    s = s & "' 売上データを既存行に書き込み（仕入番号で行を特定）" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "Sub WriteToSheetSales(ws As Worksheet, data As SaleData)" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    r = FindCarRow(ws, data.CarNumber)" & vbCrLf
    s = s & "    If r = 0 Then" & vbCrLf
    s = s & "        MsgBox ""仕入番号 "" & data.CarNumber & "" が見つかりません。"", vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
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
    s = s & "        dot  = InStrRev(fileName, ""."")" & vbCrLf
    s = s & "        base = Left(fileName, dot - 1)" & vbCrLf
    s = s & "        ext  = Mid(fileName, dot)" & vbCrLf
    s = s & "        cnt  = 1" & vbCrLf
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
    s = s & "Private m_Main    As Integer" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub SetupForm(files() As String, mainIdx As Integer)" & vbCrLf
    s = s & "    m_Files = files" & vbCrLf
    s = s & "    m_Count = UBound(files) - LBound(files) + 1" & vbCrLf
    s = s & "    m_Idx   = mainIdx" & vbCrLf
    s = s & "    m_Main  = mainIdx" & vbCrLf
    s = s & "    RefreshImage" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' === AI-OCR ここから（API課金のため無効化中）===" & vbCrLf
    s = s & "    ' Dim d As CarData" & vbCrLf
    s = s & "    ' d = CallClaudeOCR(m_Files(m_Main))" & vbCrLf
    s = s & "    ' txtDate.Text        = d.PurchaseDate" & vbCrLf
    s = s & "    ' txtSession.Text     = d.Session" & vbCrLf
    s = s & "    ' txtCarName.Text     = d.CarName" & vbCrLf
    s = s & "    ' txtYear.Text        = d.YearMonth" & vbCrLf
    s = s & "    ' txtShaken.Text      = d.Shaken" & vbCrLf
    s = s & "    ' txtColor.Text       = d.Color" & vbCrLf
    s = s & "    ' txtChassis.Text     = d.Chassis" & vbCrLf
    s = s & "    ' txtScore.Text       = d.Score" & vbCrLf
    s = s & "    ' txtMileage.Text     = d.Mileage" & vbCrLf
    s = s & "    ' txtPrice.Text       = d.Price" & vbCrLf
    s = s & "    ' txtTax.Text         = d.Tax" & vbCrLf
    s = s & "    ' txtCarTax.Text      = d.CarTax" & vbCrLf
    s = s & "    ' txtRecycle.Text     = d.Recycle" & vbCrLf
    s = s & "    ' txtFee.Text         = d.AuctionFee" & vbCrLf
    s = s & "    ' txtTotal.Text       = d.Total" & vbCrLf
    s = s & "    ' txtLoss.Text        = d.Loss" & vbCrLf
    s = s & "    ' txtSupplier.Text    = d.Supplier" & vbCrLf
    s = s & "    ' txtLotNum.Text      = d.LotNumber" & vbCrLf
    s = s & "    ' txtPlate.Text       = d.Plate" & vbCrLf
    s = s & "    ' txtOwner.Text       = d.Owner" & vbCrLf
    s = s & "    ' txtMemo.Text        = d.Memo" & vbCrLf
    s = s & "    ' === AI-OCR ここまで ===" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' --- デフォルト値（AIが有効になったら削除） ---" & vbCrLf
    s = s & "    txtDate.Text     = Format(Now, ""yyyy/m/d"")" & vbCrLf
    s = s & "    txtSession.Text  = """"" & vbCrLf
    s = s & "    txtCarName.Text  = ""スズキ クロスビー MZ""" & vbCrLf
    s = s & "    txtYear.Text     = ""2021/7""" & vbCrLf
    s = s & "    txtShaken.Text   = """"" & vbCrLf
    s = s & "    txtColor.Text    = ""ホワイトパール""" & vbCrLf
    s = s & "    txtChassis.Text  = ""MYN15S-100001""" & vbCrLf
    s = s & "    txtScore.Text    = ""4.5""" & vbCrLf
    s = s & "    txtMileage.Text  = ""35000""" & vbCrLf
    s = s & "    txtPrice.Text    = ""980000""" & vbCrLf
    s = s & "    txtTax.Text      = ""98000""" & vbCrLf
    s = s & "    txtCarTax.Text   = ""35400""" & vbCrLf
    s = s & "    txtRecycle.Text  = ""12000""" & vbCrLf
    s = s & "    txtFee.Text      = ""32000""" & vbCrLf
    s = s & "    txtTotal.Text    = """"" & vbCrLf
    s = s & "    txtLoss.Text     = """"" & vbCrLf
    s = s & "    txtSupplier.Text = ""USS大阪""" & vbCrLf
    s = s & "    txtLotNum.Text   = ""12345""" & vbCrLf
    s = s & "    txtPlate.Text    = """"" & vbCrLf
    s = s & "    txtOwner.Text    = """"" & vbCrLf
    s = s & "    txtMemo.Text     = """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub RefreshImage()" & vbCrLf
    s = s & "    If m_Count = 0 Then Exit Sub" & vbCrLf
    s = s & "    Dim prefix As String" & vbCrLf
    s = s & "    If m_Idx = m_Main Then prefix = ""★ "" Else prefix = ""　""" & vbCrLf
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
    s = s & "Public Function GetData() As CarData" & vbCrLf
    s = s & "    Dim data As CarData" & vbCrLf
    s = s & "    data.PurchaseDate = Trim(txtDate.Text)        ' B1: 仕入れ日" & vbCrLf
    s = s & "    data.Session      = Trim(txtSession.Text)     ' C1: 回次" & vbCrLf
    s = s & "    data.CarName      = Trim(txtCarName.Text)     ' F1: 車名" & vbCrLf
    s = s & "    data.YearMonth    = Trim(txtYear.Text)        ' D1: 年式/月" & vbCrLf
    s = s & "    data.Shaken       = Trim(txtShaken.Text)      ' G1: 車検" & vbCrLf
    s = s & "    data.Score        = Trim(txtScore.Text)       ' H:  評価点" & vbCrLf
    s = s & "    data.Price        = Trim(txtPrice.Text)       ' I:  車輌代" & vbCrLf
    s = s & "    data.Tax          = Trim(txtTax.Text)         ' J:  消費税" & vbCrLf
    s = s & "    data.CarTax       = Trim(txtCarTax.Text)      ' K:  自税" & vbCrLf
    s = s & "    data.Recycle      = Trim(txtRecycle.Text)     ' L:  リサイクル" & vbCrLf
    s = s & "    data.AuctionFee   = Trim(txtFee.Text)         ' M:  落札料" & vbCrLf
    s = s & "    data.Total        = Trim(txtTotal.Text)       ' N:  合計" & vbCrLf
    s = s & "    data.Loss         = Trim(txtLoss.Text)        ' O:  評価損" & vbCrLf
    s = s & "    data.Plate        = Trim(txtPlate.Text)       ' P1: 車輌番号" & vbCrLf
    s = s & "    data.Supplier     = Trim(txtSupplier.Text)    ' B2: 仕入れ先" & vbCrLf
    s = s & "    data.LotNumber    = Trim(txtLotNum.Text)      ' C2: 出品番号" & vbCrLf
    s = s & "    data.Color        = Trim(txtColor.Text)       ' D2: 色" & vbCrLf
    s = s & "    data.Chassis      = Trim(txtChassis.Text)     ' F2: 車台番号" & vbCrLf
    s = s & "    data.Mileage      = Trim(txtMileage.Text)     ' G2: 走行距離" & vbCrLf
    s = s & "    data.Owner        = Trim(txtOwner.Text)       ' P2: 所有者" & vbCrLf
    s = s & "    data.Memo         = Trim(txtMemo.Text)        ' T:  補足" & vbCrLf
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

Private Function GetFrmSalesCode() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    txtSheetName.Text = g_SheetName" & vbCrLf
    s = s & "    txtSaleDate.Text  = Format(Now, ""yyyy/m/d"")" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnOK_Click()" & vbCrLf
    s = s & "    If Trim(txtCarNum.Text) = """" Then" & vbCrLf
    s = s & "        MsgBox ""仕入番号を入力してください。"", vbExclamation" & vbCrLf
    s = s & "        txtCarNum.SetFocus : Exit Sub" & vbCrLf
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
    s = s & "    Dim data As SaleData" & vbCrLf
    s = s & "    data.CarNumber   = Trim(txtCarNum.Text)" & vbCrLf
    s = s & "    data.Meigi       = Trim(txtMeigi.Text)" & vbCrLf
    s = s & "    data.SaleDate    = Trim(txtSaleDate.Text)" & vbCrLf
    s = s & "    data.Buyer       = Trim(txtBuyer.Text)" & vbCrLf
    s = s & "    data.SaleSession = Trim(txtSaleSession.Text)" & vbCrLf
    s = s & "    data.SaleLot     = Trim(txtSaleLot.Text)" & vbCrLf
    s = s & "    data.SalePrice   = Trim(txtSalePrice.Text)" & vbCrLf
    s = s & "    data.SaleTax     = Trim(txtSaleTax.Text)" & vbCrLf
    s = s & "    data.SaleRecycle = Trim(txtSaleRecycle.Text)" & vbCrLf
    s = s & "    data.SaleTotal   = Trim(txtSaleTotal.Text)" & vbCrLf
    s = s & "    data.PaymentDate = Trim(txtPayDate.Text)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    WriteToSheetSales ws, data" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    MsgBox data.CarNumber & "" の売上を登録しました。"", vbInformation" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnCancel_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    GetFrmSalesCode = s
End Function

