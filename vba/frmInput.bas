Attribute VB_Name = "frmInput"
Option Explicit

' ── モジュールレベル変数 ─────────────────────────────
Private m_Files()    As String
Private m_Count      As Integer
Private m_Idx        As Integer
Private m_Processing As Boolean

' ── ボタンイベント用 WithEvents ──────────────────────
Private WithEvents m_btnBrowse  As MSForms.CommandButton
Private WithEvents m_btnImport  As MSForms.CommandButton
Private WithEvents m_btnCancel  As MSForms.CommandButton
Private WithEvents m_btnPrev    As MSForms.CommandButton
Private WithEvents m_btnNext    As MSForms.CommandButton

' ============================================================
' 初期化: コントロール生成 → 初期値セット
' ============================================================
Private Sub UserForm_Initialize()
    Me.Caption = "在庫管理"
    Me.Width   = 650
    Me.Height  = 720
    Me.ScrollBars   = 2
    Me.ScrollHeight = 1000

    CreateControls

    m_Processing = False
    CtlText("txtSheetName") = g_SheetName
    m_btnImport.Caption = "開始"
End Sub

' ============================================================
' 全コントロールをランタイムで生成（Designer 不要）
' ============================================================
Private Sub CreateControls()
    ' ── 設定行 ──────────────────────────────────────────
    Lbl "lblSheetName",  8,  8, 70, 18, "シート名"
    Txt "txtSheetName", 80,  6, 80, 20
    Lbl "lblFolder",   170,  8, 60, 18, "フォルダ"
    Txt "txtFolder",   235,  6, 255, 20
    Set m_btnBrowse = Btn("btnBrowse", 495, 4, 48, 22, "参照...")

    ' ── 左ペイン: 画像 ──────────────────────────────────
    Lbl "lblProgress",  8, 34, 290, 18, "画像なし"
    Dim img As Object
    Set img = Me.Controls.Add("Forms.Image.1", "imgPreview", True)
    img.Left = 8 : img.Top = 55 : img.Width = 290 : img.Height = 360
    On Error Resume Next
    img.PictureSizeMode = 3
    img.BorderStyle = 1
    On Error GoTo 0
    Lbl "lblFileName",  8, 420, 290, 18, ""
    Set m_btnPrev = Btn("btnPrev",  8, 442, 42, 22, "←")
    Set m_btnNext = Btn("btnNext", 55, 442, 42, 22, "→")

    ' ── 右ペイン: 全フィールド ──────────────────────────
    Dim bL As Single, bT As Single, rH As Single
    bL = 312 : bT = 34 : rH = 24

    Dim bLbl As Variant, bName As Variant, i As Integer
    bLbl  = Array("仕入れ日", "仕入れ先", "回次", "出品番号", _
                  "年式", "色", "車体番号", "車台番号", _
                  "車検", "距離km", "評価点", "車輌代", "消費税", "自税", _
                  "リサイクル", "落札料", "合計", "評価損", _
                  "車輌番号", "所有者", "付属品", "補足")
    bName = Array("Date", "Supplier", "Session", "LotNum", _
                  "Year", "Color", "CarName", "Chassis", _
                  "Shaken", "Mileage", "Score", "Price", "Tax", "CarTax", _
                  "Recycle", "Fee", "Total", "Loss", _
                  "Plate", "Owner", "Accessories", "Memo")

    For i = 0 To UBound(bLbl)
        Lbl "lbl" & bName(i), bL,       bT + i * rH, 105, 18, CStr(bLbl(i))
        Txt "txt" & bName(i), bL + 108, bT + i * rH, 112, 20
    Next i

    Dim sepT As Single : sepT = bT + (UBound(bLbl) + 1) * rH
    Lbl "lblSaleSep", bL, sepT, 220, 18, "── 売上 ──────────"

    Dim sLbl As Variant, sName As Variant
    sLbl  = Array("名義変更", "売上日", "売上先", "回次", _
                  "出品番号", "車輌代", "消費税", _
                  "リサイクル", "合計", "入金日")
    sName = Array("Meigi", "SaleDate", "Buyer", "SaleSession", _
                  "SaleLot", "SalePrice", "SaleTax", _
                  "SaleRecycle", "SaleTotal", "PayDate")

    Dim salT As Single : salT = sepT + rH
    For i = 0 To UBound(sLbl)
        Lbl "lbl" & sName(i), bL,       salT + i * rH, 105, 18, CStr(sLbl(i))
        Txt "txt" & sName(i), bL + 108, salT + i * rH, 112, 20
    Next i

    Dim btnT As Single : btnT = salT + (UBound(sLbl) + 1) * rH + 8
    Set m_btnImport = Btn("btnImport", bL,       btnT, 100, 28, "開始")
    Set m_btnCancel = Btn("btnCancel", bL + 110, btnT, 100, 28, "閉じる")
End Sub

' ── コントロール生成ヘルパー ─────────────────────────
Private Sub Lbl(n As String, L As Single, T As Single, W As Single, H As Single, cap As String)
    Dim c As Object
    Set c = Me.Controls.Add("Forms.Label.1", n, True)
    c.Left = L : c.Top = T : c.Width = W : c.Height = H : c.Caption = cap
End Sub

Private Sub Txt(n As String, L As Single, T As Single, W As Single, H As Single)
    Dim c As Object
    Set c = Me.Controls.Add("Forms.TextBox.1", n, True)
    c.Left = L : c.Top = T : c.Width = W : c.Height = H
End Sub

Private Function Btn(n As String, L As Single, T As Single, W As Single, H As Single, cap As String) As MSForms.CommandButton
    Dim c As MSForms.CommandButton
    Set c = Me.Controls.Add("Forms.CommandButton.1", n, True)
    c.Left = L : c.Top = T : c.Width = W : c.Height = H : c.Caption = cap
    Set Btn = c
End Function

Private Property Get CtlText(n As String) As String
    CtlText = Me.Controls(n).Text
End Property

Private Property Let CtlText(n As String, val As String)
    Me.Controls(n).Text = val
End Property

' ============================================================
' ボタンイベント
' ============================================================
Private Sub m_btnBrowse_Click()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "画像フォルダを選択してください"
        .AllowMultiSelect = False
        If .Show Then CtlText("txtFolder") = .SelectedItems(1)
    End With
End Sub

Private Sub m_btnImport_Click()
    If Not m_Processing Then StartProcessing Else ProcessCurrentImage
End Sub

Private Sub m_btnCancel_Click()
    If MsgBox("処理を中止しますか？", vbQuestion + vbYesNo) = vbYes Then Unload Me
End Sub

Private Sub m_btnPrev_Click()
    If m_Idx > 0 Then m_Idx = m_Idx - 1 : RefreshImage
End Sub

Private Sub m_btnNext_Click()
    If m_Idx < m_Count - 1 Then m_Idx = m_Idx + 1 : RefreshImage
End Sub

' ============================================================
' 開始処理
' ============================================================
Private Sub StartProcessing()
    If Trim(CtlText("txtSheetName")) = "" Then
        MsgBox "シート名を入力してください。", vbExclamation : Exit Sub
    End If
    If Trim(CtlText("txtFolder")) = "" Then
        MsgBox "フォルダを選択してください。", vbExclamation : Exit Sub
    End If
    If Dir(Trim(CtlText("txtFolder")), vbDirectory) = "" Then
        MsgBox "フォルダが見つかりません。", vbExclamation : Exit Sub
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Trim(CtlText("txtSheetName")))
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "シート「" & CtlText("txtSheetName") & "」が見つかりません。", vbExclamation : Exit Sub
    End If

    g_SheetName  = Trim(CtlText("txtSheetName"))
    g_FolderPath = Trim(CtlText("txtFolder"))

    Dim exts(3) As String
    exts(0) = "*.jpg" : exts(1) = "*.jpeg" : exts(2) = "*.png" : exts(3) = "*.bmp"
    Dim col As New Collection, i As Integer, f As String
    For i = 0 To 3
        f = Dir(g_FolderPath & "\" & exts(i))
        Do While f <> ""
            col.Add g_FolderPath & "\" & f
            f = Dir()
        Loop
    Next i

    If col.Count = 0 Then
        MsgBox "画像ファイル（JPG/PNG/BMP）が見つかりません。", vbInformation : Exit Sub
    End If

    ReDim m_Files(0 To col.Count - 1)
    For i = 0 To col.Count - 1
        m_Files(i) = col(i + 1)
    Next i
    m_Count = col.Count : m_Idx = 0 : m_Processing = True

    LoadDefaults
    RefreshImage
    m_btnImport.Caption = "取込"
End Sub

' ============================================================
' 取込処理
' ============================================================
Private Sub ProcessCurrentImage()
    If Trim(CtlText("txtCarName")) = "" Then
        MsgBox "車名を入力してください。", vbExclamation : Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(g_SheetName)

    Dim data As CarData
    data = CollectCarData()
    data.CarNumber = GetPeriodNumber(ws) & "-" & Format(GetNextNumber(ws), "000")
    WriteToSheet ws, data

    If Trim(CtlText("txtSaleDate")) <> "" Then
        Dim sd As SaleData
        sd = CollectSaleData()
        sd.CarNumber = data.CarNumber
        WriteToSheetSales ws, sd
    End If

    MoveToProcessed m_Files(m_Idx)
    m_Idx = m_Idx + 1

    If m_Idx >= m_Count Then
        MsgBox m_Count & " 件の処理が完了しました。", vbInformation
        Unload Me
    Else
        LoadDefaults
        RefreshImage
    End If
End Sub

' ============================================================
' デフォルト値（AI-OCR有効化後はここを置き換える）
' ============================================================
Private Sub LoadDefaults()
    ' === AI-OCR ここから（API課金のため無効化中）===
    ' Dim d As CarData : d = CallClaudeOCR(m_Files(m_Idx))
    ' CtlText("txtCarName") = d.CarName ... 以下同様
    ' === AI-OCR ここまで ===

    CtlText("txtDate")        = Format(Now, "yyyy/m/d")
    CtlText("txtSession")     = ""
    CtlText("txtCarName")     = "スズキ クロスビー MZ"
    CtlText("txtYear")        = "2021/7"
    CtlText("txtShaken")      = ""
    CtlText("txtScore")       = "4.5"
    CtlText("txtPrice")       = "980000"
    CtlText("txtTax")         = "98000"
    CtlText("txtCarTax")      = "35400"
    CtlText("txtRecycle")     = "12000"
    CtlText("txtFee")         = "32000"
    CtlText("txtTotal")       = ""
    CtlText("txtLoss")        = ""
    CtlText("txtPlate")       = ""
    CtlText("txtSupplier")    = "USS大阪"
    CtlText("txtLotNum")      = "12345"
    CtlText("txtColor")       = "ホワイトパール"
    CtlText("txtChassis")     = "MYN15S-100001"
    CtlText("txtMileage")     = "35000"
    CtlText("txtOwner")       = ""
    CtlText("txtAccessories") = ""
    CtlText("txtMemo")        = ""
    CtlText("txtMeigi")       = ""
    CtlText("txtSaleDate")    = ""
    CtlText("txtBuyer")       = ""
    CtlText("txtSaleSession") = ""
    CtlText("txtSaleLot")     = ""
    CtlText("txtSalePrice")   = ""
    CtlText("txtSaleTax")     = ""
    CtlText("txtSaleRecycle") = ""
    CtlText("txtSaleTotal")   = ""
    CtlText("txtPayDate")     = ""
End Sub

Private Sub RefreshImage()
    If m_Count = 0 Then Exit Sub
    Me.Controls("lblProgress").Caption = (m_Idx + 1) & " / " & m_Count
    Me.Controls("lblFileName").Caption = Mid(m_Files(m_Idx), InStrRev(m_Files(m_Idx), "\") + 1)
    On Error Resume Next
    Me.Controls("imgPreview").Picture = LoadPicture(m_Files(m_Idx))
    On Error GoTo 0
End Sub

' ============================================================
' データ収集
' ============================================================
Private Function CollectCarData() As CarData
    Dim d As CarData
    d.PurchaseDate = Trim(CtlText("txtDate"))
    d.Session      = Trim(CtlText("txtSession"))
    d.CarName      = Trim(CtlText("txtCarName"))
    d.YearMonth    = Trim(CtlText("txtYear"))
    d.Shaken       = Trim(CtlText("txtShaken"))
    d.Score        = Trim(CtlText("txtScore"))
    d.Price        = Trim(CtlText("txtPrice"))
    d.Tax          = Trim(CtlText("txtTax"))
    d.CarTax       = Trim(CtlText("txtCarTax"))
    d.Recycle      = Trim(CtlText("txtRecycle"))
    d.AuctionFee   = Trim(CtlText("txtFee"))
    d.Total        = Trim(CtlText("txtTotal"))
    d.Loss         = Trim(CtlText("txtLoss"))
    d.Plate        = Trim(CtlText("txtPlate"))
    d.Supplier     = Trim(CtlText("txtSupplier"))
    d.LotNumber    = Trim(CtlText("txtLotNum"))
    d.Color        = Trim(CtlText("txtColor"))
    d.Chassis      = Trim(CtlText("txtChassis"))
    d.Mileage      = Trim(CtlText("txtMileage"))
    d.Owner        = Trim(CtlText("txtOwner"))
    d.Accessories  = Trim(CtlText("txtAccessories"))
    d.Memo         = Trim(CtlText("txtMemo"))
    CollectCarData = d
End Function

Private Function CollectSaleData() As SaleData
    Dim d As SaleData
    d.Meigi       = Trim(CtlText("txtMeigi"))
    d.SaleDate    = Trim(CtlText("txtSaleDate"))
    d.Buyer       = Trim(CtlText("txtBuyer"))
    d.SaleSession = Trim(CtlText("txtSaleSession"))
    d.SaleLot     = Trim(CtlText("txtSaleLot"))
    d.SalePrice   = Trim(CtlText("txtSalePrice"))
    d.SaleTax     = Trim(CtlText("txtSaleTax"))
    d.SaleRecycle = Trim(CtlText("txtSaleRecycle"))
    d.SaleTotal   = Trim(CtlText("txtSaleTotal"))
    d.PaymentDate = Trim(CtlText("txtPayDate"))
    CollectSaleData = d
End Function
