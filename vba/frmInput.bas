Attribute VB_Name = "frmInput"
Option Explicit

Private m_Files()    As String
Private m_Count      As Integer
Private m_Idx        As Integer
Private m_Processing As Boolean

' ============================================================
' 初期化
' ============================================================
Private Sub UserForm_Initialize()
    txtSheetName.Text = g_SheetName
    m_Processing      = False
    btnImport.Caption = "開始"
End Sub

' ============================================================
' フォルダ参照
' ============================================================
Private Sub btnBrowse_Click()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "画像フォルダを選択してください"
        .AllowMultiSelect = False
        If .Show Then txtFolder.Text = .SelectedItems(1)
    End With
End Sub

' ============================================================
' 取込ボタン（開始 / 取込 で挙動切り替え）
' ============================================================
Private Sub btnImport_Click()
    If Not m_Processing Then
        StartProcessing
    Else
        ProcessCurrentImage
    End If
End Sub

' ============================================================
' 開始処理: バリデーション → 画像収集 → 1枚目表示
' ============================================================
Private Sub StartProcessing()
    If Trim(txtSheetName.Text) = "" Then
        MsgBox "シート名を入力してください。", vbExclamation
        txtSheetName.SetFocus : Exit Sub
    End If
    If Trim(txtFolder.Text) = "" Then
        MsgBox "フォルダを選択してください。", vbExclamation : Exit Sub
    End If
    If Dir(Trim(txtFolder.Text), vbDirectory) = "" Then
        MsgBox "フォルダが見つかりません。", vbExclamation : Exit Sub
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Trim(txtSheetName.Text))
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "シート「" & txtSheetName.Text & "」が見つかりません。", vbExclamation : Exit Sub
    End If

    g_SheetName  = Trim(txtSheetName.Text)
    g_FolderPath = Trim(txtFolder.Text)

    Dim exts(3) As String
    exts(0) = "*.jpg" : exts(1) = "*.jpeg" : exts(2) = "*.png" : exts(3) = "*.bmp"

    Dim col As New Collection
    Dim i As Integer, f As String
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
    m_Count      = col.Count
    m_Idx        = 0
    m_Processing = True

    LoadDefaults
    RefreshImage
    btnImport.Caption = "取込"
End Sub

' ============================================================
' 取込処理: 書き込み → 次の画像 or 完了
' ============================================================
Private Sub ProcessCurrentImage()
    If Trim(txtCarName.Text) = "" Then
        MsgBox "車名を入力してください。", vbExclamation
        txtCarName.SetFocus : Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(g_SheetName)

    Dim data As CarData
    data           = CollectCarData()
    data.CarNumber = GetPeriodNumber(ws) & "-" & Format(GetNextNumber(ws), "000")
    WriteToSheet ws, data

    If Trim(txtSaleDate.Text) <> "" Then
        Dim sdata As SaleData
        sdata           = CollectSaleData()
        sdata.CarNumber = data.CarNumber
        WriteToSheetSales ws, sdata
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
' デフォルト値セット（AI-OCR有効化後はここを置き換える）
' ============================================================
Private Sub LoadDefaults()
    ' === AI-OCR ここから（API課金のため無効化中）===
    ' Dim d As CarData
    ' d = CallClaudeOCR(m_Files(m_Idx))
    ' txtDate.Text        = d.PurchaseDate
    ' txtSession.Text     = d.Session
    ' txtCarName.Text     = d.CarName
    ' ...
    ' === AI-OCR ここまで ===

    txtDate.Text     = Format(Now, "yyyy/m/d")
    txtSession.Text  = ""
    txtCarName.Text  = "スズキ クロスビー MZ"
    txtYear.Text     = "2021/7"
    txtShaken.Text   = ""
    txtScore.Text    = "4.5"
    txtPrice.Text    = "980000"
    txtTax.Text      = "98000"
    txtCarTax.Text   = "35400"
    txtRecycle.Text  = "12000"
    txtFee.Text      = "32000"
    txtTotal.Text    = ""
    txtLoss.Text     = ""
    txtPlate.Text    = ""
    txtSupplier.Text = "USS大阪"
    txtLotNum.Text   = "12345"
    txtColor.Text    = "ホワイトパール"
    txtChassis.Text  = "MYN15S-100001"
    txtMileage.Text  = "35000"
    txtOwner.Text    = ""
    txtMemo.Text     = ""
    txtMeigi.Text    = ""
    txtSaleDate.Text = ""
    txtBuyer.Text    = ""
    txtSaleSession.Text  = ""
    txtSaleLot.Text      = ""
    txtSalePrice.Text    = ""
    txtSaleTax.Text      = ""
    txtSaleRecycle.Text  = ""
    txtSaleTotal.Text    = ""
    txtPayDate.Text      = ""
End Sub

' ============================================================
' 画像表示更新
' ============================================================
Private Sub RefreshImage()
    If m_Count = 0 Then Exit Sub
    lblProgress.Caption = (m_Idx + 1) & " / " & m_Count
    lblFileName.Caption = Mid(m_Files(m_Idx), InStrRev(m_Files(m_Idx), "\") + 1)
    On Error Resume Next
    imgPreview.Picture = LoadPicture(m_Files(m_Idx))
    On Error GoTo 0
End Sub

Private Sub btnPrev_Click()
    If m_Idx > 0 Then m_Idx = m_Idx - 1 : RefreshImage
End Sub

Private Sub btnNext_Click()
    If m_Idx < m_Count - 1 Then m_Idx = m_Idx + 1 : RefreshImage
End Sub

Private Sub btnCancel_Click()
    If MsgBox("処理を中止しますか？", vbQuestion + vbYesNo) = vbYes Then Unload Me
End Sub

' ============================================================
' フォームから CarData を収集
' ============================================================
Private Function CollectCarData() As CarData
    Dim d As CarData
    d.PurchaseDate = Trim(txtDate.Text)
    d.Session      = Trim(txtSession.Text)
    d.CarName      = Trim(txtCarName.Text)
    d.YearMonth    = Trim(txtYear.Text)
    d.Shaken       = Trim(txtShaken.Text)
    d.Score        = Trim(txtScore.Text)
    d.Price        = Trim(txtPrice.Text)
    d.Tax          = Trim(txtTax.Text)
    d.CarTax       = Trim(txtCarTax.Text)
    d.Recycle      = Trim(txtRecycle.Text)
    d.AuctionFee   = Trim(txtFee.Text)
    d.Total        = Trim(txtTotal.Text)
    d.Loss         = Trim(txtLoss.Text)
    d.Plate        = Trim(txtPlate.Text)
    d.Supplier     = Trim(txtSupplier.Text)
    d.LotNumber    = Trim(txtLotNum.Text)
    d.Color        = Trim(txtColor.Text)
    d.Chassis      = Trim(txtChassis.Text)
    d.Mileage      = Trim(txtMileage.Text)
    d.Owner        = Trim(txtOwner.Text)
    d.Memo         = Trim(txtMemo.Text)
    CollectCarData = d
End Function

' ============================================================
' フォームから SaleData を収集
' ============================================================
Private Function CollectSaleData() As SaleData
    Dim d As SaleData
    d.Meigi       = Trim(txtMeigi.Text)
    d.SaleDate    = Trim(txtSaleDate.Text)
    d.Buyer       = Trim(txtBuyer.Text)
    d.SaleSession = Trim(txtSaleSession.Text)
    d.SaleLot     = Trim(txtSaleLot.Text)
    d.SalePrice   = Trim(txtSalePrice.Text)
    d.SaleTax     = Trim(txtSaleTax.Text)
    d.SaleRecycle = Trim(txtSaleRecycle.Text)
    d.SaleTotal   = Trim(txtSaleTotal.Text)
    d.PaymentDate = Trim(txtPayDate.Text)
    CollectSaleData = d
End Function
