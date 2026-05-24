Attribute VB_Name = "frmInput"
Option Explicit

Private m_Files() As String
Private m_Count   As Integer
Private m_Idx     As Integer
Private m_Main    As Integer   ' 取込対象のインデックス（ナビで変わらない）

' ============================================================
' 初期化: 全画像配列 + 今回の取込対象インデックス
' ============================================================
Public Sub SetupForm(files() As String, mainIdx As Integer)
    m_Files = files
    m_Count = UBound(files) - LBound(files) + 1
    m_Idx   = mainIdx
    m_Main  = mainIdx

    RefreshImage

    ' === AI-OCR ここから（API課金のため無効化中）===
    ' Dim d As CarData
    ' d = CallClaudeOCR(m_Files(m_Main))
    ' txtCarName.Text  = d.CarName   ' 車名（グレード含む）
    ' txtYear.Text     = d.YearMonth ' 年式/月（例: 2021/7）
    ' txtColor.Text    = d.Color
    ' txtChassis.Text  = d.Chassis
    ' txtScore.Text    = d.Score
    ' txtMileage.Text  = d.Mileage
    ' txtPrice.Text    = d.Price
    ' txtTax.Text      = d.Tax
    ' txtCarTax.Text   = d.CarTax
    ' txtRecycle.Text  = d.Recycle
    ' txtFee.Text      = d.AuctionFee
    ' txtVenue.Text    = d.Venue
    ' txtLotNum.Text   = d.LotNumber
    ' === AI-OCR ここまで ===

    ' --- デフォルト値（AIが有効になったら削除） ---
    txtCarName.Text  = "スズキ クロスビー MZ"
    txtYear.Text     = "2021/7"
    txtColor.Text    = "ホワイトパール"
    txtChassis.Text  = "MYN15S-100001"
    txtScore.Text    = "4.5"
    txtMileage.Text  = "35000"
    txtPrice.Text    = "980000"
    txtTax.Text      = "98000"
    txtCarTax.Text   = "35400"
    txtRecycle.Text  = "12000"
    txtFee.Text      = "32000"
    txtVenue.Text    = "USS大阪"
    txtLotNum.Text   = "12345"
End Sub

' ============================================================
' 画像表示更新（ナビ用。m_Idxが変わるだけ）
' ============================================================
Private Sub RefreshImage()
    If m_Count = 0 Then Exit Sub
    Dim prefix As String
    If m_Idx = m_Main Then
        prefix = "★ "   ' 取込対象
    Else
        prefix = "　"
    End If
    lblProgress.Caption = prefix & (m_Idx + 1) & " / " & m_Count
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

' ============================================================
' 入力値を CarData で返す（グレード・月は統合済みのため不要）
' ============================================================
Public Function GetData() As CarData
    Dim data As CarData
    data.CarName     = Trim(txtCarName.Text)   ' 車名（グレード含む）
    data.YearMonth   = Trim(txtYear.Text)      ' 年式/月（例: 2021/7）
    data.Score       = Trim(txtScore.Text)
    data.Price       = Trim(txtPrice.Text)
    data.Tax         = Trim(txtTax.Text)
    data.CarTax      = Trim(txtCarTax.Text)
    data.Recycle     = Trim(txtRecycle.Text)
    data.AuctionFee  = Trim(txtFee.Text)
    data.Venue       = Trim(txtVenue.Text)
    data.LotNumber   = Trim(txtLotNum.Text)
    data.Color       = Trim(txtColor.Text)
    data.Chassis     = Trim(txtChassis.Text)
    data.Mileage     = Trim(txtMileage.Text)
    GetData = data
End Function

Private Sub btnImport_Click()
    If Trim(txtCarName.Text) = "" Then
        MsgBox "車名を入力してください。", vbExclamation
        txtCarName.SetFocus
        Exit Sub
    End If
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    If MsgBox("処理を中止しますか？", vbQuestion + vbYesNo) = vbYes Then
        g_Cancelled = True
        Me.Hide
    End If
End Sub
