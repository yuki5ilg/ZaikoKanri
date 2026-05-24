Attribute VB_Name = "frmInput"
Option Explicit

Private m_Files() As String
Private m_Count   As Integer
Private m_Idx     As Integer
Private m_Main    As Integer

Public Sub SetupForm(files() As String, mainIdx As Integer)
    m_Files = files
    m_Count = UBound(files) - LBound(files) + 1
    m_Idx   = mainIdx
    m_Main  = mainIdx
    RefreshImage

    ' === AI-OCR ここから（API課金のため無効化中）===
    ' Dim d As CarData
    ' d = CallClaudeOCR(m_Files(m_Main))
    ' txtDate.Text        = d.PurchaseDate
    ' txtSession.Text     = d.Session
    ' txtCarName.Text     = d.CarName
    ' txtYear.Text        = d.YearMonth
    ' txtShaken.Text      = d.Shaken
    ' txtColor.Text       = d.Color
    ' txtChassis.Text     = d.Chassis
    ' txtScore.Text       = d.Score
    ' txtMileage.Text     = d.Mileage
    ' txtPrice.Text       = d.Price
    ' txtTax.Text         = d.Tax
    ' txtCarTax.Text      = d.CarTax
    ' txtRecycle.Text     = d.Recycle
    ' txtFee.Text         = d.AuctionFee
    ' txtTotal.Text       = d.Total
    ' txtLoss.Text        = d.Loss
    ' txtSupplier.Text    = d.Supplier
    ' txtLotNum.Text      = d.LotNumber
    ' txtPlate.Text       = d.Plate
    ' txtOwner.Text       = d.Owner
    ' txtMemo.Text        = d.Memo
    ' === AI-OCR ここまで ===

    ' --- デフォルト値（AIが有効になったら削除） ---
    txtDate.Text     = Format(Now, "yyyy/m/d")
    txtSession.Text  = ""
    txtCarName.Text  = "スズキ クロスビー MZ"
    txtYear.Text     = "2021/7"
    txtShaken.Text   = ""
    txtColor.Text    = "ホワイトパール"
    txtChassis.Text  = "MYN15S-100001"
    txtScore.Text    = "4.5"
    txtMileage.Text  = "35000"
    txtPrice.Text    = "980000"
    txtTax.Text      = "98000"
    txtCarTax.Text   = "35400"
    txtRecycle.Text  = "12000"
    txtFee.Text      = "32000"
    txtTotal.Text    = ""
    txtLoss.Text     = ""
    txtSupplier.Text = "USS大阪"
    txtLotNum.Text   = "12345"
    txtPlate.Text    = ""
    txtOwner.Text    = ""
    txtMemo.Text     = ""
End Sub

Private Sub RefreshImage()
    If m_Count = 0 Then Exit Sub
    Dim prefix As String
    If m_Idx = m_Main Then prefix = "★ " Else prefix = "　"
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

Public Function GetData() As CarData
    Dim data As CarData
    data.PurchaseDate = Trim(txtDate.Text)        ' B1: 仕入れ日
    data.Session      = Trim(txtSession.Text)     ' C1: 回次
    data.CarName      = Trim(txtCarName.Text)     ' F1: 車名
    data.YearMonth    = Trim(txtYear.Text)        ' D1: 年式/月
    data.Shaken       = Trim(txtShaken.Text)      ' G1: 車検
    data.Score        = Trim(txtScore.Text)       ' H:  評価点
    data.Price        = Trim(txtPrice.Text)       ' I:  車輌代
    data.Tax          = Trim(txtTax.Text)         ' J:  消費税
    data.CarTax       = Trim(txtCarTax.Text)      ' K:  自税
    data.Recycle      = Trim(txtRecycle.Text)     ' L:  リサイクル
    data.AuctionFee   = Trim(txtFee.Text)         ' M:  落札料
    data.Total        = Trim(txtTotal.Text)       ' N:  合計
    data.Loss         = Trim(txtLoss.Text)        ' O:  評価損
    data.Plate        = Trim(txtPlate.Text)       ' P1: 車輌番号
    data.Supplier     = Trim(txtSupplier.Text)    ' B2: 仕入れ先
    data.LotNumber    = Trim(txtLotNum.Text)      ' C2: 出品番号
    data.Color        = Trim(txtColor.Text)       ' D2: 色
    data.Chassis      = Trim(txtChassis.Text)     ' F2: 車台番号
    data.Mileage      = Trim(txtMileage.Text)     ' G2: 走行距離
    data.Owner        = Trim(txtOwner.Text)       ' P2: 所有者
    data.Memo         = Trim(txtMemo.Text)        ' T:  補足
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
