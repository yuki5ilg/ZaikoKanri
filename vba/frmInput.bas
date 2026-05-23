Attribute VB_Name = "frmInput"
Option Explicit

' ============================================================
' frmInput - 画像プレビュー＋入力フォーム
'
' コントロール一覧（VBAデザイナーで作成）:
'
'   【左ペイン - 画像エリア】
'   Label        lblProgress     Caption="1 / 5"  （右上）
'   Image        imgPreview      PictureSizeMode=3（Zoom）
'   Label        lblFileName     Caption=""
'
'   【右ペイン - 入力フォーム】
'   ラベル/テキストボックスのペア（各行）:
'     lblCarName    / txtCarName    「車名」
'     lblGrade      / txtGrade      「グレード」
'     lblYear       / txtYear       「年式」
'     lblMonth      / txtMonth      「月」
'     lblColor      / txtColor      「色」
'     lblChassis    / txtChassis    「車台番号」
'     lblScore      / txtScore      「評価点」
'     lblMileage    / txtMileage    「走行距離(km)」
'     lblPrice      / txtPrice      「落札価格(円)」
'     lblTax        / txtTax        「消費税(円)」
'     lblCarTax     / txtCarTax     「自動車税(円)」
'     lblRecycle    / txtRecycle    「リサイクル料(円)」
'     lblFee        / txtFee        「落札手数料(円)」
'     lblVenue      / txtVenue      「オークション会場」
'     lblLotNum     / txtLotNum     「出品番号」
'
'   【ボタン】
'   CommandButton btnImport   Caption="取込"  （Default=True）
'   CommandButton btnCancel   Caption="中止"
' ============================================================

' ============================================================
' 初期化：画像・進捗・ダミーデフォルト値をセット
' ============================================================
Public Sub SetupForm(filePath As String, current As Integer, total As Integer)
    ' 進捗
    lblProgress.Caption = current & " / " & total

    ' ファイル名
    lblFileName.Caption = Mid(filePath, InStrRev(filePath, "\") + 1)

    ' 画像プレビュー（読み込み失敗しても続行）
    On Error Resume Next
    imgPreview.Picture = LoadPicture(filePath)
    On Error GoTo 0

    ' ダミーデフォルト値
    txtCarName.Text  = "スズキ"
    txtGrade.Text    = "クロスビー MZ"
    txtYear.Text     = "2021"
    txtMonth.Text    = "7"
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
' 入力値をCarData型で返す
' ============================================================
Public Function GetData() As CarData
    Dim data As CarData

    ' 車名＋グレードを結合
    Dim carName As String
    carName = Trim(txtCarName.Text)
    If Trim(txtGrade.Text) <> "" Then
        carName = carName & " " & Trim(txtGrade.Text)
    End If

    ' 年式/月を結合
    Dim yearMonth As String
    yearMonth = Trim(txtYear.Text)
    If Trim(txtMonth.Text) <> "" Then
        yearMonth = yearMonth & "/" & Trim(txtMonth.Text)
    End If

    data.YearMonth   = yearMonth
    data.CarName     = carName
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

' ============================================================
' ボタンイベント
' ============================================================
Private Sub btnImport_Click()
    ' 簡易バリデーション
    If Trim(txtCarName.Text) = "" Then
        MsgBox "車名を入力してください。", vbExclamation
        txtCarName.SetFocus
        Exit Sub
    End If

    Me.Hide   ' 呼び出し元に制御を返す（Unloadはしない）
End Sub

Private Sub btnCancel_Click()
    If MsgBox("処理を中止しますか？", vbQuestion + vbYesNo) = vbYes Then
        g_Cancelled = True
        Me.Hide
    End If
End Sub
