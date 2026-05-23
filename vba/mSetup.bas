Attribute VB_Name = "mSetup"
Option Explicit

' ============================================================
' フォームを自動生成するセットアップマクロ
'
' 【事前設定】
' Excelのオプション → トラストセンター → マクロの設定
' →「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェック
'
' 【使い方】
' Alt+F8 → CreateForms を実行
' ============================================================

Public Sub CreateForms()
    Dim vbp As Object
    On Error Resume Next
    Set vbp = ThisWorkbook.VBProject
    On Error GoTo 0

    If vbp Is Nothing Then
        MsgBox "VBAプロジェクトへのアクセスが無効です。" & vbCrLf & vbCrLf & _
               "Excelのオプション → トラストセンター → マクロの設定 →" & vbCrLf & _
               "「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」" & vbCrLf & _
               "にチェックを入れてから再実行してください。", vbExclamation
        Exit Sub
    End If

    ' 既存フォームを削除
    Dim comp   As Object
    Dim killMe As New Collection
    For Each comp In vbp.VBComponents
        If comp.Name = "frmSettings" Or comp.Name = "frmInput" Then
            killMe.Add comp
        End If
    Next comp
    Dim v As Variant
    For Each v In killMe
        vbp.VBComponents.Remove v
    Next v

    BuildFrmSettings vbp
    BuildFrmInput vbp

    MsgBox "フォームの生成が完了しました。" & vbCrLf & _
           "Alt+F8 → Main を実行して使用してください。", vbInformation
End Sub

' ============================================================
' frmSettings を生成
' ============================================================
Private Sub BuildFrmSettings(vbp As Object)
    Dim comp As Object
    Dim frm  As Object
    Dim c    As Object

    Set comp = vbp.VBComponents.Add(2)   ' 2 = vbext_ct_MSForm
    comp.Name = "frmSettings"

    Set frm = comp.Designer
    frm.Caption         = "設定"
    frm.Width           = 400
    frm.Height          = 172
    frm.StartUpPosition = 1

    ' シート名
    Set c = frm.Controls.Add("Forms.Label.1", "lblSheetName")
    c.Caption = "対象シート名" : c.Left = 10 : c.Top = 16 : c.Width = 92 : c.Height = 18

    Set c = frm.Controls.Add("Forms.TextBox.1", "txtSheetName")
    c.Text = "27期" : c.Left = 108 : c.Top = 12 : c.Width = 160 : c.Height = 22

    ' フォルダ
    Set c = frm.Controls.Add("Forms.Label.1", "lblFolder")
    c.Caption = "画像フォルダ" : c.Left = 10 : c.Top = 48 : c.Width = 92 : c.Height = 18

    Set c = frm.Controls.Add("Forms.TextBox.1", "txtFolder")
    c.Left = 108 : c.Top = 44 : c.Width = 216 : c.Height = 22

    Set c = frm.Controls.Add("Forms.CommandButton.1", "btnBrowse")
    c.Caption = "参照..." : c.Left = 330 : c.Top = 42 : c.Width = 52 : c.Height = 24

    ' ボタン
    Set c = frm.Controls.Add("Forms.CommandButton.1", "btnOK")
    c.Caption = "OK" : c.Default = True
    c.Left = 222 : c.Top = 84 : c.Width = 72 : c.Height = 28

    Set c = frm.Controls.Add("Forms.CommandButton.1", "btnCancel")
    c.Caption = "キャンセル" : c.Cancel = True
    c.Left = 302 : c.Top = 84 : c.Width = 80 : c.Height = 28

    comp.CodeModule.AddFromString FrmSettingsCode()
End Sub

' ============================================================
' frmInput を生成
' ============================================================
Private Sub BuildFrmInput(vbp As Object)
    Dim comp As Object
    Dim frm  As Object
    Dim c    As Object

    Set comp = vbp.VBComponents.Add(2)
    comp.Name = "frmInput"

    Set frm = comp.Designer
    frm.Caption         = "在庫取込"
    frm.Width           = 756
    frm.Height          = 560
    frm.StartUpPosition = 1

    ' 進捗ラベル
    Set c = frm.Controls.Add("Forms.Label.1", "lblProgress")
    c.Caption = "" : c.TextAlign = 3
    c.Left = 228 : c.Top = 5 : c.Width = 66 : c.Height = 16

    ' 画像プレビュー
    Set c = frm.Controls.Add("Forms.Image.1", "imgPreview")
    c.PictureSizeMode = 3 : c.BorderStyle = 1 : c.BackColor = &H808080
    c.Left = 6 : c.Top = 24 : c.Width = 288 : c.Height = 372

    ' ファイル名
    Set c = frm.Controls.Add("Forms.Label.1", "lblFileName")
    c.Caption = "" : c.WordWrap = True
    c.Left = 6 : c.Top = 400 : c.Width = 288 : c.Height = 28

    ' 入力フィールド（15項目）
    AddInputFields frm

    ' 取込ボタン
    Set c = frm.Controls.Add("Forms.CommandButton.1", "btnImport")
    c.Caption = "取込" : c.Default = True : c.Font.Bold = True : c.Font.Size = 12
    c.Left = 354 : c.Top = 466 : c.Width = 156 : c.Height = 38

    ' 中止ボタン
    Set c = frm.Controls.Add("Forms.CommandButton.1", "btnCancel")
    c.Caption = "中止" : c.Cancel = True
    c.Left = 520 : c.Top = 470 : c.Width = 80 : c.Height = 32

    comp.CodeModule.AddFromString FrmInputCode()
End Sub

' ============================================================
' 入力フィールド 15行を追加
' ============================================================
Private Sub AddInputFields(frm As Object)
    Dim labels(14) As String
    Dim names(14)  As String
    Dim c          As Object
    Dim i          As Integer
    Dim topY       As Integer

    labels(0)  = "車名"              : names(0)  = "txtCarName"
    labels(1)  = "グレード"          : names(1)  = "txtGrade"
    labels(2)  = "年式"              : names(2)  = "txtYear"
    labels(3)  = "月"                : names(3)  = "txtMonth"
    labels(4)  = "色"                : names(4)  = "txtColor"
    labels(5)  = "車台番号"          : names(5)  = "txtChassis"
    labels(6)  = "評価点"            : names(6)  = "txtScore"
    labels(7)  = "走行距離(km)"      : names(7)  = "txtMileage"
    labels(8)  = "落札価格(円)"      : names(8)  = "txtPrice"
    labels(9)  = "消費税(円)"        : names(9)  = "txtTax"
    labels(10) = "自動車税(円)"      : names(10) = "txtCarTax"
    labels(11) = "リサイクル料(円)"  : names(11) = "txtRecycle"
    labels(12) = "落札手数料(円)"    : names(12) = "txtFee"
    labels(13) = "オークション会場"  : names(13) = "txtVenue"
    labels(14) = "出品番号"          : names(14) = "txtLotNum"

    topY = 8
    For i = 0 To 14
        Set c = frm.Controls.Add("Forms.Label.1", "lbl" & Mid(names(i), 4))
        c.Caption = labels(i) : c.TextAlign = 3
        c.Left = 306 : c.Top = topY + 3 : c.Width = 96 : c.Height = 18

        Set c = frm.Controls.Add("Forms.TextBox.1", names(i))
        c.Left = 408 : c.Top = topY : c.Width = 192 : c.Height = 22

        topY = topY + 30
    Next i
End Sub

' ============================================================
' frmSettings のコード
' ============================================================
Private Function FrmSettingsCode() As String
    Dim q As String : q = Chr(34)
    Dim n As String : n = vbCrLf
    Dim c As String

    c = "Option Explicit" & n & n

    ' Initialize
    c = c & "Private Sub UserForm_Initialize()" & n
    c = c & "    txtSheetName.Text = g_SheetName" & n
    c = c & "    txtFolder.Text    = g_FolderPath" & n
    c = c & "End Sub" & n & n

    ' Browse
    c = c & "Private Sub btnBrowse_Click()" & n
    c = c & "    With Application.FileDialog(msoFileDialogFolderPicker)" & n
    c = c & "        .Title = " & q & "画像フォルダを選択してください" & q & n
    c = c & "        .AllowMultiSelect = False" & n
    c = c & "        If .Show = True Then txtFolder.Text = .SelectedItems(1)" & n
    c = c & "    End With" & n
    c = c & "End Sub" & n & n

    ' OK
    c = c & "Private Sub btnOK_Click()" & n
    c = c & "    If Trim(txtSheetName.Text) = " & q & q & " Then" & n
    c = c & "        MsgBox " & q & "シート名を入力してください。" & q & ", vbExclamation" & n
    c = c & "        txtSheetName.SetFocus : Exit Sub" & n
    c = c & "    End If" & n
    c = c & "    If Trim(txtFolder.Text) = " & q & q & " Then" & n
    c = c & "        MsgBox " & q & "画像フォルダを選択してください。" & q & ", vbExclamation : Exit Sub" & n
    c = c & "    End If" & n
    c = c & "    If Dir(Trim(txtFolder.Text), vbDirectory) = " & q & q & " Then" & n
    c = c & "        MsgBox " & q & "フォルダが見つかりません。" & q & ", vbExclamation : Exit Sub" & n
    c = c & "    End If" & n
    c = c & "    Dim ws As Worksheet" & n
    c = c & "    On Error Resume Next" & n
    c = c & "    Set ws = ThisWorkbook.Sheets(Trim(txtSheetName.Text))" & n
    c = c & "    On Error GoTo 0" & n
    c = c & "    If ws Is Nothing Then" & n
    c = c & "        MsgBox " & q & "シート " & q & " & txtSheetName.Text & " & q & " が見つかりません。" & q & ", vbExclamation : Exit Sub" & n
    c = c & "    End If" & n
    c = c & "    g_SheetName  = Trim(txtSheetName.Text)" & n
    c = c & "    g_FolderPath = Trim(txtFolder.Text)" & n
    ' ファイル収集
    c = c & "    Dim fileList As New Collection" & n
    c = c & "    Dim exts(3) As String" & n
    c = c & "    exts(0) = " & q & "*.jpg" & q & " : exts(1) = " & q & "*.jpeg" & q & n
    c = c & "    exts(2) = " & q & "*.png" & q & " : exts(3) = " & q & "*.bmp" & q & n
    c = c & "    Dim ei As Integer : Dim f As String" & n
    c = c & "    For ei = 0 To 3" & n
    c = c & "        f = Dir(g_FolderPath & " & q & "\" & q & " & exts(ei))" & n
    c = c & "        Do While f <> " & q & q & n
    c = c & "            fileList.Add g_FolderPath & " & q & "\" & q & " & f" & n
    c = c & "            f = Dir()" & n
    c = c & "        Loop" & n
    c = c & "    Next ei" & n
    c = c & "    If fileList.Count = 0 Then" & n
    c = c & "        MsgBox " & q & "画像ファイルが見つかりませんでした。" & q & ", vbInformation : Exit Sub" & n
    c = c & "    End If" & n
    c = c & "    Me.Hide" & n
    ' 処理ループ
    c = c & "    Dim processed As Integer : Dim total As Integer" & n
    c = c & "    Dim item As Variant" & n
    c = c & "    processed = 0 : total = fileList.Count" & n
    c = c & "    For Each item In fileList" & n
    c = c & "        frmInput.SetupForm CStr(item), processed + 1, total" & n
    c = c & "        frmInput.Show vbModal" & n
    c = c & "        If g_Cancelled Then" & n
    c = c & "            MsgBox " & q & "中止しました。(" & q & " & processed & " & q & "/" & q & " & total & " & q & "枚完了)" & q & ", vbInformation" & n
    c = c & "            Unload frmInput : Unload Me : Exit Sub" & n
    c = c & "        End If" & n
    c = c & "        Dim data As CarData" & n
    c = c & "        data = frmInput.GetData()" & n
    c = c & "        data.CarNumber = GetPeriodNumber(ws) & " & q & "-" & q & " & Format(GetNextNumber(ws), " & q & "000" & q & ")" & n
    c = c & "        WriteToSheet ws, data" & n
    c = c & "        MoveToProcessed CStr(item)" & n
    c = c & "        processed = processed + 1" & n
    c = c & "    Next item" & n
    c = c & "    MsgBox processed & " & q & "枚を処理しました。" & q & ", vbInformation" & n
    c = c & "    Unload frmInput : Unload Me" & n
    c = c & "End Sub" & n & n

    ' Cancel
    c = c & "Private Sub btnCancel_Click()" & n
    c = c & "    Unload Me" & n
    c = c & "End Sub" & n

    FrmSettingsCode = c
End Function

' ============================================================
' frmInput のコード
' ============================================================
Private Function FrmInputCode() As String
    Dim q As String : q = Chr(34)
    Dim n As String : n = vbCrLf
    Dim c As String

    c = "Option Explicit" & n & n

    ' SetupForm
    c = c & "Public Sub SetupForm(filePath As String, current As Integer, total As Integer)" & n
    c = c & "    lblProgress.Caption = current & " & q & " / " & q & " & total" & n
    c = c & "    lblFileName.Caption = Mid(filePath, InStrRev(filePath, " & q & "\" & q & ") + 1)" & n
    c = c & "    On Error Resume Next" & n
    c = c & "    imgPreview.Picture = LoadPicture(filePath)" & n
    c = c & "    On Error GoTo 0" & n
    c = c & "    txtCarName.Text  = " & q & "スズキ" & q & n
    c = c & "    txtGrade.Text    = " & q & "クロスビー MZ" & q & n
    c = c & "    txtYear.Text     = " & q & "2021" & q & n
    c = c & "    txtMonth.Text    = " & q & "7" & q & n
    c = c & "    txtColor.Text    = " & q & "ホワイトパール" & q & n
    c = c & "    txtChassis.Text  = " & q & "MYN15S-100001" & q & n
    c = c & "    txtScore.Text    = " & q & "4.5" & q & n
    c = c & "    txtMileage.Text  = " & q & "35000" & q & n
    c = c & "    txtPrice.Text    = " & q & "980000" & q & n
    c = c & "    txtTax.Text      = " & q & "98000" & q & n
    c = c & "    txtCarTax.Text   = " & q & "35400" & q & n
    c = c & "    txtRecycle.Text  = " & q & "12000" & q & n
    c = c & "    txtFee.Text      = " & q & "32000" & q & n
    c = c & "    txtVenue.Text    = " & q & "USS大阪" & q & n
    c = c & "    txtLotNum.Text   = " & q & "12345" & q & n
    c = c & "End Sub" & n & n

    ' GetData
    c = c & "Public Function GetData() As CarData" & n
    c = c & "    Dim data As CarData" & n
    c = c & "    Dim carName As String : carName = Trim(txtCarName.Text)" & n
    c = c & "    If Trim(txtGrade.Text) <> " & q & q & " Then carName = carName & " & q & " " & q & " & Trim(txtGrade.Text)" & n
    c = c & "    Dim ym As String : ym = Trim(txtYear.Text)" & n
    c = c & "    If Trim(txtMonth.Text) <> " & q & q & " Then ym = ym & " & q & "/" & q & " & Trim(txtMonth.Text)" & n
    c = c & "    data.YearMonth  = ym" & n
    c = c & "    data.CarName    = carName" & n
    c = c & "    data.Score      = Trim(txtScore.Text)" & n
    c = c & "    data.Price      = Trim(txtPrice.Text)" & n
    c = c & "    data.Tax        = Trim(txtTax.Text)" & n
    c = c & "    data.CarTax     = Trim(txtCarTax.Text)" & n
    c = c & "    data.Recycle    = Trim(txtRecycle.Text)" & n
    c = c & "    data.AuctionFee = Trim(txtFee.Text)" & n
    c = c & "    data.Venue      = Trim(txtVenue.Text)" & n
    c = c & "    data.LotNumber  = Trim(txtLotNum.Text)" & n
    c = c & "    data.Color      = Trim(txtColor.Text)" & n
    c = c & "    data.Chassis    = Trim(txtChassis.Text)" & n
    c = c & "    data.Mileage    = Trim(txtMileage.Text)" & n
    c = c & "    GetData = data" & n
    c = c & "End Function" & n & n

    ' btnImport
    c = c & "Private Sub btnImport_Click()" & n
    c = c & "    If Trim(txtCarName.Text) = " & q & q & " Then" & n
    c = c & "        MsgBox " & q & "車名を入力してください。" & q & ", vbExclamation" & n
    c = c & "        txtCarName.SetFocus : Exit Sub" & n
    c = c & "    End If" & n
    c = c & "    Me.Hide" & n
    c = c & "End Sub" & n & n

    ' btnCancel
    c = c & "Private Sub btnCancel_Click()" & n
    c = c & "    If MsgBox(" & q & "処理を中止しますか？" & q & ", vbQuestion + vbYesNo) = vbYes Then" & n
    c = c & "        g_Cancelled = True : Me.Hide" & n
    c = c & "    End If" & n
    c = c & "End Sub" & n

    FrmInputCode = c
End Function
