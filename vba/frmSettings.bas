Attribute VB_Name = "frmSettings"
Option Explicit

' ============================================================
' frmSettings - 設定フォーム
'
' コントロール一覧（VBAデザイナーで作成）:
'   Label        lblSheetName    Caption="対象シート名"
'   TextBox      txtSheetName    Text="27期"
'   Label        lblFolder       Caption="画像フォルダ"
'   TextBox      txtFolder       （幅広め）
'   CommandButton btnBrowse      Caption="参照..."
'   CommandButton btnOK          Caption="OK"
'   CommandButton btnCancel      Caption="キャンセル"
' ============================================================

Private Sub UserForm_Initialize()
    txtSheetName.Text = g_SheetName
    txtFolder.Text    = g_FolderPath
End Sub

' フォルダ選択ダイアログ
Private Sub btnBrowse_Click()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "画像フォルダを選択してください"
        .AllowMultiSelect = False
        If .Show = True Then
            txtFolder.Text = .SelectedItems(1)
        End If
    End With
End Sub

Private Sub btnOK_Click()
    ' --- バリデーション ---
    If Trim(txtSheetName.Text) = "" Then
        MsgBox "シート名を入力してください。", vbExclamation
        txtSheetName.SetFocus
        Exit Sub
    End If

    If Trim(txtFolder.Text) = "" Then
        MsgBox "画像フォルダを選択してください。", vbExclamation
        Exit Sub
    End If

    If Dir(Trim(txtFolder.Text), vbDirectory) = "" Then
        MsgBox "フォルダが見つかりません:" & vbCrLf & txtFolder.Text, vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Trim(txtSheetName.Text))
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "シート「" & txtSheetName.Text & "」が見つかりません。", vbExclamation
        txtSheetName.SetFocus
        Exit Sub
    End If

    ' --- 設定確定 ---
    g_SheetName  = Trim(txtSheetName.Text)
    g_FolderPath = Trim(txtFolder.Text)

    ' --- 画像ファイル収集 ---
    Dim fileList As New Collection
    Dim exts(3)  As String
    exts(0) = "*.jpg" : exts(1) = "*.jpeg"
    exts(2) = "*.png" : exts(3) = "*.bmp"

    Dim i As Integer
    Dim f As String
    For i = 0 To 3
        f = Dir(g_FolderPath & "\" & exts(i))
        Do While f <> ""
            fileList.Add g_FolderPath & "\" & f
            f = Dir()
        Loop
    Next i

    If fileList.Count = 0 Then
        MsgBox "画像ファイル（JPG/PNG/BMP）が見つかりませんでした。", vbInformation
        Exit Sub
    End If

    Me.Hide

    ' --- 画像ごとに入力フォーム表示 ---
    Dim processed As Integer
    Dim total     As Integer
    Dim item      As Variant

    processed = 0
    total     = fileList.Count

    For Each item In fileList
        frmInput.SetupForm CStr(item), processed + 1, total
        frmInput.Show vbModal

        If g_Cancelled Then
            MsgBox "処理を中止しました。（" & processed & " / " & total & " 枚完了）", _
                   vbInformation
            Unload frmInput
            Unload Me
            Exit Sub
        End If

        ' 管理番号をここで採番（書き込み直前）
        Dim data As CarData
        data           = frmInput.GetData()
        data.CarNumber = GetPeriodNumber(ws) & "-" & _
                         Format(GetNextNumber(ws), "000")

        WriteToSheet ws, data
        MoveToProcessed CStr(item)
        processed = processed + 1
    Next item

    MsgBox processed & "枚の画像を処理しました。", vbInformation
    Unload frmInput
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
