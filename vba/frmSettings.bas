Attribute VB_Name = "frmSettings"
Option Explicit

Private Sub UserForm_Initialize()
    txtSheetName.Text = g_SheetName
    txtFolder.Text    = g_FolderPath
End Sub

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
    Dim exts(3) As String
    exts(0) = "*.jpg" : exts(1) = "*.jpeg"
    exts(2) = "*.png" : exts(3) = "*.bmp"

    Dim fileList As New Collection
    Dim i As Integer, f As String
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

    ' --- Collection → 配列変換 ---
    Dim files() As String
    ReDim files(0 To fileList.Count - 1)
    Dim j As Integer
    j = 0
    Dim item As Variant
    For Each item In fileList
        files(j) = CStr(item)
        j = j + 1
    Next item

    Me.Hide

    ' --- 1枚ずつ処理（1取込 = 1件追記） ---
    Dim processed As Integer
    processed = 0

    For i = 0 To UBound(files)
        frmInput.SetupForm files, i
        frmInput.Show vbModal

        If g_Cancelled Then
            MsgBox "処理を中止しました。" & vbCrLf & _
                   processed & " / " & fileList.Count & " 件処理済み", vbInformation
            Unload frmInput
            Unload Me
            Exit Sub
        End If

        Dim data As CarData
        data           = frmInput.GetData()
        data.CarNumber = GetPeriodNumber(ws) & "-" & _
                         Format(GetNextNumber(ws), "000")

        WriteToSheet ws, data
        MoveToProcessed files(i)
        processed = processed + 1
    Next i

    MsgBox processed & " 件を処理しました。", vbInformation
    Unload frmInput
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
