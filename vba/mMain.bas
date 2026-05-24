Attribute VB_Name = "mMain"
Option Explicit

' ============================================================
' グローバル変数
' ============================================================
Public g_SheetName  As String
Public g_FolderPath As String
Public g_Cancelled  As Boolean

' ============================================================
' 1台分のデータ
' ============================================================
Public Type CarData
    CarNumber    As String   ' 管理番号 (例: 27-001)
    PurchaseDate As String   ' 仕入れ日 (例: 2025/5/24)
    Session      As String   ' 回次
    YearMonth    As String   ' 年式/月 (例: 2021/7)
    CarName      As String   ' 車名（グレード含む）
    Shaken       As String   ' 車検 (例: 7/12)
    Score        As Variant  ' 評価点
    Price        As Variant  ' 車輌代
    Tax          As Variant  ' 消費税
    CarTax       As Variant  ' 自税
    Recycle      As Variant  ' リサイクル
    AuctionFee   As Variant  ' 落札料
    Venue        As String   ' 仕入れ先（会場）
    LotNumber    As String   ' 出品番号
    Color        As String   ' 色
    Chassis      As String   ' 車台番号
    Mileage      As Variant  ' 走行距離
End Type

' ============================================================
' エントリポイント
' ============================================================
Sub Main()
    g_SheetName  = "27期"
    g_FolderPath = ""
    g_Cancelled  = False

    frmSettings.Show
End Sub

' ============================================================
' A列を検索して連番の最大値+1を返す
' パターン: \d+-\d+ (例: 27-005)
' ============================================================
Function GetNextNumber(ws As Worksheet) As Long
    Dim i       As Long
    Dim maxNum  As Long
    Dim cellVal As String
    Dim parts() As String

    maxNum = 0

    For i = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row To 1 Step -1
        On Error Resume Next
        cellVal = Trim(CStr(ws.Cells(i, 1).Value2))
        On Error GoTo 0

        If InStr(cellVal, "-") > 0 And _
           InStr(cellVal, "-") = InStrRev(cellVal, "-") Then
            parts = Split(cellVal, "-")
            If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then
                Dim num As Long
                num = CLng(parts(1))
                If num > maxNum Then maxNum = num
            End If
        End If
    Next i

    GetNextNumber = maxNum + 1
End Function

' ============================================================
' 次の書き込み先行番号を返す
' Q列(17列)の「着」を末尾として次行を算出
' ============================================================
Function GetNextWriteRow(ws As Worksheet) As Long
    Dim lastRow As Long
    Dim qVal    As String

    lastRow = ws.Cells(ws.Rows.Count, 17).End(xlUp).Row

    If lastRow < 10 Then
        GetNextWriteRow = 10
        Exit Function
    End If

    qVal = Trim(CStr(ws.Cells(lastRow, 17).Value2))

    Select Case qVal
        Case "着"
            GetNextWriteRow = lastRow + 1
        Case "予定"
            GetNextWriteRow = lastRow + 2
        Case Else
            GetNextWriteRow = lastRow + 1
    End Select
End Function

' ============================================================
' シート名から期番号を取得 ("27期" → "27")
' ============================================================
Function GetPeriodNumber(ws As Worksheet) As String
    Dim s      As String
    Dim i      As Integer
    Dim result As String

    s      = ws.Name
    result = ""

    For i = 1 To Len(s)
        If IsNumeric(Mid(s, i, 1)) Then
            result = result & Mid(s, i, 1)
        Else
            Exit For
        End If
    Next i

    If result = "" Then result = "27"
    GetPeriodNumber = result
End Function

' ============================================================
' シートに2行書き込み (Value2のみ使用・書式変更なし)
' ============================================================
Sub WriteToSheet(ws As Worksheet, data As CarData)
    Dim r As Long
    r = GetNextWriteRow(ws)

    ' --- 行1 ---
    ws.Cells(r, 1).Value2  = data.CarNumber      ' A: 管理番号
    ws.Cells(r, 4).Value2  = data.YearMonth       ' D: 年式/月
    ws.Cells(r, 6).Value2  = data.CarName         ' F: 車名
    ws.Cells(r, 17).Value2 = "予定"               ' Q: 付属品

    If data.PurchaseDate <> "" Then ws.Cells(r, 2).Value2  = data.PurchaseDate  ' B: 仕入れ日
    If data.Session      <> "" Then ws.Cells(r, 3).Value2  = data.Session       ' C: 回次
    If data.Shaken       <> "" Then ws.Cells(r, 7).Value2  = data.Shaken        ' G: 車検
    If data.Score        <> "" Then ws.Cells(r, 8).Value2  = data.Score         ' H: 評価点
    If data.Price        <> "" Then ws.Cells(r, 9).Value2  = CLng(data.Price)   ' I: 車輌代
    If data.Tax          <> "" Then ws.Cells(r, 10).Value2 = CLng(data.Tax)     ' J: 消費税
    If data.CarTax       <> "" Then ws.Cells(r, 11).Value2 = CLng(data.CarTax)  ' K: 自税
    If data.Recycle      <> "" Then ws.Cells(r, 12).Value2 = CLng(data.Recycle) ' L: リサイクル
    If data.AuctionFee   <> "" Then ws.Cells(r, 13).Value2 = CLng(data.AuctionFee) ' M: 落札料

    ' --- 行2 ---
    ws.Cells(r + 1, 2).Value2  = data.Venue       ' B: 仕入れ先（会場）
    ws.Cells(r + 1, 3).Value2  = data.LotNumber   ' C: 出品番号
    ws.Cells(r + 1, 4).Value2  = data.Color        ' D: 色
    ws.Cells(r + 1, 6).Value2  = data.Chassis      ' F: 車台番号
    ws.Cells(r + 1, 17).Value2 = "着"              ' Q: 付属品

    If data.Mileage <> "" Then ws.Cells(r + 1, 7).Value2 = CLng(data.Mileage) ' G: 走行距離
End Sub

' ============================================================
' 画像を「処理済み」サブフォルダへ移動
' ============================================================
Sub MoveToProcessed(filePath As String)
    On Error GoTo ErrHandler

    Dim folder    As String
    Dim fileName  As String
    Dim processed As String
    Dim dest      As String

    folder    = Left(filePath, InStrRev(filePath, "\"))
    fileName  = Mid(filePath, InStrRev(filePath, "\") + 1)
    processed = folder & "処理済み\"
    dest      = processed & fileName

    If Dir(processed, vbDirectory) = "" Then MkDir processed

    If Dir(dest) <> "" Then
        Dim dot  As Integer
        Dim base As String
        Dim ext  As String
        Dim cnt  As Integer
        dot  = InStrRev(fileName, ".")
        base = Left(fileName, dot - 1)
        ext  = Mid(fileName, dot)
        cnt  = 1
        Do While Dir(processed & base & "_" & cnt & ext) <> ""
            cnt = cnt + 1
        Loop
        dest = processed & base & "_" & cnt & ext
    End If

    FileCopy filePath, dest
    Kill filePath
    Exit Sub

ErrHandler:
    MsgBox "ファイル移動に失敗しました:" & vbCrLf & filePath & vbCrLf & _
           Err.Description, vbExclamation
End Sub
