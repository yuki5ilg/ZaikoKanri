Attribute VB_Name = "mMain"
Option Explicit

Public g_SheetName  As String
Public g_FolderPath As String

Public Type CarData
    CarNumber    As String   ' A
    PurchaseDate As String   ' B1: 仕入れ日
    Supplier     As String   ' B2: 仕入れ先
    Session      As String   ' C1: 回次
    LotNumber    As String   ' C2: 出品番号
    YearMonth    As String   ' D1: 年式
    Color        As String   ' D2: 色
    CarName      As String   ' F1: 車体番号（車名）
    Chassis      As String   ' F2: 車台番号
    Shaken       As String   ' G1: 車検
    Mileage      As Variant  ' G2: 距離
    Score        As Variant  ' H:  評価点
    Price        As Variant  ' I:  車輌代
    Tax          As Variant  ' J:  消費税
    CarTax       As Variant  ' K:  自税
    Recycle      As Variant  ' L:  リサイクル
    AuctionFee   As Variant  ' M:  落札料
    Total        As Variant  ' N:  合計
    Loss         As Variant  ' O:  評価損
    Plate        As String   ' P1: 車輌番号
    Owner        As String   ' P2: 所有者
    Accessories  As String   ' Q:  付属品
    Memo         As String   ' T:  補足
End Type

Public Type SaleData
    CarNumber   As String
    Meigi       As String   ' U:  名義変更
    SaleDate    As String   ' V1: 売上日
    Buyer       As String   ' V2: 売上先
    SaleSession As String   ' W1: 回次
    SaleLot     As String   ' W2: 出品番号
    SalePrice   As Variant  ' X:  車輌代
    SaleTax     As Variant  ' Y:  消費税
    SaleRecycle As Variant  ' Z:  リサイクル
    SaleTotal   As Variant  ' AA: 合計
    PaymentDate As String   ' AB: 入金日
End Type

Sub Main()
    g_SheetName  = "27期"
    g_FolderPath = ""
    frmInput.Show
End Sub

Function GetNextNumber(ws As Worksheet) As Long
    Dim i As Long, maxNum As Long, cellVal As String, parts() As String
    maxNum = 0
    For i = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row To 1 Step -1
        On Error Resume Next
        cellVal = Trim(CStr(ws.Cells(i, 1).Value2))
        On Error GoTo 0
        If InStr(cellVal, "-") > 0 And InStr(cellVal, "-") = InStrRev(cellVal, "-") Then
            parts = Split(cellVal, "-")
            If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then
                Dim n As Long : n = CLng(parts(1))
                If n > maxNum Then maxNum = n
            End If
        End If
    Next i
    GetNextNumber = maxNum + 1
End Function

Function GetNextWriteRow(ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 9 Then
        GetNextWriteRow = 10
    Else
        GetNextWriteRow = lastRow + 2
    End If
End Function

Function GetPeriodNumber(ws As Worksheet) As String
    Dim s As String, i As Integer, result As String
    s = ws.Name : result = ""
    For i = 1 To Len(s)
        If IsNumeric(Mid(s, i, 1)) Then result = result & Mid(s, i, 1) Else Exit For
    Next i
    If result = "" Then result = "27"
    GetPeriodNumber = result
End Function

Function FindCarRow(ws As Worksheet, carNum As String) As Long
    Dim i As Long
    For i = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Trim(CStr(ws.Cells(i, 1).Value2)) = Trim(carNum) Then
            FindCarRow = i : Exit Function
        End If
    Next i
    FindCarRow = 0
End Function

Sub WriteToSheet(ws As Worksheet, data As CarData)
    Dim r As Long
    r = GetNextWriteRow(ws)

    ws.Cells(r, 1).Value2 = data.CarNumber

    If data.PurchaseDate <> "" Then ws.Cells(r,  2).Value2 = data.PurchaseDate
    If data.Session      <> "" Then ws.Cells(r,  3).Value2 = data.Session
    If data.YearMonth    <> "" Then ws.Cells(r,  4).Value2 = data.YearMonth
    If data.CarName      <> "" Then ws.Cells(r,  6).Value2 = data.CarName
    If data.Shaken       <> "" Then ws.Cells(r,  7).Value2 = data.Shaken
    If data.Score        <> "" Then ws.Cells(r,  8).Value2 = data.Score
    If data.Price        <> "" Then ws.Cells(r,  9).Value2 = CLng(data.Price)
    If data.Tax          <> "" Then ws.Cells(r, 10).Value2 = CLng(data.Tax)
    If data.CarTax       <> "" Then ws.Cells(r, 11).Value2 = CLng(data.CarTax)
    If data.Recycle      <> "" Then ws.Cells(r, 12).Value2 = CLng(data.Recycle)
    If data.AuctionFee   <> "" Then ws.Cells(r, 13).Value2 = CLng(data.AuctionFee)
    If data.Total        <> "" Then ws.Cells(r, 14).Value2 = CLng(data.Total)
    If data.Loss         <> "" Then ws.Cells(r, 15).Value2 = CLng(data.Loss)
    If data.Plate        <> "" Then ws.Cells(r, 16).Value2 = data.Plate
    If data.Accessories  <> "" Then ws.Cells(r, 17).Value2 = data.Accessories
    If data.Memo         <> "" Then ws.Cells(r, 20).Value2 = data.Memo

    If data.Supplier  <> "" Then ws.Cells(r+1,  2).Value2 = data.Supplier
    If data.LotNumber <> "" Then ws.Cells(r+1,  3).Value2 = data.LotNumber
    If data.Color     <> "" Then ws.Cells(r+1,  4).Value2 = data.Color
    If data.Chassis   <> "" Then ws.Cells(r+1,  6).Value2 = data.Chassis
    If data.Mileage   <> "" Then ws.Cells(r+1,  7).Value2 = CLng(data.Mileage)
    If data.Owner     <> "" Then ws.Cells(r+1, 16).Value2 = data.Owner
End Sub

Sub WriteToSheetSales(ws As Worksheet, data As SaleData)
    Dim r As Long
    r = FindCarRow(ws, data.CarNumber)
    If r = 0 Then Exit Sub

    If data.Meigi       <> "" Then ws.Cells(r,   21).Value2 = data.Meigi
    If data.SaleDate    <> "" Then ws.Cells(r,   22).Value2 = data.SaleDate
    If data.Buyer       <> "" Then ws.Cells(r+1, 22).Value2 = data.Buyer
    If data.SaleSession <> "" Then ws.Cells(r,   23).Value2 = data.SaleSession
    If data.SaleLot     <> "" Then ws.Cells(r+1, 23).Value2 = data.SaleLot
    If data.SalePrice   <> "" Then ws.Cells(r,   24).Value2 = CLng(data.SalePrice)
    If data.SaleTax     <> "" Then ws.Cells(r,   25).Value2 = CLng(data.SaleTax)
    If data.SaleRecycle <> "" Then ws.Cells(r,   26).Value2 = CLng(data.SaleRecycle)
    If data.SaleTotal   <> "" Then ws.Cells(r,   27).Value2 = CLng(data.SaleTotal)
    If data.PaymentDate <> "" Then ws.Cells(r,   28).Value2 = data.PaymentDate
End Sub

Sub MoveToProcessed(filePath As String)
    On Error GoTo ErrHandler
    Dim folder As String, fileName As String, processed As String, dest As String
    folder    = Left(filePath, InStrRev(filePath, "\"))
    fileName  = Mid(filePath, InStrRev(filePath, "\") + 1)
    processed = folder & "処理済み\"
    dest      = processed & fileName
    If Dir(processed, vbDirectory) = "" Then MkDir processed
    If Dir(dest) <> "" Then
        Dim dot As Integer, base As String, ext As String, cnt As Integer
        dot = InStrRev(fileName, ".")
        base = Left(fileName, dot - 1) : ext = Mid(fileName, dot) : cnt = 1
        Do While Dir(processed & base & "_" & cnt & ext) <> ""
            cnt = cnt + 1
        Loop
        dest = processed & base & "_" & cnt & ext
    End If
    FileCopy filePath, dest
    Kill filePath
    Exit Sub
ErrHandler:
    MsgBox "ファイル移動に失敗しました:" & vbCrLf & filePath & vbCrLf & Err.Description, vbExclamation
End Sub
