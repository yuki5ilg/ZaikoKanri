Attribute VB_Name = "frmSales"
Option Explicit

Private Sub UserForm_Initialize()
    txtSheetName.Text = g_SheetName
    txtSaleDate.Text  = Format(Now, "yyyy/m/d")
End Sub

Private Sub btnOK_Click()
    If Trim(txtCarNum.Text) = "" Then
        MsgBox "仕入番号を入力してください。", vbExclamation
        txtCarNum.SetFocus : Exit Sub
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Trim(txtSheetName.Text))
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "シート「" & txtSheetName.Text & "」が見つかりません。", vbExclamation : Exit Sub
    End If

    Dim data As SaleData
    data.CarNumber   = Trim(txtCarNum.Text)
    data.Meigi       = Trim(txtMeigi.Text)
    data.SaleDate    = Trim(txtSaleDate.Text)
    data.Buyer       = Trim(txtBuyer.Text)
    data.SaleSession = Trim(txtSaleSession.Text)
    data.SaleLot     = Trim(txtSaleLot.Text)
    data.SalePrice   = Trim(txtSalePrice.Text)
    data.SaleTax     = Trim(txtSaleTax.Text)
    data.SaleRecycle = Trim(txtSaleRecycle.Text)
    data.SaleTotal   = Trim(txtSaleTotal.Text)
    data.PaymentDate = Trim(txtPayDate.Text)

    WriteToSheetSales ws, data

    MsgBox data.CarNumber & " の売上を登録しました。", vbInformation
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
