Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop

Public Class clsExcelClass
    Private oRecordSet As SAPbobsCOM.Recordset
    Private Const strInvoiceNo As String = "kr49_renr"
    Private Const strCardCode As String = "kr49_kdnr"
    Private Const strCardName As String = "pm53_ktrbez"
    'Private Const strDocDate As String = "kr49_renr"
    Private Const strDocDate As String = "kr49_redat"
    'Private Const strDueDate As String = "kr49_edate"
    Private Const strDueDate As String = "faelldt"
    Private Const strProject As String = "kr49_abnr"
    Private Const strLineTotal As String = "kr49_netto"
    Private Const strDocTotal As String = "kr49_brutto"
    Private Const strLineNo As String = "kr49_erlkto"
    Private Const strAcctCode As String = "kr49_erlkto"
    Private Const strCurrency As String = "kr49_fwaeh"
    Private Const strDocType As String = "kr49_anga"

    Public Function getExcelTemplate(ByVal strPath As String, ByRef dtExcel As System.Data.DataTable) As Boolean
        Dim _retVal As Boolean = True
        Dim excel As Application = New Application
        Dim w As Workbook = excel.Workbooks.Open(strPath)
        Try
            Dim dr As DataRow

            ' Loop over all sheets.
            For i As Integer = 1 To w.Sheets.Count
                Dim sheet As Worksheet = w.Sheets(i)
                Dim r As Range = sheet.UsedRange
                Dim array(,) As Object = r.Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault)
                If array IsNot Nothing Then

                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim intURow As Integer = array.GetUpperBound(0)
                    Dim intUCol As Integer = array.GetUpperBound(1)

                    Dim intInvoiceNo, intCardCode, intCardName, intDocDate, intDueDate, intProject, intLineTotal, intDocTotal, _
                        intLineNo, intRevAcct, intCurrency, intDocType As Integer
                    For index As Integer = 3 To 3
                        For intCol As Integer = 1 To intUCol
                            Dim xRng As Excel.Range = CType(sheet.Cells(index, intCol), Excel.Range)
                            Dim strStyle As String = xRng.Value()
                            Select Case strStyle
                                Case strInvoiceNo
                                    intInvoiceNo = intCol
                                Case strCardCode
                                    intCardCode = intCol
                                Case strCardName
                                    intCardName = intCol
                                Case strDocDate
                                    intDocDate = intCol
                                Case strDueDate
                                    intDueDate = intCol
                                Case strProject
                                    intProject = intCol
                                Case strLineTotal
                                    intLineTotal = intCol
                                Case strDocTotal
                                    intDocTotal = intCol
                                Case strCurrency
                                    intCurrency = intCol
                                Case strDocType
                                    intDocType = intCol
                                Case strLineNo, strAcctCode
                                    intLineNo = intCol
                                    intRevAcct = intCol
                            End Select
                        Next
                    Next

                    For intRow As Integer = 4 To intURow

                        '  MsgBox(array(intRow, intInvoiceNo))
                        Dim strDocNum As String = array(intRow, intInvoiceNo)

                        If strDocNum = Nothing Or strDocDate = "" Then

                        Else
                            dr = dtExcel.NewRow()
                            dr("InvoiceNo") = array(intRow, intInvoiceNo)  'InvoiceNo
                            dr("CardCode") = array(intRow, intCardCode)  'CardCode
                            dr("CardName") = array(intRow, intCardName).ToString.Replace("'", "") 'CardName
                            dr("DocDate") = array(intRow, intDocDate) 'DocDate
                            dr("DueDate") = array(intRow, intDueDate) 'DueDate
                            dr("Project") = array(intRow, intProject) 'Project
                            If IsDBNull(array(intRow, intLineTotal)) = True Then
                                dr("LineTotal") = 0 ' array(intRow, intLineTotal) 'LineTotal
                            Else
                                Try
                                    dr("LineTotal") = array(intRow, intLineTotal) 'LineTotal
                                Catch ex As Exception
                                    dr("LineTotal") = 0
                                End Try

                            End If
                            Try
                                If IsDBNull(array(intRow, intDocTotal)) = True Then
                                    dr("DocTotal") = 0
                                Else
                                    dr("DocTotal") = array(intRow, intDocTotal) 'DocTotal
                                End If
                            Catch ex As Exception
                                dr("DocTotal") = 0
                            End Try

                            dr("Line") = array(intRow, intLineNo) 'Line
                            dr("RevAcct") = array(intRow, intLineNo) 'Revenue Account
                            dr("Currency") = array(intRow, intCurrency) 'Currency
                            dr("DocType") = array(intRow, intDocType) 'Document Type
                            dtExcel.Rows.Add(dr)
                        End If
                    Next

                End If
            Next
            w.Close()
            excel.Quit()
            Return _retVal
        Catch ex As Exception
            w.Close()
            excel.Quit()
            Throw ex
        Finally
            ReleaseComObject(w)
            ReleaseComObject(excel)
        End Try
    End Function

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        End Try
    End Sub

End Class
