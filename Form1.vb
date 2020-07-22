Imports System.Data
Imports System.Data.SqlClient
Imports System.Threading.Thread
Imports System.Globalization
Imports System.Windows.Forms.DataGridView
Imports System.Windows.Forms.DateTimePicker
Imports Microsoft.Office.Interop
Imports STROKESCRIBELib
Imports System.Runtime.InteropServices.COMException
Imports Microsoft.Office.Interop.Excel.XlInsertFormatOrigin
Imports QRCoder



Public Class frmPrintBom
    Dim xlApp As Excel.Application = New Excel.Application



    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        ShowData()
    End Sub
    Private Sub ShowData()
        Dim sql As String
        Dim datePickerFrom As String = dtpFrom.Value
        Dim datePickerTo As String = dtpTo.Value
        Dim txtMoVal As String = txtMO.Text
        Dim Crow As Integer


        Try
            conn.Open()
            sql = "SELECT FORMAT(MOH.RequestDate,'dd-MM-yyyy'),MOH.MfgOrderNo,MODT.ItemName,MODT.ItemQty,ML.LocationName "
            sql &= "FROM MfgOrderHeader MOH INNER JOIN MfgOrderDetail MODT  "
            sql &= "ON MOH.id = MODT.MfgOrderHeaderId "
            sql &= "INNER JOIN MasterItem MI ON MODT.MasterItemId = MI.Id "
            sql &= "INNER JOIN MasterLocation ML ON MODT.MasterLocationId = ML.Id "
            sql &= "WHERE MOH.RequestDate BETWEEN '" & datePickerFrom & "' AND '" & datePickerTo & "' AND MOH.MfgOrderNo LIKE '%" & txtMoVal & "%'"
            sql &= "AND MI.ItemGroup ='FG' "  'AND MI.ItemShortName = 'Roter Assy ' "
            sql &= "ORDER BY MODT.ItemQty DESC "
            Dim cmd As New SqlCommand(sql, conn)
            Dim ad As New SqlDataAdapter(cmd)
            Dim dt As New DataSet
            ad.Fill(dt, "Data")
            Crow = dt.Tables("Data").Rows.Count - 1
            conn.Close()
            Dim dataTable As New DataTable
            Dim dataRow As DataRow
            dataTable.Clear()
            dataTable.Columns.Add("No")
            dataTable.Columns.Add("Date")
            dataTable.Columns.Add("MfgOrder")
            dataTable.Columns.Add("ItemName")
            dataTable.Columns.Add("ItemQty")
            dataTable.Columns.Add("LocationName")

            For i As Integer = 0 To Crow
                dataRow = dataTable.NewRow
                dataRow("No") = i + 1
                dataRow("Date") = dt.Tables("Data").Rows(i)(0).ToString
                dataRow("MfgOrder") = dt.Tables("Data").Rows(i)(1).ToString
                dataRow("ItemName") = dt.Tables("Data").Rows(i)(2).ToString
                dataRow("ItemQty") = dt.Tables("Data").Rows(i)(3).ToString
                dataRow("LocationName") = dt.Tables("Data").Rows(i)(4).ToString
                dataTable.Rows.Add(dataRow)
            Next
            DataGri.DataSource = dataTable
            dt = Nothing
            ad = Nothing
            cmd = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub exportExcel(ByVal moOrder As String, ByVal qtyOrder As Integer)
        Dim sql As String, sqli As String
        Dim datePickerFrom As String = dtpFrom.Value
        Dim datePickerTo As String = dtpTo.Value
        Dim Crow As Integer
        Dim countBear As Integer
        Dim numlotn As Integer

        'getValue for Method 
        Dim moValue As String = moOrder
        Dim qtyValue As Integer = qtyOrder
        Dim mIo As String, mODTlineno As String, iTODlineno As String
        Dim mIitemno As String, mohHead As String, miLocationcode As String
        Dim qrValue As String
        Dim qtyValuedot As Double = qtyOrder
        Dim errorMo As String
        'Value Show Hearder
        Dim duedate As String, model As String, itemNo As String, itemGroup As String, local As String, idStruc As String
        'get Value Row in Excel
        Dim rowLot As Short = 5
        Dim rowExcel As Short = 6
        Dim rowValExcel As Short = 4


        'Value for Qty for SQL 
        Dim qtySQl As Integer
        Dim arrQty(5) As Integer
        Dim arri As Short = 0
        Dim resutlQty As Integer

        Dim numLot As Integer
        Dim x As Short = 1
        Dim formats As String = "dd/MM/yyyy"
    
        ' Try
        conn.Open()
        sql = "SELECT FORMAT(MOH.RequestDate,'dd/MM/yyyy') AS DueDate,MOH.MfgOrderNo,MI.ItemNo, "
        sql &= "MODT.ItemName,MPL.ProductionLineName,ML.LocationName,MI.LotSize,MI.ItemName,MI.ModelNo,MI.Id,ITOH.TransferOrderNo,MODT.MfgOrderLineNo,ITOD.TransferOrderLineNo,ML.LocationCode, "
        sql &= "CASE WHEN EnumDayNightId = 1 THEN 'Day' WHEN EnumDayNightId = 2 THEN 'Night' END AS Shift, "
        sql &= "MI.ItemGroup "
        sql &= "FROM MfgOrderHeader MOH "
        sql &= "INNER JOIN MfgOrderDetail MODT ON MOH.id = MODT.MfgOrderHeaderId "
        sql &= "INNER JOIN MasterLocation ML ON MODT.MasterLocationId = ML.Id "
        sql &= "INNER JOIN MasterItem MI ON MODT.MasterItemId = MI.Id "
        sql &= "INNER JOIN MasterStructure2 MS ON MI.Id = MS.MasterParentItemId "
        sql &= "INNER JOIN MasterproductionLine MPL ON MODT.MasterproductionLineId = MPL.Id  "
        sql &= "INNER JOIN InventoryTransferOrderHeader ITOH ON ITOH.MfgOrderDetailId = MODT.Id  "
        sql &= "INNER JOIN InventoryTransferOrderDetail ITOD ON ITOD.InventoryTransferOrderHeaderId = ITOH.Id "

        sql &= "WHERE MOH.MfgOrderNo = '" & moValue & "' AND MI.ItemGroup = 'FG' AND ITOD.ItemGroup ='RM' "

        'sqli = "DECLARE  @Idmaster AS int "
        'sqli &= "SELECT @Idmaster = MI.Id FROM MasterItem MI "
        'sqli &= "INNER JOIN MfgOrderDetail MODT ON MI.Id = MODT.MasterItemId "
        'sqli &= "INNER JOIN MfgOrderHeader MOH ON  MOH.id = MODT.MfgOrderHeaderId "
        'sqli &= " WHERE  MOH.MfgOrderNo = '" & moValue & "' "



        Dim cmd As New SqlCommand(sql, conn)
        Dim ad As New SqlDataAdapter(cmd)
        Dim dt As New DataSet

        ad.Fill(dt, "Data")




        Crow = dt.Tables("Data").Rows.Count


        If Crow < 1 Then

            MsgBox("ไม่สามารถพิมพ์ข้อมูลได้ เนื่องจากไม่มีข้อมูลที่อยู่ในกลุ่ม BALL BEARING ")
            conn.Close()
            xlApp.Workbooks.Close()
            xlApp.Quit()
        Else
            'get value Calculate 
            qtySQl = CInt(dt.Tables("Data").Rows(0)(6).ToString)
            'get Value Header document
            duedate = dt.Tables("Data").Rows(0)(0).ToString
            model = dt.Tables("Data").Rows(0)(3).ToString
            itemNo = dt.Tables("Data").Rows(0)(2).ToString
            itemGroup = dt.Tables("Data").Rows(0)(8).ToString
            local = dt.Tables("Data").Rows(0)(4).ToString



            dt = Nothing
            ad = Nothing
            cmd = Nothing
            conn.Close()

            ''''''Export Excel ''''''


            xlApp.Workbooks.Add()



            ' resutlQty = (qtyValue / qtySQl) * 3
            numLot = qtyValue / qtySQl
            If numLot < 1 Then

                resutlQty = 1
                numLot = 1
                numlotn = 1
            ElseIf numLot < 2 Then



                resutlQty = numLot

            Else

                resutlQty = numLot * 16.2

            End If

            ' prgb1.Minimum = 1
            ' prgb1.Maximum = resutlQty


            'load tab

            ' Loop Until qtySQl < qtyValue
            'End Calculate Qty

            'Header Column



            'Line of excel
            ' lineRow(rowExcel + i)
            'End Set Line Column

            'Header Column

            conn.Open()
            Dim cmds As New SqlCommand(sql, conn)
            Dim ads As New SqlDataAdapter(cmds)
            Dim dts As New DataSet
            ads.Fill(dts, "Datas")

            mohHead = dts.Tables("Datas").Rows(0)(1).ToString()

            idStruc = dts.Tables("Datas").Rows(0)(9).ToString
            mIo = dts.Tables("Datas").Rows(0)(10).ToString
            mODTlineno = dts.Tables("Datas").Rows(0)(11).ToString
            iTODlineno = dts.Tables("Datas").Rows(0)(12).ToString
            miLocationcode = dts.Tables("Datas").Rows(0)(13).ToString
            'MsgBox(mIo)
            'MsgBox(mODTlineno)
            'MsgBox(iTODlineno)
            sqli = "SELECT  MI.id,MI.ItemNo,MI.ModelNo,MI.ItemShortName,MI.ItemName,MI.ItemGroup "
            sqli &= "FROM MasterStructure2  MS "
            sqli &= "INNER JOIN MasterItem MI ON  MS.MasterChildItemId = MI.Id "
            sqli &= "WHERE MS.MasterParentItemId = '" & idStruc & "' AND MI.ItemShortName = 'BALL BEARING' "

            Dim cmdz As New SqlCommand(sqli, conn)
            Dim adz As New SqlDataAdapter(cmdz)
            Dim dtz As New DataSet
            adz.Fill(dtz, "Dataz")


            conn.Close()

            'value qr code

            'MsgBox(qrValue)

            Dim rowCell As Integer
            Dim lot As Integer

            countBear = dtz.Tables("Dataz").Rows.Count
            errorMo = dts.Tables("Datas").Rows(0)(1).ToString

            If countBear >= 1 Then

                mIitemno = dtz.Tables("Dataz").Rows(0)(1).ToString
                qrValue = mIitemno & "	" & qtySQl.ToString("###.000") & "	" & miLocationcode & "	" & mIo & "	" & iTODlineno & "	" & mohHead & "	" & mODTlineno

                For lot = 1 To numLot



                    For rowCell = 0 To resutlQty

                        ''''''Export Excel ''''''



                        xlApp.Cells(rowCell + 1, 1).Value = "PrintDate :" & Date.Today.ToString(formats)
                        xlApp.Range(xlApp.Cells(rowCell + 1, 1), xlApp.Cells(rowCell + 1, 1)).ColumnWidth = 19

                        xlApp.Cells(rowCell + 1, 2).Value = "PRODUCTION ORDER & BOM (BEARING)"
                        xlApp.Cells(rowCell + 1, 2).Font.Size = 12
                        xlApp.Cells(rowCell + 1, 2).Font.Bold = True
                        xlApp.Cells(rowCell + 1, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        xlApp.Cells(rowCell + 8, 4).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                        xlApp.Range(xlApp.Cells(rowCell + 1, 2), xlApp.Cells(rowCell + 1, 3)).MergeCells = True
                        xlApp.Range(xlApp.Cells(rowCell + 1, 2), xlApp.Cells(rowCell + 1, 2)).ColumnWidth = 26
                        xlApp.Range(xlApp.Cells(rowCell + 1, 3), xlApp.Cells(rowCell + 1, 3)).ColumnWidth = 16
                        xlApp.Range(xlApp.Cells(rowCell + 1, 5), xlApp.Cells(rowCell + 1, 5)).ColumnWidth = 9
                        xlApp.Range(xlApp.Cells(rowCell + 1, 4), xlApp.Cells(rowCell + 1, 4)).ColumnWidth = 11

                        xlApp.Cells(rowCell + 1, 5).Value = "Pallet :" & lot & "/ " & numLot
                        xlApp.Cells(rowCell + 4, 3).Value = "Plan Qty :  "
                        xlApp.Cells(rowCell + 4, 4).Value = qtyValue.ToString("0,00")
                        xlApp.Cells(rowCell + 8, 3).Font.Size = 13



                        xlApp.Cells(rowCell + 4, 5).Value = "  Pcs."
                        xlApp.Cells(rowCell + 3, 1).Value = "Mfg Order No :  "
                        xlApp.Cells(rowCell + 3, 2).Value = dts.Tables("Datas").Rows(0)(1).ToString
                        xlApp.Cells(rowCell + 3, 3).Value = "Drawing No :  "
                        xlApp.Cells(rowCell + 3, 4).Value = itemNo
                        xlApp.Cells(rowCell + 4, 1).Value = "Part Name :  "
                        xlApp.Cells(rowCell + 4, 2).Value = dts.Tables("Datas").Rows(0)(7).ToString
                        xlApp.Cells(rowCell + 6, 1).Value = "Line :  "
                        xlApp.Cells(rowCell + 6, 2).Value = dts.Tables("Datas").Rows(0)(4).ToString

                        xlApp.Cells(rowCell + 5, 1).Value = "Model :  "
                        xlApp.Cells(rowCell + 5, 2).Value = dts.Tables("Datas").Rows(0)(8).ToString
                        xlApp.Cells(rowCell + 5, 3).Value = "Production Date :  "
                        xlApp.Cells(rowCell + 5, 4).Value = duedate
                        ' xlApp.Cells(rowCell + 6, 3).Value = "Pallet Qty :  "
                        'xlApp.Cells(rowCell + 6, 4).Value = qtySQl & "   Pcs."
                        xlApp.Range(xlApp.Cells(rowCell + 7, 1), xlApp.Cells(rowCell + 7, 2)).MergeCells = True
                        xlApp.Cells(rowCell + 7, 1).Value = "*** Request Parts \ Material "
                        xlApp.Cells(rowCell + 7, 1).Font.Bold = True
                        xlApp.Cells(rowCell + 8, 1).Value = "Item Code   "
                        'xlApp.Range(xlApp.Cells(rowCell + 8, 1), xlApp.Cells(rowCell + 8, 2)).MergeCells = True
                        xlApp.Cells(rowCell + 8, 2).Value = dtz.Tables("Dataz").Rows(0)(1).ToString
                        xlApp.Cells(rowCell + 8, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                        xlApp.Cells(rowCell + 7, 3).Value = "Pallet Qty"

                        If numlotn = 1 Then

                            xlApp.Cells(rowCell + 8, 3).Value = qtyValue & "  PCS."

                        Else

                            xlApp.Cells(rowCell + 8, 3).Value = qtySQl & "  PCS."

                        End If
                        xlApp.Cells(rowCell + 8, 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        xlApp.Cells(rowCell + 8, 3).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                        xlApp.Cells(rowCell + 7, 4).Value = "QR CODE"
                        xlApp.Range(xlApp.Cells(rowCell + 8, 3), xlApp.Cells(rowCell + 10, 3)).MergeCells = True
                        xlApp.Range(xlApp.Cells(rowCell + 7, 4), xlApp.Cells(rowCell + 7, 5)).MergeCells = True
                        xlApp.Range(xlApp.Cells(rowCell + 9, 4), xlApp.Cells(rowCell + 14, 5)).MergeCells = True
                        xlApp.Range(xlApp.Cells(rowCell + 11, 1), xlApp.Cells(rowCell + 12, 1)).MergeCells = True
                        xlApp.Range(xlApp.Cells(rowCell + 13, 1), xlApp.Cells(rowCell + 14, 1)).MergeCells = True
                        xlApp.Range(xlApp.Cells(rowCell + 11, 2), xlApp.Cells(rowCell + 12, 3)).MergeCells = True
                        xlApp.Range(xlApp.Cells(rowCell + 13, 2), xlApp.Cells(rowCell + 14, 3)).MergeCells = True
                        xlApp.Cells(rowCell + 11, 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                        xlApp.Cells(rowCell + 11, 1).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                        xlApp.Cells(rowCell + 13, 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                        xlApp.Cells(rowCell + 13, 1).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                        xlApp.Cells(rowCell + 7, 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        xlApp.Cells(rowCell + 7, 4).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

                        xlApp.Cells(rowCell + 9, 1).Value = "Part Name   "
                        ' xlApp.Range(xlApp.Cells(rowCell + 9, 1), xlApp.Cells(rowCell + 9, 2)).MergeCells = True
                        xlApp.Cells(rowCell + 9, 2).Value = dtz.Tables("Dataz").Rows(0)(4).ToString
                        xlApp.Cells(rowCell + 10, 1).Value = "Model      "
                        'xlApp.Range(xlApp.Cells(rowCell + 10, 1), xlApp.Cells(rowCell + 10, 2)).MergeCells = True
                        xlApp.Cells(rowCell + 10, 2).Value = dtz.Tables("Dataz").Rows(0)(2).ToString
                        xlApp.Cells(rowCell + 11, 1).Value = "REQUEST SING"
                        xlApp.Cells(rowCell + 13, 1).Value = "ISSUE SING"
                        xlApp.Cells(rowCell + 16, 5).Value = "OCT-PC-FM047 Rev.00"
                        xlApp.Cells(rowCell + 16, 5).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        xlApp.Cells(rowCell + 16, 5).Font.Size = 8
                        'xlApp.Cells(rowCell + 8 + rowValExcel, 4).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                        'xlApp.Range(xlApp.Cells(rowCell + 7, 1), xlApp.Cells(rowCell + 7, 2)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                        'xlApp.Range(xlApp.Cells(rowCell + 7, 1), xlApp.Cells(rowCell + 7, 2)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                        'xlApp.Range(xlApp.Cells(rowCell + 7, 1), xlApp.Cells(rowCell + 7, 2)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                        'xlApp.Range(xlApp.Cells(rowCell + 7, 1), xlApp.Cells(rowCell + 7, 2)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous

                        xlApp.Cells(rowCell + 7, 3).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Cells(rowCell + 7, 3).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Cells(rowCell + 7, 3).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Cells(rowCell + 7, 3).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous

                        xlApp.Range(xlApp.Cells(rowCell + 7, 4), xlApp.Cells(rowCell + 7, 5)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                        ' xlApp.Range(xlApp.Cells(rowCell + 7, 4), xlApp.Cells(rowCell + 7, 5)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 7, 4), xlApp.Cells(rowCell + 7, 5)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 7, 4), xlApp.Cells(rowCell + 7, 5)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous


                        xlApp.Range(xlApp.Cells(rowCell + 7, 4), xlApp.Cells(rowCell + 10, 4)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 7, 4), xlApp.Cells(rowCell + 10, 4)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 7, 4), xlApp.Cells(rowCell + 10, 4)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 7, 4), xlApp.Cells(rowCell + 10, 4)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous

                        '  xlApp.Range(xlApp.Cells(rowCell + 8, 4), xlApp.Cells(rowCell + 14, 5)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                        '   xlApp.Range(xlApp.Cells(rowCell + 8, 4), xlApp.Cells(rowCell + 14, 5)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 8, 4), xlApp.Cells(rowCell + 14, 5)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 8, 4), xlApp.Cells(rowCell + 14, 5)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous

                        xlApp.Range(xlApp.Cells(rowCell + 8, 1), xlApp.Cells(rowCell + 10, 2)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 8, 1), xlApp.Cells(rowCell + 10, 2)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 8, 1), xlApp.Cells(rowCell + 10, 2)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 8, 1), xlApp.Cells(rowCell + 10, 2)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous

                        xlApp.Range(xlApp.Cells(rowCell + 8, 3), xlApp.Cells(rowCell + 10, 3)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 8, 3), xlApp.Cells(rowCell + 10, 3)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 8, 3), xlApp.Cells(rowCell + 10, 3)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 8, 3), xlApp.Cells(rowCell + 10, 3)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous

                        xlApp.Range(xlApp.Cells(rowCell + 11, 1), xlApp.Cells(rowCell + 12, 1)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 11, 1), xlApp.Cells(rowCell + 12, 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 11, 1), xlApp.Cells(rowCell + 12, 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 11, 1), xlApp.Cells(rowCell + 12, 1)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous

                        xlApp.Range(xlApp.Cells(rowCell + 11, 2), xlApp.Cells(rowCell + 12, 3)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 11, 2), xlApp.Cells(rowCell + 12, 3)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 11, 2), xlApp.Cells(rowCell + 12, 3)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 11, 2), xlApp.Cells(rowCell + 12, 3)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous

                        xlApp.Range(xlApp.Cells(rowCell + 13, 1), xlApp.Cells(rowCell + 14, 1)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 13, 1), xlApp.Cells(rowCell + 14, 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 13, 1), xlApp.Cells(rowCell + 14, 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 13, 1), xlApp.Cells(rowCell + 14, 1)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous

                        xlApp.Range(xlApp.Cells(rowCell + 13, 2), xlApp.Cells(rowCell + 14, 3)).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 13, 2), xlApp.Cells(rowCell + 14, 3)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 13, 2), xlApp.Cells(rowCell + 14, 3)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                        xlApp.Range(xlApp.Cells(rowCell + 13, 2), xlApp.Cells(rowCell + 14, 3)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous


                        xlApp.Range(xlApp.Cells(rowCell + 16, 1), xlApp.Cells(rowCell + 16, 5)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDot


                        generateQrCode(rowCell, qrValue)
                        rowCell += 16

                        Select Case rowCell

                            Case 50
                                xlApp.Range(xlApp.Cells(rowCell, 1), xlApp.Cells(rowCell, 5)).Borders.Color = Color.Transparent
                                rowCell -= 1

                            Case 100
                                xlApp.Range(xlApp.Cells(rowCell, 1), xlApp.Cells(rowCell, 5)).Borders.Color = Color.Transparent

                                rowCell -= 1

                            Case 150
                                xlApp.Range(xlApp.Cells(rowCell, 1), xlApp.Cells(rowCell, 5)).Borders.Color = Color.Transparent

                                rowCell -= 1

                            Case 200
                                xlApp.Range(xlApp.Cells(rowCell, 1), xlApp.Cells(rowCell, 5)).Borders.Color = Color.Transparent

                                rowCell -= 1

                        End Select



                        'x += 1
                        'String.Format((rowCell / resutlQty) * 100) & "%"
                        lot += 1
                        '  prgb1.Value = rowCell
                        Label3.Text = "Creating Excel File....    " & ((rowCell / resutlQty) * 100).ToString("###") & "%"
                        Threading.Thread.Sleep(0)
                        'Application.DoEvents()
                        Label3.ForeColor = Color.BlueViolet
                        Label3.Visible = True
                        ' prgb1.Visible = False

                    Next



                    xlApp.Visible = True
                    '  prgb1.Visible = False
                    Label3.Text = "Creating Excel File....  100 %"
                    Label3.ForeColor = Color.LimeGreen
                    Label3.Visible = True

                Next
                dts = Nothing
                ads = Nothing
                cmds = Nothing

                ' Catch ex As Exception
                'MsgBox("Error" & ex.Message)
                ' End Try

            Else

                MsgBox(errorMo & "  ไม่สามารถพิมพ์ข้อมูลได้ เนื่องจากไม่มีข้อมูลที่อยู่ในกลุ่ม BALL BEARING ")
                prgb1.Value = rowCell
                Label3.Text = "Creating Excel File....    " & ((rowCell / resutlQty) * 100).ToString("###") & "%"
                Threading.Thread.Sleep(0)

                'prgb1.Visible = False
                Label3.Visible = False
                xlApp.Quit()
            End If

        End If

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnExcel.Click




        If DataGri.Rows.Count < 1 Then


            MsgBox("ไม่พบหมายเลย MO")
            xlApp.Workbooks.Close()
            xlApp.Quit()
        Else

            Dim CurRow As Integer = DataGri.CurrentRow.Index
            Dim mfgorder As String = DataGri.Rows(CurRow).Cells("MfgOrder").Value.ToString()
            Dim qtyOrder As String = DataGri.Rows(CurRow).Cells("ItemQty").Value.ToString()
            exportExcel(mfgorder, qtyOrder)


        End If
    End Sub

    Private Sub frmPrintBom_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        prgb1.Visible = False
        Label3.Visible = False
    End Sub

    Private Sub lineRow(ByVal rowExcel As Integer)
        'Set Line Column

        'Line of excel
        For i As Integer = 1 To 9
            With xlApp.Cells(rowExcel, i)
                .Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            End With
        Next

        'End Set Line Column

        'Header Column
    End Sub

    Private Function ExcelCellHorizontalAlignmentType() As Object
        Throw New NotImplementedException
    End Function

    Private Function ExcelCellVerticalAlignmentType() As Object
        Throw New NotImplementedException
    End Function

    Private Sub generateQrCode(ByVal row As Integer, qrValue As String)
        Dim qrcode As New QRCodeGenerator
        Dim data = qrcode.CreateQrCode(qrValue, QRCodeGenerator.ECCLevel.Q)
        Dim code As New QRCode(data)
        
        'picQRcode.Image = code.GetGraphic(6)
        'Export Excel
        Dim bm As Bitmap

        bm = New Bitmap(code.GetGraphic(3))

        System.Windows.Forms.Clipboard.SetDataObject(bm, False)
        
        'tImer1.Interval = 1000

        ' tImer1.Start()
        ' Threading.Thread.Sleep(3000)

        Delay(2)
        xlApp.Cells(row + 8, 4).PasteSpecial("Bitmap")

        Try

        Catch xlApp As Exception


        End Try



    End Sub

    Private Sub Delay(ByVal DelayInSeconds As Integer)
        Dim ts As TimeSpan
        Dim targetTime As DateTime = DateTime.Now.AddSeconds(DelayInSeconds)
        Do
            ts = targetTime.Subtract(DateTime.Now)
            Application.DoEvents() ' keep app responsive
            System.Threading.Thread.Sleep(50) ' reduce CPU usage
        Loop While ts.TotalSeconds > 0
    End Sub

    
End Class



