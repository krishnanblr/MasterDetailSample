Imports System.Data.OleDb

Module modMasterDtlsRpt
    Public para As String
    '==============FOR REPORT

    Dim _CurrentPage As Integer = 0
    'font 
    Dim MyHeaderFont As Font = New Font("Times New Roman", 15, FontStyle.Bold)
    Dim MyTitleFont As Font = New Font("Times New Roman", 10, FontStyle.Bold)
    Dim MyRptFont As Font = New Font("Garamond", 10, FontStyle.Regular)

    Dim leftMargins As Double = 0
    Dim rightMargin As Double = 0
    Dim topMargins As Double = 0
    Dim bottomMargin As Double = 0
    Dim pageHeight As Double = 0
    Dim pageWidth As Double = 0
    Dim LeftMargin As Double, TopMargin As Double, BtmMargin As Double
    Dim SlNo As Integer = 0

    Dim oMD As OleDbDataReader
    Dim dtMD As DataTable

    Dim TotAmt As Double = 0, TotPack As Double = 0
    Dim sPONo As String, sPODate As String, sMRPNo As String, sLedName As String, sLedAdd As String, sLedCity As String, sLedState As String, sLedCountry As String
    Dim sAmendNo As String, sAmendDate As DateTime
    Dim sLedCSTTIN As String, sPONote As String, sCtc As String, sOpenPara As String, sClosePara As String, sLedID As String
    Dim sLedGSTIN As String
    Dim RcdNo As Integer, dPOAmt As Double, dPOQty As Double
    Dim PartNo_LineNo As Integer = 0
    Dim k As Integer = 0, j As Integer = 0, m As Integer


    '==============END OF REPORT
    Public Sub Print_Report(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        Dim X As Double, x1 As SizeF, x2 As SizeF, PgWidth As Double, PgHeight As Double
        Dim Y As Double, XA1 As Double
        Dim XPos As Double

        Dim HeaderText As String
        Dim MaxLine As Single

        Dim g As Graphics = e.Graphics
        Dim myBrush As Brush = New SolidBrush(Color.Black)
        Dim dots As Pen = New Pen(myBrush, 0.007)


        Dim RightMargins As Double
        Dim CXPos As Double, CYPos As Double

        dots.DashStyle = Drawing2D.DashStyle.Solid
        e.Graphics.PageUnit = GraphicsUnit.Inch

        PgWidth = (e.PageSettings.PaperSize.Width / 100)
        PgHeight = e.PageSettings.PaperSize.Height / 100

        leftMargins = e.MarginBounds.Left
        rightMargin = e.MarginBounds.Right
        topMargins = e.MarginBounds.Top
        bottomMargin = e.MarginBounds.Bottom
        pageHeight = bottomMargin - topMargins
        pageWidth = rightMargin - leftMargins

        Dim headerHeight As Double = MyTitleFont.GetHeight(g)
        Dim footerHeight As Double = MyTitleFont.GetHeight(g)
        Dim ReportheaderR As New RectangleF(LeftMargin, TopMargin, PgWidth, PgHeight)
        Dim bodyR As New RectangleF(leftMargins, ReportheaderR.Bottom, pageWidth, pageHeight - ReportheaderR.Height - footerHeight)
        Dim bodyFontHeight As Double = footerHeight
        Dim linesPerPage As Double = Convert.ToInt32(bodyR.Height / MyTitleFont.Height) - 2
        Dim c As Integer, sArray As Array
        Dim MaxPrintSize As Double = PgHeight - (PgHeight * 0.07)

        Y = 0.1
        BtmMargin = 0.5
        LeftMargin = 0.5
        RightMargins = (PgWidth) - 0.5
        TopMargin = 0.2
        'Header Info
        CXPos = LeftMargin : CYPos = TopMargin
        HeaderText = "NORTHWIND MASTER DETAIL REPORT"

        Try
            'sl, orderid, productid, prodcut, unitid, qty disc  

            g.DrawString(HeaderText, MyHeaderFont, Brushes.Black, CXPos, CYPos)
            'CXPos = LeftMargin
            CYPos = CYPos + 0.25


            g.DrawLine(dots, Convert.ToSingle(LeftMargin), Convert.ToSingle(CYPos), Convert.ToSingle(PgWidth - LeftMargin), Convert.ToSingle(CYPos))
            MaxLine = 16

            CXPos = CXPos : CYPos = CYPos
                g.DrawString("Sl", MyTitleFont, Brushes.Black, CXPos, CYPos)

            CXPos += 0.2 : CYPos = CYPos
            g.DrawString("Order ID", MyTitleFont, Brushes.Black, CXPos, CYPos)

            CXPos += 1 : CYPos = CYPos
            g.DrawString("Product ID", MyTitleFont, Brushes.Black, CXPos, CYPos)

            CXPos += 1 : CYPos = CYPos
            g.DrawString("Product Name", MyTitleFont, Brushes.Black, CXPos, CYPos)

            CXPos += 2 : CYPos = CYPos
            g.DrawString("Unit Price", MyTitleFont, Brushes.Black, CXPos, CYPos)

            CXPos += 1 : CYPos = CYPos
            g.DrawString("Quantity", MyTitleFont, Brushes.Black, CXPos, CYPos)

            CXPos += 1 : CYPos = CYPos
            g.DrawString("Discount", MyTitleFont, Brushes.Black, CXPos, CYPos)

            CXPos = LeftMargin : CYPos = CYPos + 0.2

            g.DrawLine(dots, Convert.ToSingle(LeftMargin), Convert.ToSingle(CYPos), Convert.ToSingle(PgWidth - LeftMargin), Convert.ToSingle(CYPos))
            dPOAmt = 0

            Do While MaxLine + 1 < linesPerPage And CYPos + 1 <= MaxPrintSize And k <= dtMD.Rows.Count - 1
                With dtMD.Rows(k)
                    RcdNo += 1
                    CXPos = LeftMargin : CYPos = CYPos
                    g.DrawString(RcdNo, MyRptFont, Brushes.Black, CXPos, CYPos)

                    CXPos += 0.2 : CYPos = CYPos

                    g.DrawString(.Item("orderid"), MyRptFont, Brushes.Black, CXPos, CYPos)

                    CXPos += 1 : CYPos = CYPos

                    g.DrawString(.Item("productid"), MyRptFont, Brushes.Black, CXPos, CYPos)

                    CXPos += 1 : CYPos = CYPos

                    g.DrawString(.Item("productname"), MyRptFont, Brushes.Black, CXPos, CYPos)
                    CXPos += 2 : CYPos = CYPos

                    g.DrawString(.Item("unitprice"), MyRptFont, Brushes.Black, CXPos, CYPos)
                    CXPos += 1 : CYPos = CYPos

                    g.DrawString(.Item("quantity"), MyRptFont, Brushes.Black, CXPos, CYPos)

                    CXPos += 1 : CYPos = CYPos

                    g.DrawString(.Item("discount"), MyRptFont, Brushes.Black, CXPos, CYPos)
                    CYPos += 0.25

                    MaxLine += 1
                    k += 1
                End With
            Loop


            If CYPos > MaxPrintSize Or k < dtMD.Rows.Count - 1 Then 'Or MaxLine + 1 > linesPerPage


                CXPos = 1 : CYPos += 0.2

                g.DrawString("Report by Shivaram", MyRptFont, Brushes.Black, CXPos, CYPos)
                CYPos = TopMargin : MaxLine = 0
                e.HasMorePages = True
            Else
                CXPos = 1 : CYPos += 0.2

                g.DrawString("Report by Shivaram", MyRptFont, Brushes.Black, CXPos, CYPos)
                e.HasMorePages = False
            End If
        Catch ex As Exception
            MessageBox.Show(Err.Description)
        End Try
    End Sub

    Public Sub BeginPrint_Report(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs)
        Dim sSQL As String

        _CurrentPage = 0 : PartNo_LineNo = 0
        j = 0
        m = 0
        TotAmt = 0 : TotPack = 0
        k = 0
        Dim con As New OleDbConnection(My.Settings.nwindConnectionString)

        Dim cmd As New OleDbCommand
        con.Open()
        cmd.Connection = con
        'sSQL = ""
        'cmd.CommandText = sSQL

        sSQL = "SELECT * from Invoices where customerid='" & para & "'"
        cmd.CommandText = sSQL
        'oMD = cmd.ExecuteReader


        Dim ds As New DataSet
        Dim da As New OleDbDataAdapter(cmd)
        da.Fill(ds, "Invoices")
        con.Close()
        dtMD = ds.Tables("Invoices")

        RcdNo = 0

    End Sub

    Public Sub QueryPageSettings(ByVal sender As Object, ByVal e As System.Drawing.Printing.QueryPageSettingsEventArgs)
        _CurrentPage += 1
    End Sub

End Module