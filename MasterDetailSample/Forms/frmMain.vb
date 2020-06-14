Imports System.Drawing.Printing

Public Class frmMain
    'Developed by: Gani Weecom
    'Email:ganiweecom@yahoo.com
    Dim masterDetail As MasterControl


    'for print
    Public CurrentCopyName As String = "ORIGINAL FOR RECIPIENT"
    Public WithEvents print_document As New PrintDocument

    Sub clearFields()
        panelView.Controls.Clear()
        masterDetail = Nothing
        Refresh()
    End Sub
    Sub loadData()
        clearFields()
        Me.OrderReportsTableAdapter.Fill(Me.NwindDataSet.OrderReports)
        Me.InvoicesTableAdapter.Fill(Me.NwindDataSet.Invoices)
        Me.CustomersTableAdapter.Fill(Me.NwindDataSet.Customers)
        createMasterDetailView()

        If masterDetail.Rows.Count > 0 Then
            btnReport.Enabled = True
        End If

    End Sub
    Sub createMasterDetailView()
        masterDetail = New MasterControl(NwindDataSet)
        panelView.Controls.Add(masterDetail)
        masterDetail.setParentSource(NwindDataSet.Customers.TableName, "CustomerID")
        masterDetail.childView.Add(NwindDataSet.OrderReports.TableName, "Orders")
        masterDetail.childView.Add(NwindDataSet.Invoices.TableName, "Invoices")
    End Sub
    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        loadData()
    End Sub

    Private Sub btnReport_Click(sender As Object, e As EventArgs) Handles btnReport.Click

        loadReport()
    End Sub
    Sub loadReport()
        Dim PrintPreviewDlg As New PrintPreviewDialog

        'adding toolbar for different actions

        Dim PrintPreviewToolbar As ToolStrip = CType(PrintPreviewDlg.Controls(1), ToolStrip)
        Dim CustomPrintButton As ToolStripButton = New ToolStripButton
        DirectCast(DirectCast(PrintPreviewDlg.Controls(1), ToolStrip).Items(0), ToolStripButton).Visible = False 'Hides the original print button from the print preview dialogs toolstrip
        DirectCast(DirectCast(PrintPreviewDlg.Controls(1), ToolStrip).Items(9), ToolStripButton).Visible = False 'hides the original close button

        CustomPrintButton.Name = "Print"
        CustomPrintButton.ImageIndex = 0
        CustomPrintButton.ToolTipText = "Prints current report"


        'Add a Click event handler to the custom button (so it calls a function when clicked)
        AddHandler CustomPrintButton.Click, AddressOf PrintPreview_Click
        'AddHandler CustomCloseButton.Click, AddressOf btnClose_Click

        PrintPreviewToolbar.Items.Add(CustomPrintButton) 'Add the custom button to the toolstrip

        PrintPreviewDlg.Name = "PrintDocument"
        PrintPreviewDlg.Document = PreparePrintDocument()

        ' PrintPreviewDlg.MdiParent = MdiParent
        With PrintPreviewDlg

            .TopLevel = False

            Form1.Panel1.Controls.Add(PrintPreviewDlg)
            .Dock = DockStyle.Fill
            .AutoScaleMode = AutoScaleMode.Dpi
            .FormBorderStyle = FormBorderStyle.None

            .AutoScroll = True
            .BringToFront()
            .Show()

        End With
        Form1.ShowDialog()
        PrintPreviewDlg.BringToFront()
        PrintPreviewDlg.Show()
        PrintPreviewDlg.PrintPreviewControl.Zoom = 1


    End Sub
    Public Function PreparePrintDocument() As PrintDocument
        PreparePrintDocument = Nothing
        ' Make the PrintDocument object.
        print_document = New PrintDocument

        Dim pgS As System.Drawing.Printing.PageSettings = New System.Drawing.Printing.PageSettings

        'defining paper size, here custom size is defined.

        Dim sPaper As New System.Drawing.Printing.PaperSize With {
            .RawKind = Printing.PaperKind.Custom,
            .Height = 1170,
            .Width = 830
        }
        print_document.DefaultPageSettings.PaperSize = sPaper
        print_document.DefaultPageSettings.Landscape = False
        para = masterDetail.SelectedCells(0).Value
        AddHandler print_document.BeginPrint, AddressOf BeginPrint_Report
        AddHandler print_document.QueryPageSettings, AddressOf QueryPageSettings
        AddHandler print_document.PrintPage, AddressOf Print_Report
        print_document.DocumentName = "Northwind Master Detail Report"

        Return print_document
    End Function

    'to set the printer defaults
    Public Sub SetPrinter(ByVal prtName As String)
        Try
            Dim s As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser
            Dim m As Microsoft.Win32.RegistryKey
            m = s.OpenSubKey("Software")
            m = m.OpenSubKey("Microsoft")
            m = m.OpenSubKey("Windows NT")
            m = m.OpenSubKey("CurrentVersion")
            m = m.OpenSubKey("Windows", True)
            m.SetValue("Device", prtName & ",winspool,FILE:")
            m.Flush()
            m.Close()
            s.Close()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub PrintPreview_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim prndlg As New PrintDialog
        With prndlg
            .UseEXDialog = True
            .AllowCurrentPage = True
            .AllowPrintToFile = True
            .AllowSelection = True
            .AllowSomePages = True
            .ShowHelp = True
            .ShowNetwork = True

        End With
        Dim result As DialogResult = prndlg.ShowDialog()
        Dim r = prndlg.PrinterSettings.Copies

        Dim strPrnName As String = prndlg.PrinterSettings.PrinterName.ToString
        SetPrinter(strPrnName)

        If (result = System.Windows.Forms.DialogResult.OK) Then

            For p = 0 To r - 1
                Select Case p
                    Case 0
                        CurrentCopyName = "ORIGINAL FOR RECIPIENT"
                    Case 1
                        CurrentCopyName = "DUPLICATE FOR SUPPLIER/TRANSPORT"
                    Case 2
                        CurrentCopyName = "TRIPLICATE FOR SUPPLIER"
                    Case 3
                        CurrentCopyName = "QUADRUPLICATE - EXTRA COPY"
                    Case 4
                        CurrentCopyName = "QUINTUPLET"
                    Case 5
                        CurrentCopyName = "SEXTUPLET"
                    Case 6
                        CurrentCopyName = "SEPTUPLET"

                        '5-quintuplet
                        '6-sextuplet
                        '7-septuplet
                        '8-octuplet
                        '9-nonuplet
                        '10-decuplet
                End Select
                print_document.Print()

            Next

            'End If

        End If
    End Sub



End Class
