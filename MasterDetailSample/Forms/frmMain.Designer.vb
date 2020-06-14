<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.txtTitle = New System.Windows.Forms.Label()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.panelView = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.NwindDataSet = New MasterDetailSample.nwindDataSet()
        Me.CustomersTableAdapter = New MasterDetailSample.nwindDataSetTableAdapters.CustomersTableAdapter()
        Me.TableAdapterManager = New MasterDetailSample.nwindDataSetTableAdapters.TableAdapterManager()
        Me.InvoicesTableAdapter = New MasterDetailSample.nwindDataSetTableAdapters.InvoicesTableAdapter()
        Me.OrderReportsTableAdapter = New MasterDetailSample.nwindDataSetTableAdapters.OrderReportsTableAdapter()
        Me.btnReport = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        CType(Me.NwindDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.txtTitle)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Panel1.Size = New System.Drawing.Size(666, 83)
        Me.Panel1.TabIndex = 20
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(31, 15)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(47, 47)
        Me.PictureBox1.TabIndex = 11
        Me.PictureBox1.TabStop = False
        '
        'txtTitle
        '
        Me.txtTitle.AutoSize = True
        Me.txtTitle.Font = New System.Drawing.Font("Segoe UI Symbol", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTitle.ForeColor = System.Drawing.Color.DodgerBlue
        Me.txtTitle.Location = New System.Drawing.Point(84, 22)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(529, 37)
        Me.txtTitle.TabIndex = 10
        Me.txtTitle.Text = "Master Detail Demo - NorthWind Database"
        '
        'btnLoad
        '
        Me.btnLoad.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLoad.FlatAppearance.BorderColor = System.Drawing.Color.LightSkyBlue
        Me.btnLoad.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnLoad.Location = New System.Drawing.Point(394, 21)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 21
        Me.btnLoad.Text = "Load Data"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'panelView
        '
        Me.panelView.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panelView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.panelView.Location = New System.Drawing.Point(0, 83)
        Me.panelView.Name = "panelView"
        Me.panelView.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.panelView.Size = New System.Drawing.Size(666, 427)
        Me.panelView.TabIndex = 22
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.btnReport)
        Me.Panel3.Controls.Add(Me.btnLoad)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel3.Location = New System.Drawing.Point(0, 510)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Panel3.Size = New System.Drawing.Size(666, 67)
        Me.Panel3.TabIndex = 23
        '
        'NwindDataSet
        '
        Me.NwindDataSet.DataSetName = "nwindDataSet"
        Me.NwindDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'CustomersTableAdapter
        '
        Me.CustomersTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager
        '
        Me.TableAdapterManager.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager.CustomersTableAdapter = Me.CustomersTableAdapter
        Me.TableAdapterManager.UpdateOrder = MasterDetailSample.nwindDataSetTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'InvoicesTableAdapter
        '
        Me.InvoicesTableAdapter.ClearBeforeFill = True
        '
        'OrderReportsTableAdapter
        '
        Me.OrderReportsTableAdapter.ClearBeforeFill = True
        '
        'btnReport
        '
        Me.btnReport.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnReport.Enabled = False
        Me.btnReport.FlatAppearance.BorderColor = System.Drawing.Color.LightSkyBlue
        Me.btnReport.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnReport.Location = New System.Drawing.Point(498, 21)
        Me.btnReport.Name = "btnReport"
        Me.btnReport.Size = New System.Drawing.Size(75, 23)
        Me.btnReport.TabIndex = 22
        Me.btnReport.Text = "Report"
        Me.btnReport.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(227, Byte), Integer), CType(CType(241, Byte), Integer), CType(CType(254, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(666, 577)
        Me.Controls.Add(Me.panelView)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMain"
        Me.Text = "Master-Detail"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        CType(Me.NwindDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents txtTitle As System.Windows.Forms.Label
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents panelView As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents NwindDataSet As MasterDetailSample.nwindDataSet
    Friend WithEvents CustomersTableAdapter As MasterDetailSample.nwindDataSetTableAdapters.CustomersTableAdapter
    Friend WithEvents TableAdapterManager As MasterDetailSample.nwindDataSetTableAdapters.TableAdapterManager
    Friend WithEvents InvoicesTableAdapter As MasterDetailSample.nwindDataSetTableAdapters.InvoicesTableAdapter
    Friend WithEvents OrderReportsTableAdapter As MasterDetailSample.nwindDataSetTableAdapters.OrderReportsTableAdapter
    Friend WithEvents btnReport As Button
End Class
