<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPrintBom
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPrintBom))
        Me.txtDate = New System.Windows.Forms.Label()
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dtpTo = New System.Windows.Forms.DateTimePicker()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.DataGri = New System.Windows.Forms.DataGridView()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtMO = New System.Windows.Forms.TextBox()
        Me.btnExcel = New System.Windows.Forms.Button()
        Me.prgb1 = New System.Windows.Forms.ProgressBar()
        Me.Label3 = New System.Windows.Forms.Label()
        CType(Me.DataGri, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtDate
        '
        Me.txtDate.AutoSize = True
        Me.txtDate.Location = New System.Drawing.Point(21, 28)
        Me.txtDate.Name = "txtDate"
        Me.txtDate.Size = New System.Drawing.Size(38, 13)
        Me.txtDate.TabIndex = 0
        Me.txtDate.Text = "FROM"
        '
        'dtpFrom
        '
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFrom.Location = New System.Drawing.Point(65, 25)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(105, 20)
        Me.dtpFrom.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(175, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(22, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "TO"
        '
        'dtpTo
        '
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpTo.Location = New System.Drawing.Point(201, 25)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(105, 20)
        Me.dtpTo.TabIndex = 3
        '
        'btnSearch
        '
        Me.btnSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSearch.BackColor = System.Drawing.SystemColors.HighlightText
        Me.btnSearch.ForeColor = System.Drawing.SystemColors.Highlight
        Me.btnSearch.Location = New System.Drawing.Point(585, 17)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(101, 35)
        Me.btnSearch.TabIndex = 4
        Me.btnSearch.Text = "SEARCH"
        Me.btnSearch.UseVisualStyleBackColor = False
        '
        'DataGri
        '
        Me.DataGri.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGri.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGri.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight
        Me.DataGri.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.DataGri.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.RaisedHorizontal
        Me.DataGri.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGri.GridColor = System.Drawing.SystemColors.ActiveCaption
        Me.DataGri.Location = New System.Drawing.Point(24, 67)
        Me.DataGri.Name = "DataGri"
        Me.DataGri.Size = New System.Drawing.Size(701, 348)
        Me.DataGri.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(325, 28)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(24, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "MO"
        '
        'txtMO
        '
        Me.txtMO.Location = New System.Drawing.Point(367, 25)
        Me.txtMO.Name = "txtMO"
        Me.txtMO.Size = New System.Drawing.Size(145, 20)
        Me.txtMO.TabIndex = 7
        '
        'btnExcel
        '
        Me.btnExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExcel.BackColor = System.Drawing.Color.Snow
        Me.btnExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnExcel.ForeColor = System.Drawing.Color.Green
        Me.btnExcel.Location = New System.Drawing.Point(570, 421)
        Me.btnExcel.Name = "btnExcel"
        Me.btnExcel.Size = New System.Drawing.Size(120, 45)
        Me.btnExcel.TabIndex = 8
        Me.btnExcel.Text = "EXCEL"
        Me.btnExcel.UseVisualStyleBackColor = False
        '
        'prgb1
        '
        Me.prgb1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.prgb1.Location = New System.Drawing.Point(25, 436)
        Me.prgb1.MarqueeAnimationSpeed = 500
        Me.prgb1.Name = "prgb1"
        Me.prgb1.Size = New System.Drawing.Size(172, 23)
        Me.prgb1.Step = 5
        Me.prgb1.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.prgb1.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(228, 437)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(16, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "..."
        '
        'frmPrintBom
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(737, 471)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.prgb1)
        Me.Controls.Add(Me.btnExcel)
        Me.Controls.Add(Me.txtMO)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.DataGri)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.dtpTo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.txtDate)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmPrintBom"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Mfg Order Bom Bearing (Thomas Program)"
        CType(Me.DataGri, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtDate As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents DataGri As System.Windows.Forms.DataGridView
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtMO As System.Windows.Forms.TextBox
    Friend WithEvents btnExcel As System.Windows.Forms.Button
    Friend WithEvents prgb1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label3 As System.Windows.Forms.Label

End Class
