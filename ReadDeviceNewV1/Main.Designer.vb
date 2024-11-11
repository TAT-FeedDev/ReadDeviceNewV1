<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main
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
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.btnLoadExcel = New System.Windows.Forms.Button()
        Me.btnUpdateDatabase = New System.Windows.Forms.Button()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.tsProgressBar = New System.Windows.Forms.ToolStripProgressBar()
        Me.tsStatus = New System.Windows.Forms.ToolStripStatusLabel()
        Me.cboStation = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cmdDatabaseRoute = New System.Windows.Forms.ComboBox()
        Me.cmdDatabaseBatching = New System.Windows.Forms.ComboBox()
        Me.btnSaveConfigDb = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.StatusStrip1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Location = New System.Drawing.Point(0, 74)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1308, 663)
        Me.TabControl1.TabIndex = 0
        '
        'btnLoadExcel
        '
        Me.btnLoadExcel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnLoadExcel.Location = New System.Drawing.Point(7, 12)
        Me.btnLoadExcel.Name = "btnLoadExcel"
        Me.btnLoadExcel.Size = New System.Drawing.Size(103, 23)
        Me.btnLoadExcel.TabIndex = 1
        Me.btnLoadExcel.Text = "Load Excel"
        Me.btnLoadExcel.UseVisualStyleBackColor = True
        '
        'btnUpdateDatabase
        '
        Me.btnUpdateDatabase.Location = New System.Drawing.Point(7, 41)
        Me.btnUpdateDatabase.Name = "btnUpdateDatabase"
        Me.btnUpdateDatabase.Size = New System.Drawing.Size(103, 23)
        Me.btnUpdateDatabase.TabIndex = 3
        Me.btnUpdateDatabase.Text = "Update"
        Me.btnUpdateDatabase.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsProgressBar, Me.tsStatus})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 740)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1308, 22)
        Me.StatusStrip1.TabIndex = 3
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tsProgressBar
        '
        Me.tsProgressBar.Name = "tsProgressBar"
        Me.tsProgressBar.Size = New System.Drawing.Size(100, 16)
        Me.tsProgressBar.Step = 1
        '
        'tsStatus
        '
        Me.tsStatus.Name = "tsStatus"
        Me.tsStatus.Size = New System.Drawing.Size(112, 17)
        Me.tsStatus.Text = "toolStripStatusLabel"
        '
        'cboStation
        '
        Me.cboStation.FormattingEnabled = True
        Me.cboStation.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20"})
        Me.cboStation.Location = New System.Drawing.Point(170, 12)
        Me.cboStation.Name = "cboStation"
        Me.cboStation.Size = New System.Drawing.Size(68, 21)
        Me.cboStation.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(119, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "PLC No."
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.GroupBox1.Controls.Add(Me.cmdDatabaseRoute)
        Me.GroupBox1.Controls.Add(Me.cmdDatabaseBatching)
        Me.GroupBox1.Controls.Add(Me.btnSaveConfigDb)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtPassword)
        Me.GroupBox1.Controls.Add(Me.txtUsername)
        Me.GroupBox1.Controls.Add(Me.txtServer)
        Me.GroupBox1.Location = New System.Drawing.Point(263, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(632, 69)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Config DB"
        '
        'cmdDatabaseRoute
        '
        Me.cmdDatabaseRoute.FormattingEnabled = True
        Me.cmdDatabaseRoute.Location = New System.Drawing.Point(375, 41)
        Me.cmdDatabaseRoute.Name = "cmdDatabaseRoute"
        Me.cmdDatabaseRoute.Size = New System.Drawing.Size(155, 21)
        Me.cmdDatabaseRoute.TabIndex = 3
        '
        'cmdDatabaseBatching
        '
        Me.cmdDatabaseBatching.FormattingEnabled = True
        Me.cmdDatabaseBatching.Location = New System.Drawing.Point(110, 42)
        Me.cmdDatabaseBatching.Name = "cmdDatabaseBatching"
        Me.cmdDatabaseBatching.Size = New System.Drawing.Size(155, 21)
        Me.cmdDatabaseBatching.TabIndex = 3
        '
        'btnSaveConfigDb
        '
        Me.btnSaveConfigDb.Location = New System.Drawing.Point(545, 40)
        Me.btnSaveConfigDb.Name = "btnSaveConfigDb"
        Me.btnSaveConfigDb.Size = New System.Drawing.Size(75, 23)
        Me.btnSaveConfigDb.TabIndex = 2
        Me.btnSaveConfigDb.Text = "Save"
        Me.btnSaveConfigDb.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(334, 17)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(63, 13)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Password : "
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(274, 44)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(95, 13)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "Database Route : "
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(157, 17)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(65, 13)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Username : "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 44)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(107, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Database Batching : "
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 17)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Server : "
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(403, 14)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(100, 21)
        Me.txtPassword.TabIndex = 0
        Me.txtPassword.UseSystemPasswordChar = True
        '
        'txtUsername
        '
        Me.txtUsername.Location = New System.Drawing.Point(228, 14)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(100, 21)
        Me.txtUsername.TabIndex = 0
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(59, 14)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(90, 21)
        Me.txtServer.TabIndex = 0
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1308, 762)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboStation)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnUpdateDatabase)
        Me.Controls.Add(Me.btnLoadExcel)
        Me.Controls.Add(Me.TabControl1)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Name = "Main"
        Me.Text = "Main"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents btnLoadExcel As Button
    Friend WithEvents btnUpdateDatabase As Button
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents tsStatus As ToolStripStatusLabel
    Friend WithEvents tsProgressBar As ToolStripProgressBar
    Friend WithEvents cboStation As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents btnSaveConfigDb As Button
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents txtPassword As TextBox
    Friend WithEvents txtUsername As TextBox
    Friend WithEvents txtServer As TextBox
    Friend WithEvents cmdDatabaseBatching As ComboBox
    Friend WithEvents cmdDatabaseRoute As ComboBox
    Friend WithEvents Label6 As Label
End Class
