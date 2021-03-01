<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Login
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Login))
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.label4 = New System.Windows.Forms.Label()
        Me.label1 = New System.Windows.Forms.Label()
        Me.label2 = New System.Windows.Forms.Label()
        Me.linkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.button1 = New System.Windows.Forms.Button()
        Me.panel4 = New System.Windows.Forms.Panel()
        Me.textBox2 = New System.Windows.Forms.TextBox()
        Me.panel3 = New System.Windows.Forms.Panel()
        Me.textBox1 = New System.Windows.Forms.TextBox()
        Me.pictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.FromArgb(CType(CType(37, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(38, Byte), Integer))
        Me.Button3.FlatAppearance.BorderSize = 0
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Image = Global.Design_01.My.Resources.Resources.close_window_25px
        Me.Button3.Location = New System.Drawing.Point(296, 12)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(26, 23)
        Me.Button3.TabIndex = 22
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.FromArgb(CType(CType(37, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(38, Byte), Integer))
        Me.Button4.FlatAppearance.BorderSize = 0
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button4.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.Image = Global.Design_01.My.Resources.Resources.minimize_window_25px
        Me.Button4.Location = New System.Drawing.Point(264, 12)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(26, 23)
        Me.Button4.TabIndex = 23
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(45, Byte), Integer), CType(CType(48, Byte), Integer))
        Me.Panel1.Controls.Add(Me.label4)
        Me.Panel1.Controls.Add(Me.label1)
        Me.Panel1.Controls.Add(Me.label2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(581, 583)
        Me.Panel1.TabIndex = 24
        '
        'label4
        '
        Me.label4.AutoSize = True
        Me.label4.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label4.ForeColor = System.Drawing.Color.White
        Me.label4.Location = New System.Drawing.Point(37, 215)
        Me.label4.Name = "label4"
        Me.label4.Size = New System.Drawing.Size(303, 21)
        Me.label4.TabIndex = 11
        Me.label4.Text = "Welcome to Grading Management System"
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label1.ForeColor = System.Drawing.Color.White
        Me.label1.Location = New System.Drawing.Point(38, 125)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(184, 15)
        Me.label1.TabIndex = 10
        Me.label1.Text = "Powered By Mobarak Dimalotang"
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Font = New System.Drawing.Font("Bahnschrift Condensed", 27.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label2.ForeColor = System.Drawing.Color.White
        Me.label2.Location = New System.Drawing.Point(33, 80)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(372, 45)
        Me.label2.TabIndex = 9
        Me.label2.Text = "Grading Management System"
        '
        'linkLabel1
        '
        Me.linkLabel1.ActiveLinkColor = System.Drawing.Color.Gray
        Me.linkLabel1.AutoSize = True
        Me.linkLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.linkLabel1.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline
        Me.linkLabel1.LinkColor = System.Drawing.Color.White
        Me.linkLabel1.Location = New System.Drawing.Point(227, 407)
        Me.linkLabel1.Name = "linkLabel1"
        Me.linkLabel1.Size = New System.Drawing.Size(81, 17)
        Me.linkLabel1.TabIndex = 14
        Me.linkLabel1.TabStop = True
        Me.linkLabel1.Text = "Need Help ?"
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(37, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(38, Byte), Integer))
        Me.Panel2.Controls.Add(Me.linkLabel1)
        Me.Panel2.Controls.Add(Me.Button4)
        Me.Panel2.Controls.Add(Me.button1)
        Me.Panel2.Controls.Add(Me.Button3)
        Me.Panel2.Controls.Add(Me.panel4)
        Me.Panel2.Controls.Add(Me.textBox2)
        Me.Panel2.Controls.Add(Me.panel3)
        Me.Panel2.Controls.Add(Me.textBox1)
        Me.Panel2.Controls.Add(Me.pictureBox1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(581, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(334, 583)
        Me.Panel2.TabIndex = 25
        '
        'button1
        '
        Me.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.button1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.button1.ForeColor = System.Drawing.Color.White
        Me.button1.Image = Global.Design_01.My.Resources.Resources.login_rounded_up_40px
        Me.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.button1.Location = New System.Drawing.Point(26, 444)
        Me.button1.Name = "button1"
        Me.button1.Size = New System.Drawing.Size(282, 41)
        Me.button1.TabIndex = 13
        Me.button1.Text = "Login"
        Me.button1.UseVisualStyleBackColor = True
        '
        'panel4
        '
        Me.panel4.BackColor = System.Drawing.Color.White
        Me.panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panel4.Location = New System.Drawing.Point(26, 390)
        Me.panel4.Name = "panel4"
        Me.panel4.Size = New System.Drawing.Size(282, 3)
        Me.panel4.TabIndex = 11
        '
        'textBox2
        '
        Me.textBox2.BackColor = System.Drawing.Color.FromArgb(CType(CType(37, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(38, Byte), Integer))
        Me.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.textBox2.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textBox2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.textBox2.Location = New System.Drawing.Point(26, 362)
        Me.textBox2.Name = "textBox2"
        Me.textBox2.Size = New System.Drawing.Size(282, 22)
        Me.textBox2.TabIndex = 12
        Me.textBox2.Text = "Password"
        Me.textBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.textBox2.UseSystemPasswordChar = True
        '
        'panel3
        '
        Me.panel3.BackColor = System.Drawing.Color.White
        Me.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panel3.Location = New System.Drawing.Point(26, 338)
        Me.panel3.Name = "panel3"
        Me.panel3.Size = New System.Drawing.Size(282, 3)
        Me.panel3.TabIndex = 9
        '
        'textBox1
        '
        Me.textBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(37, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(38, Byte), Integer))
        Me.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.textBox1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textBox1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.textBox1.Location = New System.Drawing.Point(26, 310)
        Me.textBox1.Name = "textBox1"
        Me.textBox1.Size = New System.Drawing.Size(282, 22)
        Me.textBox1.TabIndex = 10
        Me.textBox1.Text = "Username"
        Me.textBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'pictureBox1
        '
        Me.pictureBox1.Image = Global.Design_01.My.Resources.Resources.school_director_120px
        Me.pictureBox1.Location = New System.Drawing.Point(108, 116)
        Me.pictureBox1.Name = "pictureBox1"
        Me.pictureBox1.Size = New System.Drawing.Size(120, 120)
        Me.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pictureBox1.TabIndex = 1
        Me.pictureBox1.TabStop = False
        '
        'Login
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.AliceBlue
        Me.ClientSize = New System.Drawing.Size(915, 583)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "Login"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Grading Management System"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Private WithEvents pictureBox1 As System.Windows.Forms.PictureBox
    Private WithEvents linkLabel1 As System.Windows.Forms.LinkLabel
    Private WithEvents button1 As System.Windows.Forms.Button
    Private WithEvents panel4 As System.Windows.Forms.Panel
    Private WithEvents textBox2 As System.Windows.Forms.TextBox
    Private WithEvents panel3 As System.Windows.Forms.Panel
    Private WithEvents textBox1 As System.Windows.Forms.TextBox
    Private WithEvents label2 As System.Windows.Forms.Label
    Private WithEvents label1 As System.Windows.Forms.Label
    Private WithEvents label4 As System.Windows.Forms.Label
End Class
