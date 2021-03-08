<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CheckForm
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
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

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Yepbtn = New System.Windows.Forms.Button()
        Me.Fonbtn = New System.Windows.Forms.Button()
        Me.Delbtn = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.ConfirmBtn = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ClearBtn = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Yepbtn
        '
        Me.Yepbtn.BackColor = System.Drawing.Color.Yellow
        Me.Yepbtn.Location = New System.Drawing.Point(136, 123)
        Me.Yepbtn.Name = "Yepbtn"
        Me.Yepbtn.Size = New System.Drawing.Size(75, 23)
        Me.Yepbtn.TabIndex = 0
        Me.Yepbtn.Text = "Классно!"
        Me.Yepbtn.UseVisualStyleBackColor = False
        '
        'Fonbtn
        '
        Me.Fonbtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Fonbtn.Location = New System.Drawing.Point(217, 123)
        Me.Fonbtn.Name = "Fonbtn"
        Me.Fonbtn.Size = New System.Drawing.Size(75, 23)
        Me.Fonbtn.TabIndex = 1
        Me.Fonbtn.Text = "Фон"
        Me.Fonbtn.UseVisualStyleBackColor = False
        '
        'Delbtn
        '
        Me.Delbtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Delbtn.Location = New System.Drawing.Point(180, 94)
        Me.Delbtn.Name = "Delbtn"
        Me.Delbtn.Size = New System.Drawing.Size(75, 23)
        Me.Delbtn.TabIndex = 2
        Me.Delbtn.Text = "Удалить"
        Me.Delbtn.UseVisualStyleBackColor = False
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(12, 5)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(413, 20)
        Me.TextBox1.TabIndex = 3
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(12, 38)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(413, 20)
        Me.TextBox2.TabIndex = 3
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(127, 165)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(191, 21)
        Me.ComboBox1.TabIndex = 4
        '
        'ConfirmBtn
        '
        Me.ConfirmBtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ConfirmBtn.Location = New System.Drawing.Point(324, 165)
        Me.ConfirmBtn.Name = "ConfirmBtn"
        Me.ConfirmBtn.Size = New System.Drawing.Size(92, 33)
        Me.ConfirmBtn.TabIndex = 5
        Me.ConfirmBtn.Text = "Отправить"
        Me.ConfirmBtn.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(253, 149)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(13, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "0"
        '
        'ClearBtn
        '
        Me.ClearBtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ClearBtn.Location = New System.Drawing.Point(127, 192)
        Me.ClearBtn.Name = "ClearBtn"
        Me.ClearBtn.Size = New System.Drawing.Size(85, 23)
        Me.ClearBtn.TabIndex = 7
        Me.ClearBtn.Text = "Clear"
        Me.ClearBtn.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(142, 149)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(70, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Всего папок"
        '
        'CheckForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(433, 389)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ClearBtn)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ConfirmBtn)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Delbtn)
        Me.Controls.Add(Me.Fonbtn)
        Me.Controls.Add(Me.Yepbtn)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CheckForm"
        Me.Text = "CheckForm"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Yepbtn As System.Windows.Forms.Button
    Friend WithEvents Fonbtn As System.Windows.Forms.Button
    Friend WithEvents Delbtn As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents ConfirmBtn As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ClearBtn As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
