<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.lblnomb_archivo = New System.Windows.Forms.Label()
        Me.btnadjuntar = New System.Windows.Forms.Button()
        Me.txtcontraseña = New System.Windows.Forms.TextBox()
        Me.lblcontrasena = New System.Windows.Forms.Label()
        Me.txtcuerpo_msg = New System.Windows.Forms.TextBox()
        Me.txtasunto = New System.Windows.Forms.TextBox()
        Me.txtpara = New System.Windows.Forms.TextBox()
        Me.txtdesde = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnenviar = New System.Windows.Forms.Button()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cmbdescripcion = New System.Windows.Forms.ComboBox()
        Me.lbldefault = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtservidor = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtpuerto = New System.Windows.Forms.TextBox()
        Me.abrir = New System.Windows.Forms.OpenFileDialog()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(63, 31)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(617, 317)
        Me.TabControl1.TabIndex = 10
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.lblnomb_archivo)
        Me.TabPage1.Controls.Add(Me.btnadjuntar)
        Me.TabPage1.Controls.Add(Me.txtcontraseña)
        Me.TabPage1.Controls.Add(Me.lblcontrasena)
        Me.TabPage1.Controls.Add(Me.txtcuerpo_msg)
        Me.TabPage1.Controls.Add(Me.txtasunto)
        Me.TabPage1.Controls.Add(Me.txtpara)
        Me.TabPage1.Controls.Add(Me.txtdesde)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.btnenviar)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(609, 291)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Enviar"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'lblnomb_archivo
        '
        Me.lblnomb_archivo.AutoSize = True
        Me.lblnomb_archivo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblnomb_archivo.Location = New System.Drawing.Point(131, 131)
        Me.lblnomb_archivo.Name = "lblnomb_archivo"
        Me.lblnomb_archivo.Size = New System.Drawing.Size(0, 13)
        Me.lblnomb_archivo.TabIndex = 21
        Me.lblnomb_archivo.Visible = False
        '
        'btnadjuntar
        '
        Me.btnadjuntar.Location = New System.Drawing.Point(96, 123)
        Me.btnadjuntar.Name = "btnadjuntar"
        Me.btnadjuntar.Size = New System.Drawing.Size(29, 28)
        Me.btnadjuntar.TabIndex = 20
        Me.btnadjuntar.UseVisualStyleBackColor = True
        '
        'txtcontraseña
        '
        Me.txtcontraseña.Location = New System.Drawing.Point(384, 25)
        Me.txtcontraseña.Name = "txtcontraseña"
        Me.txtcontraseña.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtcontraseña.Size = New System.Drawing.Size(115, 20)
        Me.txtcontraseña.TabIndex = 19
        '
        'lblcontrasena
        '
        Me.lblcontrasena.AutoSize = True
        Me.lblcontrasena.Location = New System.Drawing.Point(311, 28)
        Me.lblcontrasena.Name = "lblcontrasena"
        Me.lblcontrasena.Size = New System.Drawing.Size(67, 13)
        Me.lblcontrasena.TabIndex = 18
        Me.lblcontrasena.Text = "Contraseña :"
        '
        'txtcuerpo_msg
        '
        Me.txtcuerpo_msg.Location = New System.Drawing.Point(53, 170)
        Me.txtcuerpo_msg.Multiline = True
        Me.txtcuerpo_msg.Name = "txtcuerpo_msg"
        Me.txtcuerpo_msg.Size = New System.Drawing.Size(446, 102)
        Me.txtcuerpo_msg.TabIndex = 17
        '
        'txtasunto
        '
        Me.txtasunto.Location = New System.Drawing.Point(134, 89)
        Me.txtasunto.Name = "txtasunto"
        Me.txtasunto.Size = New System.Drawing.Size(365, 20)
        Me.txtasunto.TabIndex = 16
        '
        'txtpara
        '
        Me.txtpara.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtpara.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.txtpara.Location = New System.Drawing.Point(134, 58)
        Me.txtpara.Name = "txtpara"
        Me.txtpara.Size = New System.Drawing.Size(365, 20)
        Me.txtpara.TabIndex = 15
        '
        'txtdesde
        '
        Me.txtdesde.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdesde.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.txtdesde.Location = New System.Drawing.Point(134, 25)
        Me.txtdesde.Name = "txtdesde"
        Me.txtdesde.Size = New System.Drawing.Size(171, 20)
        Me.txtdesde.TabIndex = 14
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(22, 154)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(106, 13)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Cuerpo del mensaje :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(82, 92)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 13)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Asunto :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(93, 61)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(35, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Para :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(84, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(44, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Desde :"
        '
        'btnenviar
        '
        Me.btnenviar.Location = New System.Drawing.Point(515, 242)
        Me.btnenviar.Name = "btnenviar"
        Me.btnenviar.Size = New System.Drawing.Size(78, 30)
        Me.btnenviar.TabIndex = 9
        Me.btnenviar.Text = "Enviar"
        Me.btnenviar.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.GroupBox1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(609, 291)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Configurar Servidor"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cmbdescripcion)
        Me.GroupBox1.Controls.Add(Me.lbldefault)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.txtservidor)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtpuerto)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 15)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(585, 263)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Configuracion"
        '
        'cmbdescripcion
        '
        Me.cmbdescripcion.FormattingEnabled = True
        Me.cmbdescripcion.Items.AddRange(New Object() {"Outlook", "Gmail"})
        Me.cmbdescripcion.Location = New System.Drawing.Point(180, 86)
        Me.cmbdescripcion.Name = "cmbdescripcion"
        Me.cmbdescripcion.Size = New System.Drawing.Size(128, 21)
        Me.cmbdescripcion.TabIndex = 5
        '
        'lbldefault
        '
        Me.lbldefault.AutoSize = True
        Me.lbldefault.Location = New System.Drawing.Point(300, 155)
        Me.lbldefault.Name = "lbldefault"
        Me.lbldefault.Size = New System.Drawing.Size(0, 13)
        Me.lbldefault.TabIndex = 7
        Me.lbldefault.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(122, 121)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(52, 13)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Servidor :"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(247, 155)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(47, 13)
        Me.Label8.TabIndex = 6
        Me.Label8.Text = "Default :"
        '
        'txtservidor
        '
        Me.txtservidor.Location = New System.Drawing.Point(180, 118)
        Me.txtservidor.Name = "txtservidor"
        Me.txtservidor.Size = New System.Drawing.Size(183, 20)
        Me.txtservidor.TabIndex = 1
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(130, 155)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(44, 13)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "Puerto :"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(105, 89)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(69, 13)
        Me.Label7.TabIndex = 4
        Me.Label7.Text = "Descripcion :"
        '
        'txtpuerto
        '
        Me.txtpuerto.Location = New System.Drawing.Point(180, 152)
        Me.txtpuerto.Name = "txtpuerto"
        Me.txtpuerto.Size = New System.Drawing.Size(51, 20)
        Me.txtpuerto.TabIndex = 3
        '
        'abrir
        '
        Me.abrir.Multiselect = True
        Me.abrir.Title = "Selecione el Achivo Adjunto"
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(757, 394)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "Form2"
        Me.Text = "Form2"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents lblnomb_archivo As System.Windows.Forms.Label
    Friend WithEvents btnadjuntar As System.Windows.Forms.Button
    Friend WithEvents txtcontraseña As System.Windows.Forms.TextBox
    Friend WithEvents lblcontrasena As System.Windows.Forms.Label
    Friend WithEvents txtcuerpo_msg As System.Windows.Forms.TextBox
    Friend WithEvents txtasunto As System.Windows.Forms.TextBox
    Friend WithEvents txtpara As System.Windows.Forms.TextBox
    Friend WithEvents txtdesde As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnenviar As System.Windows.Forms.Button
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmbdescripcion As System.Windows.Forms.ComboBox
    Friend WithEvents lbldefault As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtservidor As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtpuerto As System.Windows.Forms.TextBox
    Friend WithEvents abrir As System.Windows.Forms.OpenFileDialog
End Class
