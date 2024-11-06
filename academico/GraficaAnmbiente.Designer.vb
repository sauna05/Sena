<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class GraficaAnmbiente
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.GroupBox10 = New System.Windows.Forms.GroupBox()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.Dtpfecha = New System.Windows.Forms.DateTimePicker()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label58 = New System.Windows.Forms.Label()
        Me.Label57 = New System.Windows.Forms.Label()
        Me.ComboBox8 = New System.Windows.Forms.ComboBox()
        Me.Button26 = New System.Windows.Forms.Button()
        Me.TextBox12 = New System.Windows.Forms.TextBox()
        Me.Label71 = New System.Windows.Forms.Label()
        Me.Dtambiente = New System.Windows.Forms.DataGridView()
        Me.dt_grafica_semana = New System.Windows.Forms.DataGridView()
        Me.GroupBox10.SuspendLayout()
        CType(Me.Dtambiente, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dt_grafica_semana, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox10
        '
        Me.GroupBox10.Controls.Add(Me.Label29)
        Me.GroupBox10.Controls.Add(Me.Dtpfecha)
        Me.GroupBox10.Controls.Add(Me.TextBox2)
        Me.GroupBox10.Controls.Add(Me.Label58)
        Me.GroupBox10.Controls.Add(Me.Label57)
        Me.GroupBox10.Controls.Add(Me.ComboBox8)
        Me.GroupBox10.Controls.Add(Me.Button26)
        Me.GroupBox10.Controls.Add(Me.TextBox12)
        Me.GroupBox10.Controls.Add(Me.Label71)
        Me.GroupBox10.Location = New System.Drawing.Point(23, 4)
        Me.GroupBox10.Name = "GroupBox10"
        Me.GroupBox10.Size = New System.Drawing.Size(1329, 116)
        Me.GroupBox10.TabIndex = 74
        Me.GroupBox10.TabStop = False
        Me.GroupBox10.Text = "Ambiente"
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(36, 90)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(203, 13)
        Me.Label29.TabIndex = 74
        Me.Label29.Text = "FECHA PARA SELECCION DE SEMANA"
        '
        'Dtpfecha
        '
        Me.Dtpfecha.Location = New System.Drawing.Point(283, 90)
        Me.Dtpfecha.Name = "Dtpfecha"
        Me.Dtpfecha.Size = New System.Drawing.Size(200, 20)
        Me.Dtpfecha.TabIndex = 73
        '
        'TextBox2
        '
        Me.TextBox2.Enabled = False
        Me.TextBox2.Location = New System.Drawing.Point(94, 10)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(43, 20)
        Me.TextBox2.TabIndex = 72
        '
        'Label58
        '
        Me.Label58.AutoSize = True
        Me.Label58.Location = New System.Drawing.Point(70, 12)
        Me.Label58.Name = "Label58"
        Me.Label58.Size = New System.Drawing.Size(18, 13)
        Me.Label58.TabIndex = 71
        Me.Label58.Text = "ID"
        '
        'Label57
        '
        Me.Label57.AutoSize = True
        Me.Label57.Location = New System.Drawing.Point(36, 65)
        Me.Label57.Name = "Label57"
        Me.Label57.Size = New System.Drawing.Size(52, 13)
        Me.Label57.TabIndex = 70
        Me.Label57.Text = "Municipio"
        '
        'ComboBox8
        '
        Me.ComboBox8.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.ComboBox8.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox8.FormattingEnabled = True
        Me.ComboBox8.Location = New System.Drawing.Point(94, 62)
        Me.ComboBox8.Name = "ComboBox8"
        Me.ComboBox8.Size = New System.Drawing.Size(325, 21)
        Me.ComboBox8.TabIndex = 69
        '
        'Button26
        '
        Me.Button26.Location = New System.Drawing.Point(681, 28)
        Me.Button26.Name = "Button26"
        Me.Button26.Size = New System.Drawing.Size(128, 24)
        Me.Button26.TabIndex = 28
        Me.Button26.Text = "Ver"
        Me.Button26.UseVisualStyleBackColor = True
        '
        'TextBox12
        '
        Me.TextBox12.Location = New System.Drawing.Point(94, 36)
        Me.TextBox12.Name = "TextBox12"
        Me.TextBox12.Size = New System.Drawing.Size(507, 20)
        Me.TextBox12.TabIndex = 6
        '
        'Label71
        '
        Me.Label71.AutoSize = True
        Me.Label71.Location = New System.Drawing.Point(44, 39)
        Me.Label71.Name = "Label71"
        Me.Label71.Size = New System.Drawing.Size(44, 13)
        Me.Label71.TabIndex = 5
        Me.Label71.Text = "Nombre"
        '
        'Dtambiente
        '
        Me.Dtambiente.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Dtambiente.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.Dtambiente.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Dtambiente.DefaultCellStyle = DataGridViewCellStyle2
        Me.Dtambiente.Location = New System.Drawing.Point(23, 126)
        Me.Dtambiente.Name = "Dtambiente"
        Me.Dtambiente.ReadOnly = True
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Dtambiente.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.Dtambiente.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.Dtambiente.Size = New System.Drawing.Size(306, 541)
        Me.Dtambiente.TabIndex = 76
        '
        'dt_grafica_semana
        '
        Me.dt_grafica_semana.AllowUserToAddRows = False
        Me.dt_grafica_semana.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dt_grafica_semana.Location = New System.Drawing.Point(335, 126)
        Me.dt_grafica_semana.Name = "dt_grafica_semana"
        Me.dt_grafica_semana.Size = New System.Drawing.Size(1017, 541)
        Me.dt_grafica_semana.TabIndex = 77
        '
        'GraficaAnmbiente
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1364, 659)
        Me.Controls.Add(Me.dt_grafica_semana)
        Me.Controls.Add(Me.Dtambiente)
        Me.Controls.Add(Me.GroupBox10)
        Me.Name = "GraficaAnmbiente"
        Me.Text = "GraficaAnmbiente"
        Me.GroupBox10.ResumeLayout(False)
        Me.GroupBox10.PerformLayout()
        CType(Me.Dtambiente, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dt_grafica_semana, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox10 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label58 As System.Windows.Forms.Label
    Friend WithEvents Label57 As System.Windows.Forms.Label
    Friend WithEvents ComboBox8 As System.Windows.Forms.ComboBox
    Friend WithEvents Button26 As System.Windows.Forms.Button
    Friend WithEvents TextBox12 As System.Windows.Forms.TextBox
    Friend WithEvents Label71 As System.Windows.Forms.Label
    Friend WithEvents Dtambiente As System.Windows.Forms.DataGridView
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Dtpfecha As System.Windows.Forms.DateTimePicker
    Friend WithEvents dt_grafica_semana As System.Windows.Forms.DataGridView
End Class
