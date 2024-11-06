<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Modalidad
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
        Me.datagridmodalidad = New System.Windows.Forms.DataGridView()
        CType(Me.datagridmodalidad, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'datagridmodalidad
        '
        Me.datagridmodalidad.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.datagridmodalidad.Location = New System.Drawing.Point(12, 111)
        Me.datagridmodalidad.Name = "datagridmodalidad"
        Me.datagridmodalidad.Size = New System.Drawing.Size(444, 150)
        Me.datagridmodalidad.TabIndex = 0
        '
        'Modalidad
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(468, 273)
        Me.Controls.Add(Me.datagridmodalidad)
        Me.Name = "Modalidad"
        Me.Text = "Modalidad"
        CType(Me.datagridmodalidad, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents datagridmodalidad As System.Windows.Forms.DataGridView
End Class
