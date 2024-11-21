Imports System.Data.SqlClient

Public Class Ambientes

    Private Sub Ambientes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LlenarCombo()
    End Sub

    Sub LlenarCombo()
        Dim sql As String = "SELECT ambiente FROM ambientes ORDER BY ambiente"

        Dim cmd As New SqlCommand(sql, cnn)
        Dim adapter As New SqlDataAdapter(cmd)
        Dim dataTable As New DataTable()

        Try
            conectado()
            adapter.Fill(dataTable)
        Catch ex As Exception
            MessageBox.Show("Error Con la base de datos:" & ex.Message)
        Finally
            cerrar_conexion()
        End Try

        cmbAmbientes.DataSource = dataTable
        cmbAmbientes.DisplayMember = "ambiente"
        cmbAmbientes.ValueMember = "ambiente"

    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click

        If cmbAmbientes.SelectedItem Is Nothing Then
            MessageBox.Show("Por favor, seleccione un ambiente.")
        End If

        Dim sql As String = "  SELECT  instructor, ambiente_lunes AS Ambiente,  CONCAT(hlunes_iniciada, ' - ', hlunes_terminada) AS Lunes, CONCAT(hmartes_iniciada, ' - ', hmartes_terminada) AS Martes, CONCAT(hmiercoles_iniciada, ' - ', hmiercoles_terminada) AS Miercoles,CONCAT(hjueves_iniciada, ' - ', hjueves_terminada) AS Jueves, CONCAT(hviernes_iniciada, ' - ', hviernes_terminada) AS Viernes, CONCAT(hsabado_iniciada, ' - ', hsabado_terminada) AS Sabado, CONCAT(hdomingo_iniciada, ' - ', hdomingo_terminada) AS Domingo, competencia AS Competencia, fecha_de_inicio As Inicio, fecha_de_terminacion AS Terminacion from programacion where instructor is not null and ambiente_lunes = @Ambiente  and fecha_de_inicio <= @Fecha and fecha_de_terminacion >= @Fecha ORDER BY hlunes_iniciada ASC"

        Dim cmd As New SqlCommand(sql, cnn)
        cmd.Parameters.AddWithValue("@Ambiente", cmbAmbientes.SelectedValue)
        cmd.Parameters.AddWithValue("@Fecha", dtpFecha.Value)

        Dim adapter As New SqlDataAdapter(cmd)
        Dim dataTable As New DataTable()

        Try
            conectado()
            adapter.Fill(dataTable)
        Catch ex As Exception
            MessageBox.Show("Error Con la base de datos:" & ex.Message)

        End Try
        cerrar_conexion()

        dgvAmbientes.DataSource = dataTable
        dgvAmbientes.Columns("instructor").Width = 200
        dgvAmbientes.Columns("Lunes").Width = 80
        dgvAmbientes.Columns("Martes").Width = 80
        dgvAmbientes.Columns("Miercoles").Width = 80
        dgvAmbientes.Columns("Jueves").Width = 80
        dgvAmbientes.Columns("Viernes").Width = 80
        dgvAmbientes.Columns("Sabado").Width = 80
        dgvAmbientes.Columns("Domingo").Width = 80


    End Sub


End Class

