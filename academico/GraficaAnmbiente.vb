Imports System.Data.SqlClient

Public Class GraficaAnmbiente

    Public Function LlenarDatos(sql As String) As DataTable
        Dim datos As New DataTable()
        Dim adapter As New SqlDataAdapter(sql, cnn)
        adapter.Fill(datos)
        Return datos
    End Function
    Private Sub GraficaAnmbiente_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sql = "Select top 100 * from  ambientes"
        LlenarDataGrids(Dtambiente)
    End Sub

    Private Sub dtambiente_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Dtambiente.CellClick
        Dim FILA As Integer = Dtambiente.CurrentRow.Index.ToString
        TextBox2.Text = Dtambiente.Rows(FILA).Cells("ID").Value.ToString()
        TextBox12.Text = Dtambiente.Rows(FILA).Cells("ambiente").Value.ToString()
        ComboBox8.Text = Dtambiente.Rows(FILA).Cells("Municipio").Value.ToString()
        grafica()

    End Sub

    Sub grafica()
        'se debe programar la graficacion del amiente seleccionado los 7 dias de la semaa de acuerdo a la fecha
        Dim fecha As Date = Dtpfecha.Value
        Dim dia_semana = Weekday(fecha, FirstDayOfWeek.Monday)
        Dim fecha_lunes = fecha.AddDays(-dia_semana)
        fecha_lunes = fecha_lunes.AddDays(1)
        dt_grafica_semana.ColumnCount = 7

        Dim i As Integer
        Dim fecha_encabezado As Date = fecha_lunes
        For i = 1 To 7

            dt_grafica_semana.Columns(i - 1).HeaderText = diasemana(Weekday(fecha_encabezado, FirstDayOfWeek.Monday)) & vbCrLf & fecha_encabezado.Date
            dt_grafica_semana.Columns(i - 1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.BottomCenter
            fecha_encabezado = fecha_encabezado.AddDays(1)

        Next
        Dim hora_inicio_grafica As DateTime
        Dim hora_fin_grafica As DateTime
        hora_inicio_grafica = "06:00:00"
        hora_fin_grafica = "23:00:00"
        i = 0
        While hora_inicio_grafica <= hora_fin_grafica
            hora_inicio_grafica = hora_inicio_grafica.AddMinutes(30)
            i += 1
        End While
        dt_grafica_semana.RowCount = i
        Dim filas, columnas As Integer
        For columnas = 0 To 6
            For filas = 0 To dt_grafica_semana.RowCount - 1
                dt_grafica_semana.Rows(filas).Cells(columnas).Value = ""
                dt_grafica_semana.Rows(filas).Cells(columnas).Style.BackColor = Color.White
            Next
        Next
        hora_inicio_grafica = "06:00:00"
        hora_fin_grafica = "23:00:00"
        i = 0
        While hora_inicio_grafica <= hora_fin_grafica
            dt_grafica_semana.Rows(i).HeaderCell.Value = hora_inicio_grafica.TimeOfDay.ToString
            dt_grafica_semana.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders)
            hora_inicio_grafica = hora_inicio_grafica.AddMinutes(30)
            i += 1

        End While
        Dim text_dia As String
        fecha_encabezado = fecha_lunes
        For i = 0 To 6
            text_dia = diasemana_minuscula(Weekday(fecha_encabezado, FirstDayOfWeek.Monday))
            sql = "select ficha, h" & text_dia & "_iniciada, h" & text_dia & "_terminada, fecha_de_inicio, fecha_de_terminacion from programacion where  ambiente_" & text_dia & " = '" & TextBox12.Text & "' and '" & fecha_encabezado.Date & "' BETWEEN fecha_de_inicio AND fecha_de_terminacion"
            Clipboard.SetText(sql)

            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader
            While reader.Read
                Dim j As Integer = 0
                hora_inicio_grafica = "06:00:00"
                hora_fin_grafica = "23:00:00"
                While hora_inicio_grafica <= hora_fin_grafica
                    Dim h_ref_inicio As DateTime = reader(1) & ":00"
                    Dim h_ref_fin As DateTime = reader(2) & ":00"
                    If h_ref_inicio <= hora_inicio_grafica And hora_inicio_grafica < h_ref_fin Then


                        dt_grafica_semana.Rows(j).Cells(i).Value = reader("ficha")
                        Dim f_termina_ref As Date = reader("fecha_de_terminacion")

                        If f_termina_ref < Now.Date Then
                            dt_grafica_semana.Rows(j).Cells(i).Style.BackColor = Color.Red
                        Else
                            dt_grafica_semana.Rows(j).Cells(i).Style.BackColor = Color.Green
                        End If
                    End If

                    hora_inicio_grafica = hora_inicio_grafica.AddMinutes(30)
                    j += 1
                End While
            End While

            fecha_encabezado = fecha_encabezado.AddDays(1)
        Next

    End Sub
    Function diasemana_minuscula(ByVal dia)
        Select Case dia
            Case 1 : Return "lunes"
            Case 2 : Return "martes"
            Case 3 : Return "miercoles"
            Case 4 : Return "jueves"
            Case 5 : Return "viernes"
            Case 6 : Return "sabado"
            Case 7 : Return "domingo"
        End Select
        Return "naDA"

    End Function
    Function diasemana(ByVal dia)
        Select Case dia
            Case 1 : Return "LUNES"
            Case 2 : Return "MARTES"
            Case 3 : Return "MIERCOLES"
            Case 4 : Return "JUEVES"
            Case 5 : Return "VIERNES"
            Case 6 : Return "SABADO"
            Case 7 : Return "DOMINGO"
        End Select
        Return "naDA"

    End Function


    Private Sub dt_grafica_semana_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dt_grafica_semana.CellContentClick

    End Sub

    Private Sub Dtambiente_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Dtambiente.CellContentClick

    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        Dim rowNumber As Integer = 1
        Dim hora As Date
        ' hora.TimeOfDay.Hours = 6



        While hora < "23:00"
            ' MsgBox(hora)
            DateAdd("n", 30, hora)
        End While


        '*****************************************************Lunes

        sql = "Select * from programacion where "
        conectado()
        cmd = New SqlCommand(sql, cnn)
        reader = cmd.ExecuteReader

        While reader.Read

        End While

        reader.Close()
        cerrar_conexion()
    End Sub

    Private Sub Dtpfecha_ValueChanged(sender As Object, e As EventArgs) Handles Dtpfecha.ValueChanged

    End Sub
End Class