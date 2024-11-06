Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Office.Interop


Public Class Form3
    Dim zona1, zona2 As String



    Private Sub Form3_EnabledChanged(sender As Object, e As EventArgs) Handles Me.EnabledChanged
        If Me.Enabled Then
            actualizar_grid()
        End If
    End Sub
    Sub actualizar_grid()
        If chzonasur.Checked Then

        End If

        Try
            sql = "Select   T1.[id],T1.[curso],T1.[ficha] ,T1.[competencia],T1.[Duracion],T1.[iniciada],T1.[terminada],T1.[cesionada]"
            sql += ",T1.[hora_programada],T1.[fecha_de_inicio],T1.[fecha_no_ejecutable],T1.[fecha_de_terminacion],T1.[instructor]"
            sql += ",T1.[hlunes_iniciada],T1.[hlunes_terminada],T1.[hlunes],T1.[ambiente_lunes]"
            sql += ",T1.[hmartes_iniciada],T1.[hmartes_terminada],T1.[hmartes],T1.[ambiente_martes]"
            sql += ",T1.[hmiercoles_iniciada],T1.[hmiercoles_terminada],T1.[hmiercoles],T1.[ambiente_miercoles]"
            sql += ",T1.[hjueves_iniciada],T1.[hjueves_terminada],T1.[hjueves],T1.[ambiente_jueves]"
            sql += ",T1.[hviernes_iniciada],T1.[hviernes_terminada],T1.[hviernes],T1.[ambiente_viernes]"
            sql += ",T1.[hsabado_iniciada],T1.[hsabado_terminada],T1.[hsabado],T1.[ambiente_sabado]"
            sql += ",T1.[hdomingo_iniciada],T1.[hdomingo_terminada],T1.[hdomingo],T1.[ambiente_domingo]"
            sql += ",T1.[fecha_programacion],T1.[programado_por],T1.[fecha_de_registro_sofia],T1.[registrado_por],T1.[estado_de_registro]"
            sql += ",T1.[motivo_de_error],T1.[Aviso_terminacion],T1.[Fecha_aviso]"
            sql += " from dbo.programacion T1 INNER JOIN dbo.instructores T2 on T1.instructor = t2.NOMBRE_FUNCIONARIO where (t2.zona= '" & zona1 & "' or t2.zona= '" & zona2 & "' or t2.Zona IS NULL) and (T1.iniciada=1 or T1.terminada=1 ) and  (T1.estado_de_registro= 'Sin registrar' or T1.estado_de_registro= 'Correjido') and T1.hlunes_iniciada <> 'sist'"


            conectado()
            da = New SqlClient.SqlDataAdapter(sql, cnn)
            cb = New SqlClient.SqlCommandBuilder(da)
            ds = New DataSet

            da.Fill(ds, "programacion")
            DataGridView1.DataSource = ds
            DataGridView1.DataMember = "programacion"

            ' DataGridView1.Columns("curso").Width = 300
            '  DataGridView1.Columns("competencia").Width = 300
            cerrar_conexion()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try




    End Sub

    Private Sub Form3_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        Me.Finalize()
        Application.ExitThread()

    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        COMPRUEBA_ZONA()
        actualizar_grid()
    End Sub

    Sub llenagrid()
        Try
            da = New SqlClient.SqlDataAdapter(sql, cnn)
            cb = New SqlClient.SqlCommandBuilder(da)
            ds = New DataSet

            da.Fill(ds, "programacion")
            DataGridView1.DataSource = ds
            DataGridView1.DataMember = "programacion"
            'DataGridView1.Columns("Habilitado").Visible = False
            ' DataGridView1.Columns("ficha").Visible = False
            'DataGridView1.Columns("curso").Visible = False
            DataGridView1.Columns("curso").Width = 300
            DataGridView1.Columns("competencia").Width = 300
            cerrar_conexion()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim FILA As Integer = DataGridView1.CurrentRow.Index.ToString

        lblidcompetencia.Text = DataGridView1.Rows(FILA).Cells(0).Value.ToString()
        txtcurso.Text = DataGridView1.Rows(FILA).Cells(1).Value.ToString()
        txtficha.Text = DataGridView1.Rows(FILA).Cells(2).Value.ToString()
        txtcompetencia_programacion.Text = DataGridView1.Rows(FILA).Cells(3).Value.ToString()
        txt_horas_a_ejecutar.Text = DataGridView1.Rows(FILA).Cells(8).Value.ToString()
        Dtpfechadeinicio.Value = DataGridView1.Rows(FILA).Cells(9).Value.ToString()
        txtnoeje.Text = DataGridView1.Rows(FILA).Cells(10).Value.ToString()
        Dtpfechafin.Value = DataGridView1.Rows(FILA).Cells(11).Value.ToString()
        txtinstructor.Text = DataGridView1.Rows(FILA).Cells(12).Value.ToString()

        lbliniciolunes.Text = DataGridView1.Rows(FILA).Cells(13).Value.ToString()
        lblfinlunes.Text = DataGridView1.Rows(FILA).Cells(14).Value.ToString()
        lblHorasLunes.Text = DataGridView1.Rows(FILA).Cells(15).Value.ToString()
        lblAmbienteLunes.Text = DataGridView1.Rows(FILA).Cells(16).Value.ToString()

        lbliniciomartes.Text = DataGridView1.Rows(FILA).Cells(17).Value.ToString()
        lblFinMartes.Text = DataGridView1.Rows(FILA).Cells(18).Value.ToString()
        lblHorasMartes.Text = DataGridView1.Rows(FILA).Cells(19).Value.ToString()
        lblAmbienteMartes.Text = DataGridView1.Rows(FILA).Cells(20).Value.ToString()

        lblInicioMiercoles.Text = DataGridView1.Rows(FILA).Cells(21).Value.ToString()
        lblFinMiercoles.Text = DataGridView1.Rows(FILA).Cells(22).Value.ToString()
        lblHorasMiercoles.Text = DataGridView1.Rows(FILA).Cells(23).Value.ToString()
        lblAmbienteMiercoles.Text = DataGridView1.Rows(FILA).Cells(24).Value.ToString()

        lblInicioJueves.Text = DataGridView1.Rows(FILA).Cells(25).Value.ToString()
        lblFinJueves.Text = DataGridView1.Rows(FILA).Cells(26).Value.ToString()
        lblHorasJueves.Text = DataGridView1.Rows(FILA).Cells(27).Value.ToString()
        lblAmbienteJueves.Text = DataGridView1.Rows(FILA).Cells(28).Value.ToString()

        lblInicioViernes.Text = DataGridView1.Rows(FILA).Cells(29).Value.ToString()
        lblFinViernes.Text = DataGridView1.Rows(FILA).Cells(30).Value.ToString()
        lblHorasViernes.Text = DataGridView1.Rows(FILA).Cells(31).Value.ToString()
        lblAmbienteViernes.Text = DataGridView1.Rows(FILA).Cells(32).Value.ToString()

        lblInicioSabado.Text = DataGridView1.Rows(FILA).Cells(33).Value.ToString()
        lblFinSabado.Text = DataGridView1.Rows(FILA).Cells(34).Value.ToString()
        lblHorasSabado.Text = DataGridView1.Rows(FILA).Cells(35).Value.ToString()
        lblAmbienteSabado.Text = DataGridView1.Rows(FILA).Cells(36).Value.ToString()


        lblInicioDomingo.Text = DataGridView1.Rows(FILA).Cells(37).Value.ToString()
        lblFinDomingo.Text = DataGridView1.Rows(FILA).Cells(38).Value.ToString()
        lblHorasDomingo.Text = DataGridView1.Rows(FILA).Cells(39).Value.ToString()
        lblAmbienteDomingo.Text = DataGridView1.Rows(FILA).Cells(40).Value.ToString()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Enabled = False
        fmrinputbox.Show()

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Try
            sql = "update dbo.programacion set estado_de_registro= 'Ok', registrado_por= '" & lblusuariosofia.Text & "', fecha_de_registro_sofia= '" & Now.Date & "' Where id = " & lblidcompetencia.Text & ""
            conectado()
            cmd = New SqlCommand(sql, cnn)
            cmd.ExecuteNonQuery()
            cerrar_conexion()

            MsgBox("Programacion, Registrada con Exito")

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try


        sql = "Select * from dbo.instructores where NOMBRE_FUNCIONARIO= '" & txtinstructor.Text & "'"
        conectado()
        cmd = New SqlCommand(sql, cnn)
        reader = cmd.ExecuteReader

        If reader.Read Then
            If reader("Correo") = "" Then
            Else
                para = reader("Correo")
            End If

        End If
        reader.Close()

        asunto = "Registro de programacion Sofia plus"
        cuerpo = " <p><strong>Se&ntilde;or(a):</strong> <br />" & txtinstructor.Text & ".</p> <p> Se le ha programado la competencia " & txtcompetencia_programacion.Text & "Desde el dia " & Dtpfechadeinicio.Value.Date & ", Hasta " & Dtpfechafin.Value.Date & ", En el Programa de formacion " & txtcurso.Text & ", Con ficha: " & txtficha.Text
        cuerpo += " </p>"
        cuerpo += " <p> Por favor verifique que la información sea correcta, de lo contrario acercarse a la oficina de Sofia Plus en el menor tiempo posible para que sea corregida. "


        cuerpo = "Señor " & txtinstructor.Text & ", se le ha registrado en el aplicativo Sofia Plus la programacion correspondiente a la competencia " & txtcompetencia_programacion.Text & "Desde el dia " & Dtpfechadeinicio.Value.Date & ", Hasta " & Dtpfechafin.Value.Date & ", En el Programa de formacion " & txtcurso.Text & ", Con ficha: " & txtficha.Text
        cuerpo += vbCrLf

        cuerpo += vbCrLf
        cuerpo += vbCrLf

        cuerpo += " <p> Por favor verifique que la informaci&oacute;n sea correcta, de lo contrario acercarse en el menor tiempo posible a la coordinaci&oacute;n acad&eacute;mica para que sea corregida."

        cuerpo += "</p>"
        cuerpo += " <p></p>"
        cuerpo += " <p></p>"

        cuerpo += " <p> Cordialmente: </p>"

        cuerpo += " <p>    " & lblusuariosofia.Text
        cuerpo += " <br />    Coordinador Academico"
        cuerpo += "</p>"

        adjunto = ""
        Dim emisor As String = "cordinacionagroempresarial@gmail.com"
        Dim pass As String = "gtkpfeyahjkgjnyr"

        enviarCorreo(emisor, pass, cuerpo, asunto, para, adjunto)

        actualizar_grid()
    End Sub

    Private Sub Form3_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        DataGridView1.Width = Me.Width - 200
        DataGridView1.Height = Me.Height - 500
    End Sub
    Dim XLApp As Excel.Application  'Aplicación Excel en varaible XLApp
    Dim XLBook As Excel.Workbook    'Libro de Excel en variable XLBook
    Dim XLSheet As Excel.Worksheet  'Hoja de cálculo en variable XLSheet
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        XLApp = CreateObject("Excel.Application")
        XLBook = XLApp.Workbooks.Open(My.Computer.FileSystem.CurrentDirectory & "\FICHA2.XLS")
        XLSheet = XLBook.Worksheets(1)
        XLSheet.Name = txtficha.Text
        XLApp.Visible = True
        crealibro()
    End Sub

    Sub crealibro()
        sql = "Select * from dbo.grupos where ficha='" & txtficha.Text & "'"
        conectado()
        cmd = New SqlCommand(sql, cnn)
        reader = cmd.ExecuteReader

        If reader.Read Then
            Dim ficha As String = reader("Ficha")
            XLSheet.Cells(8, 2).Value = reader("Ficha")                  'COLOCA LA FICHA EN EL CAMPO DESIGNADO
            Dim fecha_inicio As Date = reader("Fecha_inicio")
            XLSheet.Range("K8").Value = fecha_inicio.Day
            XLSheet.Range("M8").Value = fecha_inicio.Month
            XLSheet.Range("O8").Value = fecha_inicio.Year

            Dim fecha_final As Date = reader("Fecha_terminacion")
            XLSheet.Range("V8").Value = fecha_final.Day
            XLSheet.Range("X8").Value = fecha_final.Month
            XLSheet.Range("Z8").Value = fecha_final.Year

            XLSheet.Range("AG8").Value = reader("Aprendices_matriculados")

            XLSheet.Range("AQ8").Value = reader("Aprendices_activos")

            XLSheet.Range("L10").Value = reader("Nivel") & " " & reader("Nombre_curso")

            XLSheet.Range("AQ10").Value = reader("codigo_programa") & " -V" & reader("Version")
            XLSheet.Range("H12").Value = reader("Municipio")
            XLSheet.Range("Z12").Value = reader("Lugar")
            XLSheet.Range("N17").Value = reader("Instructor_responsable")
        End If



        'XLSheet.Range("A23:A24").EntireRow.Copy()
        ' XLSheet.Range("B25").EntireRow.Insert()

        Try

            Dim contador As Integer = 23
            Dim NUMERO_COMPETENCIA As Integer = 1

            Dim fila As Integer


            sql = "Select * from dbo.programacion where ficha='" & txtficha.Text & "'"
            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader

            While reader.Read
                fila = contador

                XLSheet.Range("B" & contador & ":B" & contador + 1).EntireRow.Copy()


                contador += 2
                XLSheet.Range("B" & contador).EntireRow.Insert()
                XLSheet.Range("B" & fila).Value = NUMERO_COMPETENCIA
                NUMERO_COMPETENCIA += 1

                XLSheet.Range("C" & fila).Value = reader("competencia")
                XLSheet.Range("W" & fila).Value = reader("Duracion")
                'fecha_inicio = reader("fecha_de_inicio")
                XLSheet.Range("AA" & fila).Value = reader("fecha_de_inicio")
                XLSheet.Range("AG" & fila).Value = reader("fecha_de_terminacion")
                XLSheet.Range("AM" & fila).Value = reader("instructor")
                If reader("iniciada") Then
                    XLSheet.Range("C" & fila).Interior.Color = Color.MediumAquamarine
                    XLSheet.Range("W" & fila).Interior.Color = Color.MediumAquamarine
                    XLSheet.Range("AA" & fila).Interior.Color = Color.MediumAquamarine
                    XLSheet.Range("AG" & fila).Interior.Color = Color.MediumAquamarine
                    XLSheet.Range("AM" & fila).Interior.Color = Color.MediumAquamarine
                    XLSheet.Range("AU" & fila).Interior.Color = Color.MediumAquamarine
                End If

                If reader("terminada") Then
                    XLSheet.Range("C" & fila).Interior.Color = Color.Yellow
                    XLSheet.Range("W" & fila).Interior.Color = Color.Yellow
                    XLSheet.Range("AA" & fila).Interior.Color = Color.Yellow
                    XLSheet.Range("AG" & fila).Interior.Color = Color.Yellow
                    XLSheet.Range("AM" & fila).Interior.Color = Color.Yellow
                    XLSheet.Range("AU" & fila).Interior.Color = Color.Yellow
                End If



                XLSheet.Range("AX" & fila).Value = reader("hdomingo_iniciada")
                XLSheet.Range("AX" & fila + 1).Value = reader("ambiente_domingo")
                XLSheet.Range("AY" & fila).Value = reader("hdomingo_terminada")


                XLSheet.Range("AZ" & fila).Value = reader("hlunes_iniciada")
                XLSheet.Range("AZ" & fila + 1).Value = reader("ambiente_lunes")
                XLSheet.Range("BA" & fila).Value = reader("hlunes_terminada")

                XLSheet.Range("BB" & fila).Value = reader("hmartes_iniciada")
                XLSheet.Range("BB" & fila + 1).Value = reader("ambiente_martes")
                XLSheet.Range("BC" & fila).Value = reader("hmartes_terminada")

                XLSheet.Range("BD" & fila).Value = reader("hmiercoles_iniciada")
                XLSheet.Range("BD" & fila + 1).Value = reader("ambiente_miercoles")
                XLSheet.Range("BE" & fila).Value = reader("hmiercoles_terminada")

                XLSheet.Range("BF" & fila).Value = reader("hjueves_iniciada")
                XLSheet.Range("BF" & fila + 1).Value = reader("ambiente_jueves")
                XLSheet.Range("BG" & fila).Value = reader("hjueves_terminada")

                XLSheet.Range("BH" & fila).Value = reader("hviernes_iniciada")
                XLSheet.Range("BH" & fila + 1).Value = reader("ambiente_viernes")
                XLSheet.Range("BI" & fila).Value = reader("hviernes_terminada")

                XLSheet.Range("BJ" & fila).Value = reader("hsabado_iniciada")
                XLSheet.Range("BJ" & fila + 1).Value = reader("ambiente_sabado")
                XLSheet.Range("BK" & fila).Value = reader("hsabado_terminada")


            End While
            XLSheet.Range("B" & contador & ":B" & contador + 1).EntireRow.Delete()
            Dim i As Integer
            Dim suma As Integer = 0
            For i = 23 To contador - 2 Step 2
                suma += XLSheet.Range("W" & i).Value
            Next
            XLSheet.Range("W" & contador).Value = suma
            XLApp.Application.DisplayAlerts = False
            XLBook.SaveAs("C:\academico\" & "Book" & txtficha.Text & ".xls")

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        XLApp = CreateObject("Excel.Application")
        XLBook = XLApp.Workbooks.Open(My.Computer.FileSystem.CurrentDirectory & "\BITACORA1.XLSX")
        XLSheet = XLBook.Worksheets(1)
        XLSheet.Name = "BITACORA"
        XLApp.Visible = True
        DateTimePicker1.Value = New DateTime(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month, DateTimePicker1.Value.Day, 0, 0, 1)
        DateTimePicker2.Value = New DateTime(DateTimePicker2.Value.Year, DateTimePicker2.Value.Month, DateTimePicker2.Value.Day, 23, 59, 59)

        XLSheet.Range("C4").Value = DateTimePicker1.Value
        XLSheet.Range("E4").Value = DateTimePicker2.Value


        Try

            Dim contador As Integer = 7


            Dim fila As Integer

            Dim FECHA_TERM As Date = DateTimePicker2.Value.AddDays(1)



            sql = "Select * from dbo.programacion where fecha_de_registro_sofia >= '" & DateTimePicker1.Value.Date & "' AND fecha_de_registro_sofia <= '" & DateTimePicker2.Value.Date & "' and registrado_por= '" & lblusuariosofia.Text & "'"
            Clipboard.SetText(sql)
            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader

            While reader.Read
                fila = contador

                XLSheet.Range("A" & contador & ":B" & contador + 1).EntireRow.Copy()


                contador += 2

                XLSheet.Range("A" & contador).EntireRow.Insert()

                XLSheet.Range("A" & fila).Value = reader("FICHA")
                XLSheet.Range("B" & fila).Value = reader("curso")
                XLSheet.Range("C" & fila).Value = reader("competencia")
                XLSheet.Range("D" & fila).Value = reader("instructor")
                XLSheet.Range("E" & fila).Value = reader("fecha_de_registro_sofia")
                XLSheet.Range("F" & fila).Value = reader("fecha_de_inicio")
                XLSheet.Range("G" & fila).Value = reader("fecha_de_terminacion")



            End While


            XLApp.Application.DisplayAlerts = False
            XLBook.SaveAs("C:\academico\" & "BITACORA_" & Now.Day & "-" & Now.Month & "-" & Now.Year & " " & Now.Hour & "_" & Now.Minute & "_" & Now.Second & ".xlsx")

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try
    End Sub

    Private Sub CambiarContraseñaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CambiarContraseñaToolStripMenuItem.Click

    End Sub

    Private Sub UsuarioToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UsuarioToolStripMenuItem.Click


        LoginForm2.Show()
    End Sub

    Private Sub chezonnorte_CheckedChanged(sender As Object, e As EventArgs) Handles chezonnorte.CheckedChanged
        COMPRUEBA_ZONA()
    End Sub
    Sub COMPRUEBA_ZONA()
        If chezonnorte.Checked And chzonasur.Checked Then
            lblzona.Text = "NORTE Y SUR"
            zona1 = "Norte"
            zona2 = "Sur"
        ElseIf chezonnorte.Checked And chzonasur.Checked = 0 Then
            lblzona.Text = "NORTE"
            zona1 = "Norte"
            zona2 = "Norte"
        ElseIf chezonnorte.Checked = 0 And chzonasur.Checked Then
            lblzona.Text = "SUR"
            zona1 = "Sur"
            zona2 = "Sur"
        Else
            lblzona.Text = "NULL"
            zona1 = "Null"
            zona2 = "Null"
        End If
        actualizar_grid()
    End Sub

    Private Sub chzonasur_CheckedChanged(sender As Object, e As EventArgs) Handles chzonasur.CheckedChanged
        COMPRUEBA_ZONA()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        COMPRUEBA_ZONA()
        actualizar_grid()
    End Sub
End Class