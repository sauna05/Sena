Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Text.RegularExpressions

Public Class Form1
    Dim datagrid As String
    Public Property IsReadOnly As Boolean
    Dim horas As Decimal
    Dim dias_no_habil As String
    Dim libro_adjunto As String
    Dim iniciada, terminada, secionada As String
    Dim FILA As Integer
    Dim horaF, horaI, horaCI, horaCT As DateTime

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Txtbuscar.Text = "" Then
            llenagrid()
            GoTo salir

        End If
        If Rdbtnbuscar_ficha.Checked Then
            buscar()
        Else
            selecciona()


        End If
salir:
    End Sub

    Sub llenagridcompe()


        da = New SqlClient.SqlDataAdapter(sql, cnn)
        cb = New SqlClient.SqlCommandBuilder(da)
        ds = New DataSet
        da.Fill(ds, "compe")
        DataGridView2.DataSource = ds
        DataGridView2.DataMember = "compe"
        'DataGridView1.Columns("Habilitado").Visible = False
        ' DataGridView1.Columns("estado votacion").Visible = False
        ' DataGridView1.Columns("voto").Visible = False

        DataGridView2.Columns("id").Width = 30
        DataGridView2.Columns("curso").Width = 200
        DataGridView2.Columns("competencia").Width = 600

    End Sub



    Sub llenagrid()
        Dim cuenta_filas As Integer

        da = New SqlClient.SqlDataAdapter(sql, cnn)
        cb = New SqlClient.SqlCommandBuilder(da)
        ds = New DataSet
        If datagrid = "grupos" Then


            da.Fill(ds, "grupos")
            DataGridView1.DataSource = ds
            DataGridView1.DataMember = "grupos"
            'DataGridView1.Columns("Habilitado").Visible = False
            ' DataGridView1.Columns("estado votacion").Visible = False
            ' DataGridView1.Columns("voto").Visible = False
            DataGridView1.Columns("Nombre_curso").Width = 600
            cuenta_filas = DataGridView1.RowCount.ToString
            For i = 0 To cuenta_filas - 1
                If DataGridView1.Rows(i).Cells("Fecha_terminacion").Value < Now.Date Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                End If
            Next

            cerrar_conexion()


        ElseIf datagrid = "programacion" Then
            da.Fill(ds, "programacion")
            DataGridView1.DataSource = ds
            DataGridView1.DataMember = "programacion"
            'DataGridView1.Columns("Habilitado").Visible = False
            'DataGridView1.Columns("ficha").Visible = False
            'DataGridView1.Columns("curso").Visible = False
            DataGridView1.Columns("competencia").Width = 300


            cerrar_conexion()




            cuenta_filas = DataGridView1.RowCount.ToString
            Dim i As Integer
            For i = 0 To cuenta_filas - 1
                If DataGridView1.Rows(i).Cells("iniciada").Value Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                End If
            Next

            For i = 0 To cuenta_filas - 1
                If DataGridView1.Rows(i).Cells("terminada").Value Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    If IsDBNull(DataGridView1.Rows(i).Cells("Evaluado").Value) Then

                    Else

                        If DataGridView1.Rows(i).Cells("Evaluado").Value Then
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Green
                        End If

                    End If
                End If
            Next



        ElseIf datagrid = "competencia" Then
            da.Fill(ds, "competencia")
            DataGridView1.DataSource = ds
            DataGridView1.DataMember = "competencia"
            'DataGridView1.Columns("Habilitado").Visible = False
            DataGridView1.Columns("ficha").Visible = False
            DataGridView1.Columns("nombre_del_curso").Visible = False
            DataGridView1.Columns("competencia").Width = DataGridView1.Size.Width - 500
            DataGridView1.Columns("competencia").FillWeight = 300

            cerrar_conexion()

        ElseIf datagrid = "instructores" Then
            da.Fill(ds, "competencia")
            Dtginstructor.DataSource = ds
            Dtginstructor.DataMember = "competencia"
            'DataGridView1.Columns("Habilitado").Visible = False

            Dtginstructor.Columns("NOMBRE_FUNCIONARIO").Width = 300
            Dtginstructor.Columns("Correo").Width = 300


            cerrar_conexion()



        End If
        Try
            sql = "SELECT COUNT(T1.estado_de_registro) AS Expr1 from  programacion T1 INNER JOIN  instructores T2 on T1.instructor = t2.NOMBRE_FUNCIONARIO"
            sql += " where t2.zona= 'norte' and (T1.iniciada=1 or T1.terminada=1 ) and  (T1.estado_de_registro= 'Sin registrar' or T1.estado_de_registro= 'Correjido') and T1.hlunes_iniciada <> 'sist'"
            conectado()
            cmd = New SqlCommand(sql, cnn)

            reader = cmd.ExecuteReader
            If reader.Read Then
                lblsinregistrarNORTE.Text = reader("Expr1")
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        cerrar_conexion()
        Try
            sql = "SELECT COUNT(T1.estado_de_registro) AS Expr1 from  programacion T1 INNER JOIN  instructores T2 on T1.instructor = t2.NOMBRE_FUNCIONARIO"
            sql += " where t2.zona= 'Sur' and (T1.iniciada=1 or T1.terminada=1 ) and  (T1.estado_de_registro= 'Sin registrar' or T1.estado_de_registro= 'Correjido') and T1.hlunes_iniciada <> 'sist'"
            conectado()
            cmd = New SqlCommand(sql, cnn)

            reader = cmd.ExecuteReader
            If reader.Read Then
                lblsinregistrarsur.Text = reader("Expr1")
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        cerrar_conexion()

        Try
            sql = "SELECT COUNT(T1.estado_de_registro) AS Expr1 from  programacion T1 INNER JOIN  instructores T2 on T1.instructor = t2.NOMBRE_FUNCIONARIO"
            sql += " where t2.zona IS NULL and (T1.iniciada=1 or T1.terminada=1 ) and  (T1.estado_de_registro= 'Sin registrar' or T1.estado_de_registro= 'Correjido') and T1.hlunes_iniciada <> 'sist'"
            conectado()
            cmd = New SqlCommand(sql, cnn)

            reader = cmd.ExecuteReader
            If reader.Read Then
                lblsinregistrarNA.Text = reader("Expr1")
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        cerrar_conexion()

        Try
            sql = "SELECT COUNT(estado_de_registro) AS Expr1 from  programacion where (iniciada=1 or terminada=1) and (estado_de_registro= 'Sin registrar' or estado_de_registro= 'Correjido') and hlunes_iniciada <> 'sist'"
            conectado()
            cmd = New SqlCommand(sql, cnn)

            reader = cmd.ExecuteReader
            If reader.Read Then
                lblsinregistrar.Text = reader("Expr1")
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        cerrar_conexion()

        Try

            sql = "SELECT COUNT(estado_de_registro) AS Expr1 from  programacion where estado_de_registro= 'Error'"
            conectado()
            cmd = New SqlCommand(sql, cnn)

            reader = cmd.ExecuteReader
            If reader.Read Then
                lblerror.Text = reader("Expr1")
            End If
            cerrar_conexion()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub



    Sub selecciona()
        Dim FILA As Integer = DataGridView1.CurrentRow.Index.ToString

        txtficha.Text = DataGridView1.Rows(FILA).Cells(0).Value.ToString()
        txtcodprograma.Text = DataGridView1.Rows(FILA).Cells(1).Value.ToString()
        txtversion.Text = DataGridView1.Rows(FILA).Cells(2).Value.ToString()
        txtcurso.Text = DataGridView1.Rows(FILA).Cells(3).Value.ToString()
        txtnivel.Text = DataGridView1.Rows(FILA).Cells(4).Value.ToString()
        txtinicio.Text = DataGridView1.Rows(FILA).Cells(5).Value.Date()
        txtfin.Text = DataGridView1.Rows(FILA).Cells(6).Value.Date()
        txtmatriculados.Text = DataGridView1.Rows(FILA).Cells(7).Value.ToString()
        txtactivos.Text = DataGridView1.Rows(FILA).Cells(8).Value.ToString()
        txtlugar.Text = DataGridView1.Rows(FILA).Cells(9).Value.ToString()
        txtmunicipio.Text = DataGridView1.Rows(FILA).Cells(10).Value.ToString()
        txtinstructor.Text = DataGridView1.Rows(FILA).Cells(11).Value.ToString()
        txtproyecto.Text = DataGridView1.Rows(FILA).Cells("proyecto").Value.ToString()
        txtcodproyecto.Text = DataGridView1.Rows(FILA).Cells(14).Value.ToString()


        Try
            sql = "Select * from  proyecto where ficha= " & txtficha.Text

            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader

            If reader.Read Then
                txtproyecto.Text = reader("nombre")
                txtcodproyecto.Text = reader("codigo")
            End If

            reader.Close()
            cerrar_conexion()

        Catch ex As Exception

        End Try


        GroupBox2.Enabled = True

        ' DataGridView1.Enabled = False



    End Sub

    Sub buscar()

        sql = "Select * from  grupos where ficha=" & Txtbuscar.Text
        conectado()
        datagrid = "grupos"
        llenagrid()



        If DataGridView1.RowCount = 0 Then

            ret = MsgBox("El curso no existe")

            GoTo sale
        Else
            selecciona()

        End If
sale:
    End Sub

    Sub inicializar()

        limpiar_lbl()
        Button11.Enabled = False
        gbprogcompetencia.Visible = False
        GroupBox3.Visible = False
        txtdias_ejecutar.Text = ""

        btneliminarcompe_programacion.Enabled = False
        btneliminarcompetencia.Enabled = False
        sql = "Select * from  instructores"
        conectado()
        datagrid = "instructores"
        llenagrid()
        GroupBox7.Enabled = False

        GroupBox2.Enabled = False
        sql = "Select top 10 * from  grupos"
        conectado()
        datagrid = "grupos"
        llenagrid()

        carga_todos_los_ambientes()
        muestra_todos_los_ambientes()



        If txtficha.Text = "" Then
            rbtomar.Enabled = False
        End If
        sql = "Select * from  calendario where año=" & Cbaño.Text
        conectado()
        llenagrid2()


        Try
            sql = "Select * from  instructores ORDER BY NOMBRE_FUNCIONARIO ASC"

            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader

            While reader.Read
                ComboBox4.Items.Add(reader("NOMBRE_FUNCIONARIO"))
            End While
            reader.Close()
            cerrar_conexion()

        Catch ex As Exception

        End Try

        Try
            sql = "Select * from  municipios where departamento='44' ORDER BY Municipio ASC"

            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader
            ComboBox5.Items.Add("TODOS LOS MUNICIPIOS")
            ComboBox8.Items.Add("TODOS LOS MUNICIPIOS")
            ComboBox9.Items.Add("TODOS LOS MUNICIPIOS")
            While reader.Read
                ComboBox5.Items.Add(reader("Municipio"))
                ComboBox8.Items.Add(reader("Municipio"))
                ComboBox9.Items.Add(reader("Municipio"))
            End While
            ComboBox5.SelectedItem = ComboBox5.Items.Item(0)
            ComboBox8.SelectedItem = ComboBox8.Items.Item(0)
            '    ComboBox9.SelectedItem = ComboBox9.Items.Item(0)
            reader.Close()
            cerrar_conexion()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        ComboBox7.SelectedItem = ComboBox7.Items.Item(0)
        carga_todos_los_ambientes()

        sql = "Select * from  programacion"
        conectado()
        datagrid = "compe"
        llenagridcompe()

        CheckBox4.Checked = 1
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Finalize()
        Application.ExitThread()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtfechainicio.Value = "01/01/" & Now.Year
        sql = "select top 10 CONCAT(ficha,' | ',Nombre_curso) as combo, ficha from grupos"
        llenarcombos(cmbgrupos, "combo", "ficha")

        mostrar_aprendices_ficha()

        inicializar()



    End Sub
    Sub carga_todos_los_ambientes()
        Try
            sql = "Select * from ambientes ORDER BY ambiente ASC"
            'llenarcombos(ComboBox3, "ambiente", "ID")
            'llenarcombos(ComboBox6, "ambiente", "ID")
            completa_ambiente()
            sql = "Select * from  ambientes"
            ' LlenarDataGrids(DataGridView3)
        Catch ex As Exception

        End Try
    End Sub
    Sub completa_ambiente()
        conectado()
        cmd = New SqlCommand(sql, cnn)
        reader = cmd.ExecuteReader
        ComboBox6.Items.Add("TODOS LOS AMBIENTES")
        While reader.Read
            ComboBox3.Items.Add(reader("ambiente"))

            ComboBox6.Items.Add(reader("ambiente"))
        End While
        ComboBox3.SelectedItem = ComboBox3.Items.Item(0)
        ComboBox6.SelectedItem = ComboBox6.Items.Item(0)
        reader.Close()
        cerrar_conexion()
    End Sub
    Sub muestra_todos_los_ambientes()
        sql = "Select * from  ambientes"
        grid_ambiente()
    End Sub
    Sub grid_ambiente()
        conectado()
        da = New SqlClient.SqlDataAdapter(sql, cnn)
        cb = New SqlClient.SqlCommandBuilder(da)
        ds = New DataSet
        da.Fill(ds, "ambientes")
        DataGridView3.DataSource = ds
        DataGridView3.DataMember = "ambientes"
        'DataGridView1.Columns("Habilitado").Visible = False
        ' DataGridView3.Columns("ambiente").Width = 400

        cerrar_conexion()
    End Sub

    Private Sub Rdbtnbuscar_nombre_CheckedChanged(sender As Object, e As EventArgs) Handles Rdbtnbuscar_nombre.CheckedChanged
        Button1.Enabled = False
        lblbuscar.Text = "Curso: "

    End Sub

    Private Sub Rdbtnbuscar_ficha_CheckedChanged(sender As Object, e As EventArgs) Handles Rdbtnbuscar_ficha.CheckedChanged
        lblbuscar.Text = "Ficha: "
        Button1.Enabled = True
    End Sub

    Private Sub Txtbuscar_TextChanged(sender As Object, e As EventArgs) Handles Txtbuscar.TextChanged
        If Rdbtnbuscar_nombre.Checked Then
            sql = "Select top 50 * from  grupos where Nombre_curso LIKE '%" + Txtbuscar.Text + "%' ORDER BY Fecha_inicio DESC"
            conectado()
            datagrid = "grupos"

            llenagrid()
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        FILA = DataGridView1.CurrentRow.Index.ToString

        If datagrid = "copia_masiva_origen" Then
            Dim ficha_origen As Integer = DataGridView1.Rows(FILA).Cells("ficha").Value
            Dim filas, i As Integer

            filas = DataGridView1.RowCount - 2

            Dim ficha As Integer
            Dim competencia(filas), duracion(filas) As String

            sql = "Select * from  competencia where ficha='" & ficha_origen & "'"
            conectado()

            da = New SqlClient.SqlDataAdapter(sql, cnn)

            ds = New DataSet

            da.Fill(ds, "grupos")
            Dim dt As New DataTable
            dt = New DataTable
            dt = ds.Tables(0)


            Dim j, dura As Integer
            Dim compe As String
            Dim codigo_pro As String = DataGridView1.Rows(FILA).Cells("codigo_programa").Value
            Dim version As String = DataGridView1.Rows(FILA).Cells("version").Value
            Dim curso As String = DataGridView1.Rows(FILA).Cells("Nombre_curso").Value

            For i = 0 To filas

                If DataGridView1.Rows(i).Cells(0).Value And DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Gray Then
                    ficha = DataGridView1.Rows(i).Cells("ficha").Value

                    For j = 0 To dt.Rows.Count - 1
                        maximo_id_competencia()
                        compe = dt.Rows(j).Item(3)
                        dura = dt.Rows(j).Item(4)
                        sql = "insert into  competencia (id, cod_programa, ficha, competencia, duracion, nombre_del_curso) values ( " & maximo + 1 & ", " & codigo_pro & ", '" & ficha & "','" & compe & "', " & dura & ", '" & curso & "')"

                        conectado()
                        cmd = New SqlCommand(sql, cnn)
                        cmd.ExecuteNonQuery()
                        cerrar_conexion()

                        maximo_id_programacion()
                        sql = "insert into  programacion (id, ficha, competencia,  duracion, curso, estado_de_registro, Aviso_terminacion, iniciada, terminada, cesionada) values (" & maximo + 1 & ", '" & ficha & "', '" & compe & "', " & dura & ", '" & curso & "', 'Sin registrar',0 ,0 ,0 ,0)"
                        conectado()
                        cmd = New SqlCommand(sql, cnn)
                        cmd.ExecuteNonQuery()
                        cerrar_conexion()

                    Next
                End If
            Next

            MsgBox(" Copia masiva exitosa")

            sql = "Select * from  grupos"
            conectado()
            datagrid = "grupos"
            llenagrid()
            cerrar_conexion()

            DataGridView1.Columns.Remove("Seleccionar")
            datagrid = "oculta"
            GoTo salir
        End If

        If datagrid = "oculta" Then
            sql = "Select * from  grupos"
            conectado()
            datagrid = "grupos"
            llenagrid()
            cerrar_conexion()
            GoTo salir
        End If

        If datagrid = "grupos" Then
            selecciona()

            ' datagrid = "grupos"
        End If

        If datagrid = "competencia_otro_curso" Then

            sql = "Select * from  competencia where ficha=" & DataGridView1.Rows(FILA).Cells(0).Value.ToString()
            conectado()
            datagrid = "competencia"
            llenagrid()
            datagrid = "competencia_otro_curso1"
        End If

        If datagrid = "competencia_otro_curso1" Then

            If DataGridView1.RowCount > 1 Then

                txtcompetencia.Text = DataGridView1.Rows(FILA).Cells(3).Value.ToString()
                txthoras.Text = DataGridView1.Rows(FILA).Cells(4).Value.ToString()
                id_competencia.Text = DataGridView1.Rows(FILA).Cells(0).Value.ToString()
                btneliminarcompetencia.Enabled = True
                datagrid = "competencia_otro_curso1"
            Else
                MsgBox("El curso no tiene compretencias asignadas")

            End If

        End If

        If datagrid = "Todo_otro_curso" Then

            sql = "Select * from  competencia where ficha=" & DataGridView1.Rows(FILA).Cells(0).Value.ToString()
            conectado()
            datagrid = "competencia"
            llenagrid()
            datagrid = "competencia_otro_curso2"



        End If

        If datagrid = "competencia" Then
            txtcompetencia.Text = DataGridView1.Rows(FILA).Cells(3).Value.ToString()
            txthoras.Text = DataGridView1.Rows(FILA).Cells(4).Value.ToString()
            id_competencia.Text = DataGridView1.Rows(FILA).Cells(0).Value.ToString()
            btneliminarcompetencia.Enabled = True
        End If

        If datagrid = "programacion" Then
            ' lblidcompetencia.Text = DataGridView1.Rows(FILA).Cells(0).Value.ToString()
            ' txtcompetencia_programacion.Text = DataGridView1.Rows(FILA).Cells(3).Value.ToString()
            ' txthoras_programar.Text = DataGridView1.Rows(FILA).Cells(4).Value.ToString()
            gbprogcompetencia.Enabled = True
            btneliminarcompe_programacion.Enabled = True

            LLENALBL()



        End If

        If datagrid = "error" Then
            fmrinputbox.Show()

            fmrinputbox.txterror.Text = DataGridView1.Rows(FILA).Cells("ficha").Value.ToString() & ": " & DataGridView1.Rows(FILA).Cells(46).Value.ToString()

            fmrinputbox.Button1.Enabled = False
            fmrinputbox.Button2.Enabled = False
            '****************************************************************
            '****************************************************************
            gbprogcompetencia.Visible = True
            GroupBox3.Visible = False
            gbprogcompetencia.Enabled = True
            limpiar_lbl()

            txtficha.Text = DataGridView1.Rows(FILA).Cells(2).Value.ToString()

            sql = "Select * from  grupos where ficha=" & txtficha.Text
            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader

            If reader.Read Then
                txtcodprograma.Text = reader("codigo_programa")
                txtversion.Text = reader("Version")
                txtmatriculados.Text = reader("Aprendices_matriculados")
                txtactivos.Text = reader("Aprendices_activos")
                txtnivel.Text = reader("Nivel")
                txtinicio.Text = reader("Fecha_inicio").date
                txtfin.Text = reader("Fecha_terminacion").date
                txtmunicipio.Text = reader("Municipio")
                txtlugar.Text = reader("Lugar")
            End If
            txtcurso.Text = DataGridView1.Rows(FILA).Cells(1).Value.ToString()
            txtficha.Text = DataGridView1.Rows(FILA).Cells(2).Value.ToString()
            txtinstructor.Text = DataGridView1.Rows(FILA).Cells(12).Value.ToString()
            LLENALBL()

        End If
salir:
    End Sub

    Sub LLENALBL()
        limpiar_lbl()
        btnsentxls.Enabled = False
        lblidcompetencia.Text = DataGridView1.Rows(FILA).Cells(0).Value.ToString()


        txtcompetencia_programacion.Text = DataGridView1.Rows(FILA).Cells(3).Value.ToString()
        txt_horas_a_ejecutar.Text = DataGridView1.Rows(FILA).Cells(8).Value.ToString()

        If DataGridView1.Rows(FILA).Cells(9).Value.Equals(DBNull.Value) Then
            Dtpfechadeinicio.Value = Now
            Dtpfechafin.Value = Now
            btnsentxls.Enabled = False

        Else
            Dtpfechadeinicio.Value = DataGridView1.Rows(FILA).Cells(9).Value.ToString()
            Dtpfechafin.Value = DataGridView1.Rows(FILA).Cells(11).Value.ToString()
            btnsentxls.Enabled = True
        End If



        ComboBox4.Text = DataGridView1.Rows(FILA).Cells(12).Value.ToString()


        txthoras_programar.Text = DataGridView1.Rows(FILA).Cells(4).Value.ToString()
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


        '****************************************************************
        '****************************************************************
    End Sub

    Private Sub rbadministrar_CheckedChanged(sender As Object, e As EventArgs) Handles rbadministrar.CheckedChanged
        btnbuscarcodprograma.Enabled = False
    End Sub
    '*************************************tomar competencias de otro programa de formacion
    Private Sub rbtomar_CheckedChanged(sender As Object, e As EventArgs) Handles rbtomar.CheckedChanged
        btnbuscarcodprograma.Enabled = True
        btnactualizarcompetencia.Enabled = False
        btneliminarcompetencia.Enabled = False

        sql = "Select * from  grupos where codigo_programa=" & txtcodprograma.Text
        conectado()
        datagrid = "grupos"
        llenagrid()

        datagrid = "competencia_otro_curso"

sale:



    End Sub
    Private Sub rbtomartodo_CheckedChanged(sender As Object, e As EventArgs) Handles rbtomartodo.CheckedChanged
        btnbuscarcodprograma.Enabled = True
        btnactualizarcompetencia.Enabled = False
        btneliminarcompetencia.Enabled = False
        rbtodo()

    End Sub

    Sub rbtodo()
        sql = "Select * from  grupos where codigo_programa=" & txtcodprograma.Text & " and version= " & txtversion.Text
        conectado()
        datagrid = "grupos"
        llenagrid()
        datagrid = "Todo_otro_curso"
    End Sub



    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        gbprogcompetencia.Visible = False
        GroupBox3.Visible = True

        sql = "Select * from  competencia where ficha=" & txtficha.Text
        conectado()
        datagrid = "competencia"
        llenagrid()


    End Sub

    Private Sub btnagregarcompetencia_Click(sender As Object, e As EventArgs) Handles btnagregarcompetencia.Click
        If txtcompetencia.Text = "" Or txthoras.Text = "" Then
            MsgBox("Debe especificar competencia y duracion")


        Else
            maximo_id_competencia()
            Try

                sql = "insert into  competencia (id, cod_programa, ficha, competencia, duracion, nombre_del_curso) values ( " & maximo + 1 & ", " & txtcodprograma.Text & ", '" & txtficha.Text & "','" & txtcompetencia.Text & "', " & txthoras.Text & ", '" & txtcurso.Text & "')"
                ', '" & txtficha.Text & "', '" & txtcompetencia.Text & "', 
                conectado()
                cmd = New SqlCommand(sql, cnn)
                cmd.ExecuteNonQuery()
                cerrar_conexion()
                MsgBox("competencia creada exitosamente")
                maximo_id_programacion()
                sql = "insert into  programacion (id, ficha, iniciada, competencia, duracion, curso, estado_de_registro, Aviso_terminacion, terminada, cesionada) values ( " & maximo + 1 & ", '" & txtficha.Text & "', 0, '" & txtcompetencia.Text & "', " & txthoras.Text & ", '" & txtcurso.Text & "', 'Sin registrar', 0 ,0 ,0)"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                cmd.ExecuteNonQuery()
                cerrar_conexion()
                MsgBox("competencia agregada a programacion exitosamente")

                sql = "Select * from  programacion where ficha=" & txtficha.Text
                conectado()
                datagrid = "programacion"
                llenagrid()
                DataGridView1.Enabled = True
                limpiarcompetencia()
            Catch ex As Exception
                MsgBox(ex.ToString)

            End Try
        End If
    End Sub
    Sub limpiarcompetencia()
        txtcompetencia.Text = ""
        txthoras.Text = ""
        txtcompetencia.Select()
    End Sub
    '*************************************agrega competencia a un programa de formacion

    Private Sub btnbuscarcodprograma_Click(sender As Object, e As EventArgs) Handles btnbuscarcodprograma.Click
        If datagrid = "competencia_otro_curso2" Then
            Dim i As Integer
            Dim compe, dura As String
            For i = 0 To DataGridView1.RowCount - 2

                maximo_id_competencia()



                Try
                    compe = DataGridView1.Rows(i).Cells(3).Value.ToString()
                    dura = DataGridView1.Rows(i).Cells(4).Value.ToString()
                    sql = "insert into competencia (id, cod_programa, ficha, competencia, duracion, nombre_del_curso) values ( " & maximo + 1 & ", " & txtcodprograma.Text & ", '" & txtficha.Text & "','" & compe & "', " & dura & ", '" & txtcurso.Text & "')"
                    ', '" & txtficha.Text & "', '" & txtcompetencia.Text & "', 
                    conectado()
                    cmd = New SqlCommand(sql, cnn)
                    cmd.ExecuteNonQuery()
                    cerrar_conexion()


                    maximo_id_programacion()

                    sql = "insert into programacion (id, ficha, competencia, duracion, curso, estado_de_registro, Aviso_terminacion, iniciada, terminada, cesionada) values ( " & maximo + 1 & ", '" & txtficha.Text & "', '" & compe & "', " & dura & ", '" & txtcurso.Text & "', 'Sin registrar', 0, 0, 0, 0)"
                    conectado()
                    cmd = New SqlCommand(sql, cnn)
                    cmd.ExecuteNonQuery()
                    cerrar_conexion()




                Catch ex As Exception
                    MsgBox(ex.ToString)

                End Try

            Next
            MsgBox("todas las competencias fueron copiadas exitosamente")
            sql = "Select * from  programacion where ficha=" & txtficha.Text
            conectado()
            datagrid = "programacion"
            llenagrid()

        Else



            If txtcompetencia.Text = "" Or txthoras.Text = "" Then
                MsgBox("Debe especificar competencia y duracion")


            Else
                Try

                    sql = "insert into  competencia (id, cod_programa, ficha, competencia, duracion, nombre_del_curso) values (" & maximo + 1 & ", " & txtcodprograma.Text & ", '" & txtficha.Text & "','" & txtcompetencia.Text & "', " & txthoras.Text & ", '" & txtcurso.Text & "')"
                    ', '" & txtficha.Text & "', '" & txtcompetencia.Text & "', 
                    conectado()
                    cmd = New SqlCommand(sql, cnn)
                    cmd.ExecuteNonQuery()
                    cerrar_conexion()

                    maximo_id_programacion()
                    sql = "insert into  programacion (id, ficha, competencia, duracion, curso, estado_de_registro, Aviso_terminacion, iniciada, terminada, cesionada) values (" & maximo + 1 & ", '" & txtficha.Text & "', '" & txtcompetencia.Text & "', " & txthoras.Text & ", '" & txtcurso.Text & "', 'Sin registrar', 0, 0, 0, 0)"
                    conectado()
                    cmd = New SqlCommand(sql, cnn)
                    cmd.ExecuteNonQuery()
                    cerrar_conexion()



                    MsgBox("competencia copiada exitosamente")
                Catch ex As Exception
                    MsgBox(ex.ToString)

                End Try

            End If
        End If

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        sql = "Select * from  programacion where ficha=" & txtficha.Text
        conectado()
        datagrid = "programacion"
        llenagrid()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        sql = "Select * from  competencia where ficha=" & txtficha.Text
        conectado()
        datagrid = "competencia"
        llenagrid()
    End Sub

    Private Sub txtficha_TextChanged(sender As Object, e As EventArgs) Handles txtficha.TextChanged
        If txtficha.Text <> "" Then
            rbtomar.Enabled = True
        End If

    End Sub
    '***************************************************************************************************************************************************
    '***************************************************************************************************************************************************
    '********************PAGINA 4- CALENDARIO ACADEMICO

    Sub llenagrid2()

        da = New SqlClient.SqlDataAdapter(sql, cnn)
        cb = New SqlClient.SqlCommandBuilder(da)
        ds = New DataSet



        da.Fill(ds, "calendario")
        datacalendario.DataSource = ds
        datacalendario.DataMember = "calendario"

        cerrar_conexion()


    End Sub



    '********************************************************************************
    Private Sub TabPage4_Click(sender As Object, e As EventArgs) Handles TabPage4.Click


    End Sub



    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim año, mes, dia As Integer
        Dim fecha As Date
        Dim motivo As String


        año = DateTimePicker1.Value.Year
        mes = DateTimePicker1.Value.Month
        dia = DateTimePicker1.Value.Day
        motivo = txtmotivo.Text
        fecha = DateTimePicker1.Value.Date
        'MsgBox(año & ", " & mes & ", " & dia & ", " & fecha)

        Try

            sql = "insert into  calendario (año, mes, dia, fecha_completa, motivo) values ( " & año & ", " & mes & "," & dia & ", '" & fecha & "', '" & motivo & "')"
            Clipboard.SetText(sql)
            conectado()
            cmd = New SqlCommand(sql, cnn)
            cmd.ExecuteNonQuery()
            cerrar_conexion()
            MsgBox("fecha agregada exitosamente")
            sql = "Select * from  calendario where año=" & Cbaño.Text
            conectado()
            llenagrid2()
        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try
    End Sub



    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Dim rpt As Integer = MessageBox.Show("¿Está seguro de que desea eliminar esta fecha?", "Advertencia", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
        If rpt = 6 Then
            Dim FILA As String = datacalendario.CurrentRow.Cells(3).Value.ToString

            sql = "delete from  calendario where fecha_completa= '" & FILA & "'"
            conectado()
            cmd = New SqlCommand(sql, cnn)
            cmd.ExecuteNonQuery()
            cerrar_conexion()
            MsgBox("fecha eliminada exitosamente")
            sql = "Select * from calendario where año=" & Cbaño.Text
            conectado()
            llenagrid2()
        Else
            Exit Sub
        End If



    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        sql = "Select * from  calendario where año=" & Cbaño.Text
        conectado()
        llenagrid2()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        sql = "Select * from  calendario where mes=" & cbmes.Text
        conectado()
        llenagrid2()
    End Sub





    Sub limpiar_lbl()

        lbliniciolunes.Text = ""
        lblfinlunes.Text = ""
        lblAmbienteLunes.Text = ""
        lblHorasLunes.Text = ""

        lbliniciomartes.Text = ""
        lblInicioMiercoles.Text = ""
        lblInicioJueves.Text = ""
        lblInicioViernes.Text = ""
        lblInicioSabado.Text = ""
        lblInicioDomingo.Text = ""


        lblFinMartes.Text = ""
        lblFinMiercoles.Text = ""
        lblFinJueves.Text = ""
        lblFinViernes.Text = ""
        lblFinSabado.Text = ""
        lblFinDomingo.Text = ""

        lblAmbienteMartes.Text = ""
        lblAmbienteMiercoles.Text = ""
        lblAmbienteJueves.Text = ""
        lblAmbienteViernes.Text = ""
        lblAmbienteSabado.Text = ""
        lblAmbienteDomingo.Text = ""

        lblHorasMartes.Text = ""
        lblHorasMiercoles.Text = ""
        lblHorasJueves.Text = ""
        lblHorasViernes.Text = ""
        lblHorasSabado.Text = ""
        lblHorasDomingo.Text = ""

    End Sub


    '*************************************Agregar horario y ambiente a los label
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Button14.Enabled = True
        Button16.Enabled = True

        If cmbhorainicio.Text = cmbhorafin.Text Then
            MessageBox.Show("Las horas son iguales, No se ejecutará ninguna cantidad de horas", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        Else
            If cruce_ambiente() = 0 Then
                GoTo Continuar
            Else
                Exit Sub
            End If

            GoTo Continuar
        End If

Continuar:


        'ValidacionAmbienteConsulta()

        resta_horas()

        If Chlunes.Checked Then
            lbliniciolunes.Text = cmbhorainicio.Text
            lblfinlunes.Text = cmbhorafin.Text
            lblAmbienteLunes.Text = ComboBox3.Text
            lblHorasLunes.Text = horas
        End If

        If ChMartes.Checked Then
            lbliniciomartes.Text = cmbhorainicio.Text
            lblFinMartes.Text = cmbhorafin.Text
            lblAmbienteMartes.Text = ComboBox3.Text
            lblHorasMartes.Text = horas
        End If

        If ChMiercoles.Checked Then
            lblInicioMiercoles.Text = cmbhorainicio.Text
            lblFinMiercoles.Text = cmbhorafin.Text
            lblAmbienteMiercoles.Text = ComboBox3.Text
            lblHorasMiercoles.Text = horas
        End If

        If ChJueves.Checked Then
            lblInicioJueves.Text = cmbhorainicio.Text
            lblFinJueves.Text = cmbhorafin.Text
            lblAmbienteJueves.Text = ComboBox3.Text
            lblHorasJueves.Text = horas
        End If

        If ChViernes.Checked Then
            lblInicioViernes.Text = cmbhorainicio.Text
            lblFinViernes.Text = cmbhorafin.Text
            lblAmbienteViernes.Text = ComboBox3.Text
            lblHorasViernes.Text = horas
        End If

        If ChSabado.Checked Then
            lblInicioSabado.Text = cmbhorainicio.Text
            lblFinSabado.Text = cmbhorafin.Text
            lblAmbienteSabado.Text = ComboBox3.Text
            lblHorasSabado.Text = horas
        End If

        If ChDomingo.Checked Then
            lblInicioDomingo.Text = cmbhorainicio.Text
            lblFinDomingo.Text = cmbhorafin.Text
            lblAmbienteDomingo.Text = ComboBox3.Text
            lblHorasDomingo.Text = horas
        End If
salir:
    End Sub

    Sub resta_horas()
        Dim a As DateTime
        Dim b As DateTime
        Dim c As Integer
        Dim d As Decimal

        a = cmbhorainicio.Text
        b = cmbhorafin.Text
        c = b.Hour - a.Hour
        d = b.Minute - a.Minute
        If d < 0 Then
            c = c - 1
            d = d * (-1)

        End If

        d = d / 60
        If c < 0 Then
            MsgBox("La Hora de inicio es mayor que la fin", MsgBoxStyle.Critical)
        Else : horas = c + d

        End If

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        limpiar_lbl()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        gbprogcompetencia.Visible = True
        gbprogcompetencia.Enabled = False
        GroupBox3.Visible = False
        adaptar()
        limpiar_lbl()

        sql = "Select * from  programacion where ficha=" & txtficha.Text
        conectado()
        datagrid = "programacion"
        llenagrid()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        dias_no_habil = ""
        Dim horas_a_programar As Integer = txthoras_programar.Text
        Dim fecha_de_inicio As Date = Dtpfechadeinicio.Value.Date
        Dim fecha_de_fin_compe As Date
        Dim dia_semana As Integer
        Dim horas_a_ejecutar As Double = 0
        Dim dias_formacion As Integer





        While horas_a_programar > 0

            Try
                sql = "Select * from  calendario where fecha_completa='" & fecha_de_inicio & "'"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                If reader.Read Then
                    dias_no_habil = dias_no_habil & " - " & reader("fecha_completa")
                Else

                    dia_semana = fecha_de_inicio.DayOfWeek
                    Select Case dia_semana
                        Case 0          'Domingo
                            If lblHorasDomingo.Text <> "" Then
                                horas_a_programar -= lblHorasDomingo.Text
                                horas_a_ejecutar += lblHorasDomingo.Text
                                fecha_de_fin_compe = fecha_de_inicio
                                dias_formacion += 1
                            End If
                        Case 1
                            If lblHorasLunes.Text <> "" Then
                                horas_a_programar -= lblHorasLunes.Text
                                horas_a_ejecutar += lblHorasLunes.Text
                                fecha_de_fin_compe = fecha_de_inicio
                                dias_formacion += 1
                            End If
                        Case 2
                            If lblHorasMartes.Text <> "" Then
                                horas_a_programar -= lblHorasMartes.Text
                                horas_a_ejecutar += lblHorasMartes.Text
                                fecha_de_fin_compe = fecha_de_inicio
                                dias_formacion += 1
                            End If
                        Case 3
                            If lblHorasMiercoles.Text <> "" Then
                                horas_a_programar -= lblHorasMiercoles.Text
                                horas_a_ejecutar += lblHorasMiercoles.Text
                                fecha_de_fin_compe = fecha_de_inicio
                                dias_formacion += 1
                            End If
                        Case 4
                            If lblHorasJueves.Text <> "" Then
                                horas_a_programar -= lblHorasJueves.Text
                                horas_a_ejecutar += lblHorasJueves.Text
                                fecha_de_fin_compe = fecha_de_inicio
                                dias_formacion += 1
                            End If
                        Case 5
                            If lblHorasViernes.Text <> "" Then
                                horas_a_programar -= lblHorasViernes.Text
                                horas_a_ejecutar += lblHorasViernes.Text
                                fecha_de_fin_compe = fecha_de_inicio
                                dias_formacion += 1
                            End If
                        Case 6          'Sabado
                            If lblHorasSabado.Text <> "" Then
                                horas_a_programar -= lblHorasSabado.Text
                                horas_a_ejecutar += lblHorasSabado.Text
                                fecha_de_fin_compe = fecha_de_inicio
                                dias_formacion += 1
                            End If

                    End Select
                    reader.Close()
                End If

            Catch ex As Exception
                MsgBox(ex.ToString)

            End Try
            cerrar_conexion()




            If horas_a_programar > 0 Then
                fecha_de_inicio = fecha_de_inicio.AddDays(1)
                If fecha_de_inicio.Year > Now.Year Then
                    MsgBox("La competencia no se alcanza a terminar en esta vigencia se ejecutarian: " & horas_a_ejecutar & ", ¿Desea seccionarla? ", MsgBoxStyle.YesNo)
                    'programar seccionamiento de competencia'
                    fecha_de_inicio = fecha_de_inicio.AddDays(-1)
                    GoTo salir_calc
                End If
            End If



        End While
salir_calc:
        Dtpfechafin.Value = fecha_de_fin_compe
        txt_horas_a_ejecutar.Text = horas_a_ejecutar
        Button11.Enabled = True
        txtdias_ejecutar.Text = dias_formacion

        'ValidacionAmbienteConsulta()
        comprueba_cruce()


    End Sub

    Sub comprueba_cruce()
        '***********************************************************************
        '*******************CODIGO DE COMPROBACION DE CRUCE DE HORARIO



        Try

            sql = "Select * from  programacion where ficha='" & txtficha.Text & "'"
            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader

            While reader.Read
                If (reader("iniciada") Or reader("terminada")) And (Not IsDBNull(reader("hora_programada"))) Then

                    If reader("id") = lblidcompetencia.Text Then

                    ElseIf (reader("fecha_de_inicio") < Dtpfechadeinicio.Value.Date And reader("fecha_de_terminacion") < Dtpfechadeinicio.Value.Date) Or (reader("fecha_de_inicio") > Dtpfechafin.Value.Date And reader("fecha_de_terminacion") > Dtpfechafin.Value.Date) Then
                    Else

                        If reader("hlunes_iniciada") <> "" And lbliniciolunes.Text <> "" Then

                            cruce_hora(reader("hlunes_iniciada"), reader("hlunes_terminada"), lbliniciolunes.Text, lblfinlunes.Text)

                            If (hora_completa_fin_pro <= hora_completa_ini_pro) Then
                                MsgBox("Las horas de inicio deben ser mayores que las de terminacion", MsgBoxStyle.Critical)
                                Button11.Enabled = False
                                GoTo salir
                            End If

                            If (hora_completa_fin_pro <= hora_completa_ini_bd) Or (hora_completa_ini_pro >= hora_completa_fin_bd) Then

                            Else
                                MsgBox("Existe un cruce de horario el dia lunes con la competencia " & reader("competencia"), MsgBoxStyle.Critical)
                                Button11.Enabled = False
                                GoTo salir
                            End If
                        End If

                        If reader("hmartes_iniciada") <> "" And lbliniciomartes.Text <> "" Then
                            cruce_hora(reader("hmartes_iniciada"), reader("hmartes_terminada"), lbliniciomartes.Text, lblFinMartes.Text)
                            If (hora_completa_fin_pro <= hora_completa_ini_pro) Then
                                MsgBox("Las horas de inicio deben ser mayores que las de terminacion", MsgBoxStyle.Critical)
                                Button11.Enabled = False
                                GoTo salir
                            End If

                            If (hora_completa_fin_pro <= hora_completa_ini_bd) Or (hora_completa_ini_pro >= hora_completa_fin_bd) Then
                            Else
                                MsgBox("Existe un cruce de horario el dia martes con la competencia " & reader("competencia"), MsgBoxStyle.Critical)
                                Button11.Enabled = False
                                GoTo salir
                            End If
                        End If


                        If reader("hmiercoles_iniciada") <> "" And lblInicioMiercoles.Text <> "" Then
                            cruce_hora(reader("hmiercoles_iniciada"), reader("hmiercoles_terminada"), lblInicioMiercoles.Text, lblFinMiercoles.Text)
                            If (hora_completa_fin_pro <= hora_completa_ini_pro) Then
                                MsgBox("Las horas de inicio deben ser mayores que las de terminacion", MsgBoxStyle.Critical)
                                Button11.Enabled = False
                                GoTo salir
                            End If

                            If (hora_completa_fin_pro <= hora_completa_ini_bd) Or (hora_completa_ini_pro >= hora_completa_fin_bd) Then
                            Else
                                MsgBox("Existe un cruce de horario el dia miercoles con la competencia " & reader("competencia"), MsgBoxStyle.Critical)
                                Button11.Enabled = False
                                GoTo salir
                            End If
                        End If


                        If reader("hjueves_iniciada") <> "" And lblInicioJueves.Text <> "" Then
                            cruce_hora(reader("hjueves_iniciada"), reader("hjueves_terminada"), lblInicioJueves.Text, lblFinJueves.Text)

                            If (hora_completa_fin_pro <= hora_completa_ini_pro) Then
                                MsgBox("Las horas de inicio deben ser mayores que las de terminacion", MsgBoxStyle.Critical)
                                Button11.Enabled = False
                                GoTo salir
                            End If

                            If (hora_completa_fin_pro <= hora_completa_ini_bd) Or (hora_completa_ini_pro >= hora_completa_fin_bd) Then
                            Else
                                MsgBox("Existe un cruce de horario el dia jueves con la competencia " & reader("competencia"), MsgBoxStyle.Critical)
                                Button11.Enabled = False
                                GoTo salir
                            End If
                        End If

                        If reader("hviernes_iniciada") <> "" And lblInicioViernes.Text <> "" Then
                            cruce_hora(reader("hviernes_iniciada"), reader("hviernes_terminada"), lblInicioViernes.Text, lblFinViernes.Text)
                            If (hora_completa_fin_pro <= hora_completa_ini_pro) Then
                                MsgBox("Las horas de inicio deben ser mayores que las de terminacion", MsgBoxStyle.Critical)
                                Button11.Enabled = False
                                GoTo salir
                            End If

                            If (hora_completa_fin_pro <= hora_completa_ini_bd) Or (hora_completa_ini_pro >= hora_completa_fin_bd) Then
                            Else
                                MsgBox("Existe un cruce de horario el dia viernes con la competencia " & reader("competencia"), MsgBoxStyle.Critical)
                                Button11.Enabled = False
                                GoTo salir
                            End If
                        End If


                        If reader("hsabado_iniciada") <> "" And lblInicioSabado.Text <> "" Then
                            cruce_hora(reader("hsabado_iniciada"), reader("hsabado_terminada"), lblInicioSabado.Text, lblFinSabado.Text)
                            If (hora_completa_fin_pro <= hora_completa_ini_pro) Then
                                MsgBox("Las horas de inicio deben ser mayores que las de terminacion", MsgBoxStyle.Critical)
                                Button11.Enabled = False
                                GoTo salir
                            End If

                            If (hora_completa_fin_pro <= hora_completa_ini_bd) Or (hora_completa_ini_pro >= hora_completa_fin_bd) Then
                            Else
                                MsgBox("Existe un cruce de horario el dia sabado con la competencia " & reader("competencia"), MsgBoxStyle.Critical)
                                Button11.Enabled = False
                                GoTo salir
                            End If
                        End If

                        If reader("hdomingo_iniciada") <> "" And lblInicioDomingo.Text <> "" Then
                            cruce_hora(reader("hdomingo_iniciada"), reader("hdomingo_terminada"), lblInicioDomingo.Text, lblFinDomingo.Text)
                            If (hora_completa_fin_pro <= hora_completa_ini_pro) Then
                                MsgBox("Las horas de inicio deben ser mayores que las de terminacion", MsgBoxStyle.Critical)
                                Button11.Enabled = False
                                GoTo salir
                            End If

                            If (hora_completa_fin_pro <= hora_completa_ini_bd) Or (hora_completa_ini_pro >= hora_completa_fin_bd) Then
                            Else
                                MsgBox("Existe un cruce de horario el dia domingo con la competencia " & reader("competencia"), MsgBoxStyle.Critical)
                                Button11.Enabled = False
                                GoTo salir
                            End If
                        End If


                    End If
                End If

            End While
salir:



        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

        '*****************************************************************************************
        CRUCE_INSTRUCTOR()

        If comprueba_ambiente() = 1 Then
            Button11.Enabled = False
            Exit Sub
        End If
      
    End Sub

    Sub CRUCE_INSTRUCTOR()

        Try
            sql = "Select * from  programacion where instructor='" & ComboBox4.Text & "'"
            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader

            While reader.Read


                If reader("id") = lblidcompetencia.Text Then

                ElseIf (reader("fecha_de_inicio") < Dtpfechadeinicio.Value.Date And reader("fecha_de_terminacion") < Dtpfechadeinicio.Value.Date) Or (reader("fecha_de_inicio") > Dtpfechafin.Value.Date And reader("fecha_de_terminacion") > Dtpfechafin.Value.Date) Then
                Else

                    If reader("hlunes_iniciada") <> "" And lbliniciolunes.Text <> "" Then

                        cruce_hora(reader("hlunes_iniciada"), reader("hlunes_terminada"), lbliniciolunes.Text, lblfinlunes.Text)

                        If (hora_completa_fin_pro <= hora_completa_ini_pro) Then
                            MsgBox("Las horas de inicio deben ser mayores que las de terminacion", MsgBoxStyle.Critical)
                            Button11.Enabled = False
                            GoTo salir
                        End If

                        If (hora_completa_fin_pro <= hora_completa_ini_bd) Or (hora_completa_ini_pro >= hora_completa_fin_bd) Then

                        Else
                            MsgBox("Existe un cruce de horario del instructor " & ComboBox4.Text & ", el dia lunes con la competencia " & reader("competencia") & ", EN LA FICHA  " & reader("ficha"), MsgBoxStyle.Critical)
                            Button11.Enabled = False
                            GoTo salir
                        End If
                    End If

                    If reader("hmartes_iniciada") <> "" And lbliniciomartes.Text <> "" Then
                        cruce_hora(reader("hmartes_iniciada"), reader("hmartes_terminada"), lbliniciomartes.Text, lblFinMartes.Text)
                        If (hora_completa_fin_pro <= hora_completa_ini_pro) Then
                            MsgBox("Las horas de inicio deben ser mayores que las de terminacion", MsgBoxStyle.Critical)
                            Button11.Enabled = False
                            GoTo salir
                        End If

                        If (hora_completa_fin_pro <= hora_completa_ini_bd) Or (hora_completa_ini_pro >= hora_completa_fin_bd) Then
                        Else
                            MsgBox("Existe un cruce de horario del instructor " & ComboBox4.Text & ", el dia martes con la competencia " & reader("competencia") & ", EN LA FICHA  " & reader("ficha"), MsgBoxStyle.Critical)
                            Button11.Enabled = False
                            GoTo salir
                        End If
                    End If


                    If reader("hmiercoles_iniciada") <> "" And lblInicioMiercoles.Text <> "" Then
                        cruce_hora(reader("hmiercoles_iniciada"), reader("hmiercoles_terminada"), lblInicioMiercoles.Text, lblFinMiercoles.Text)
                        If (hora_completa_fin_pro <= hora_completa_ini_pro) Then
                            MsgBox("Las horas de inicio deben ser mayores que las de terminacion", MsgBoxStyle.Critical)
                            Button11.Enabled = False
                            GoTo salir
                        End If

                        If (hora_completa_fin_pro <= hora_completa_ini_bd) Or (hora_completa_ini_pro >= hora_completa_fin_bd) Then
                        Else
                            MsgBox("Existe un cruce de horario del instructor " & ComboBox4.Text & ", el dia miercoles con la competencia " & reader("competencia") & ", EN LA FICHA  " & reader("ficha"), MsgBoxStyle.Critical)
                            Button11.Enabled = False
                            GoTo salir
                        End If
                    End If


                    If reader("hjueves_iniciada") <> "" And lblInicioJueves.Text <> "" Then
                        cruce_hora(reader("hjueves_iniciada"), reader("hjueves_terminada"), lblInicioJueves.Text, lblFinJueves.Text)

                        If (hora_completa_fin_pro <= hora_completa_ini_pro) Then
                            MsgBox("Las horas de inicio deben ser mayores que las de terminacion", MsgBoxStyle.Critical)
                            Button11.Enabled = False
                            GoTo salir
                        End If

                        If (hora_completa_fin_pro <= hora_completa_ini_bd) Or (hora_completa_ini_pro >= hora_completa_fin_bd) Then
                        Else
                            MsgBox("Existe un cruce de horario del instructor " & ComboBox4.Text & ", el dia jueves con la competencia " & reader("competencia") & ", EN LA FICHA  " & reader("ficha"), MsgBoxStyle.Critical)
                            Button11.Enabled = False
                            GoTo salir
                        End If
                    End If

                    If reader("hviernes_iniciada") <> "" And lblInicioViernes.Text <> "" Then
                        cruce_hora(reader("hviernes_iniciada"), reader("hviernes_terminada"), lblInicioViernes.Text, lblFinViernes.Text)
                        If (hora_completa_fin_pro <= hora_completa_ini_pro) Then
                            MsgBox("Las horas de inicio deben ser mayores que las de terminacion", MsgBoxStyle.Critical)
                            Button11.Enabled = False
                            GoTo salir
                        End If

                        If (hora_completa_fin_pro <= hora_completa_ini_bd) Or (hora_completa_ini_pro >= hora_completa_fin_bd) Then
                        Else
                            MsgBox("Existe un cruce de horario del instructor " & ComboBox4.Text & ", el dia viernes con la competencia " & reader("competencia") & ", EN LA FICHA  " & reader("ficha"), MsgBoxStyle.Critical)
                            Button11.Enabled = False
                            GoTo salir
                        End If
                    End If


                    If reader("hsabado_iniciada") <> "" And lblInicioSabado.Text <> "" Then
                        cruce_hora(reader("hsabado_iniciada"), reader("hsabado_terminada"), lblInicioSabado.Text, lblFinSabado.Text)
                        If (hora_completa_fin_pro <= hora_completa_ini_pro) Then
                            MsgBox("Las horas de inicio deben ser mayores que las de terminacion", MsgBoxStyle.Critical)
                            Button11.Enabled = False
                            GoTo salir
                        End If

                        If (hora_completa_fin_pro <= hora_completa_ini_bd) Or (hora_completa_ini_pro >= hora_completa_fin_bd) Then
                        Else
                            MsgBox("Existe un cruce de horario del instructor " & ComboBox4.Text & " el dia sabado con la competencia " & reader("competencia") & ", EN LA FICHA  " & reader("ficha"), MsgBoxStyle.Critical)
                            Button11.Enabled = False
                            GoTo salir
                        End If
                    End If

                    If reader("hdomingo_iniciada") <> "" And lblInicioDomingo.Text <> "" Then
                        cruce_hora(reader("hdomingo_iniciada"), reader("hdomingo_terminada"), lblInicioDomingo.Text, lblFinDomingo.Text)
                        If (hora_completa_fin_pro <= hora_completa_ini_pro) Then
                            MsgBox("Las horas de inicio deben ser mayores que las de terminacion", MsgBoxStyle.Critical)
                            Button11.Enabled = False
                            GoTo salir
                        End If

                        If (hora_completa_fin_pro <= hora_completa_ini_bd) Or (hora_completa_ini_pro >= hora_completa_fin_bd) Then
                        Else
                            MsgBox("Existe un cruce de horario del instructor " & ComboBox4.Text & " el dia domingo con la competencia " & reader("competencia") & ", EN LA FICHA  " & reader("ficha"), MsgBoxStyle.Critical)
                            Button11.Enabled = False
                            GoTo salir
                        End If
                    End If


                End If


            End While
salir:


        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub
    '*********************************************************comprueba cruce de ambiente



    Sub mensaje()

    End Sub

    Dim hora_ini_bd, hora_ini_pro As Integer
    Dim minutos_ini_bd, minutos_ini_pro As Integer
    Dim hora_completa_ini_bd, hora_completa_ini_pro As Decimal
    Dim hora_fin_bd, hora_fin_pro As Integer
    Dim minutos_fin_bd, minutos_fin_pro As Integer
    Dim hora_completa_fin_bd, hora_completa_fin_pro As Decimal

    Sub cruce_hora(ini_bd, fin_bd, ini_pro, fin_pro)

        hora_ini_bd = Mid(ini_bd, 1, 2)
        minutos_ini_bd = Mid(ini_bd, 4, 2)
        hora_completa_ini_bd = minutos_ini_bd / 60
        hora_completa_ini_bd += hora_ini_bd
        hora_fin_bd = Mid(fin_bd, 1, 2)
        minutos_fin_bd = Mid(fin_bd, 4, 2)
        hora_completa_fin_bd = minutos_fin_bd / 60
        hora_completa_fin_bd += hora_fin_bd

        hora_ini_pro = Mid(ini_pro, 1, 2)
        minutos_ini_pro = Mid(ini_pro, 4, 2)
        hora_completa_ini_pro = minutos_ini_pro / 60
        hora_completa_ini_pro += hora_ini_pro
        hora_fin_pro = Mid(fin_pro, 1, 2)
        minutos_fin_pro = Mid(fin_pro, 4, 2)
        hora_completa_fin_pro = minutos_fin_pro / 60
        hora_completa_fin_pro += hora_fin_pro
    End Sub


    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        dias_no_habil = ""
        Dim fecha_de_inicio As Date = Dtpfechadeinicio.Value.Date
        Dim fecha_fin As Date = Dtpfechafin.Value.Date
        Dim horas_a_programar As Double = 0
        Dim dias_formacion As Integer
        Dim dia_semana As Integer
        Try
            While fecha_de_inicio <= fecha_fin


                sql = "Select * from  calendario where fecha_completa='" & fecha_de_inicio & "'"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                If reader.Read Then
                    dias_no_habil = dias_no_habil & " - " & reader("fecha_completa")
                Else


                    dia_semana = fecha_de_inicio.DayOfWeek
                    Select Case dia_semana
                        Case 0          'Domingo
                            If lblHorasDomingo.Text <> "" Then
                                horas_a_programar += lblHorasDomingo.Text
                                dias_formacion += 1
                            End If
                        Case 1
                            If lblHorasLunes.Text <> "" Then
                                horas_a_programar += lblHorasLunes.Text
                                dias_formacion += 1
                            End If
                        Case 2
                            If lblHorasMartes.Text <> "" Then
                                horas_a_programar += lblHorasMartes.Text
                                dias_formacion += 1
                            End If
                        Case 3
                            If lblHorasMiercoles.Text <> "" Then
                                horas_a_programar += lblHorasMiercoles.Text
                                dias_formacion += 1
                            End If
                        Case 4
                            If lblHorasJueves.Text <> "" Then
                                horas_a_programar += lblHorasJueves.Text
                                dias_formacion += 1
                            End If
                        Case 5
                            If lblHorasViernes.Text <> "" Then
                                horas_a_programar += lblHorasViernes.Text
                                dias_formacion += 1
                            End If
                        Case 6          'Sabado
                            If ChSabado.Checked Then
                                If lblHorasSabado.Text <> "" Then
                                    horas_a_programar += lblHorasSabado.Text
                                    dias_formacion += 1
                                Else
                                    MessageBox.Show("Debe agregar nuevamente la programacion de la competencia", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    Exit Sub
                                End If
                                ' Else
                                ' MessageBox.Show("Debe seleccionar el dia sabado del horario", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                ' Exit Sub

                            End If


                    End Select


                    reader.Close()
                End If


                cerrar_conexion()



                If horas_a_programar > 0 Then
                    fecha_de_inicio = fecha_de_inicio.AddDays(1)
                    If fecha_de_inicio.Year > Now.Year Then
                        MsgBox("La competencia no se alcanza a terminar en esta vigencia se ejecutarian: " & horas_a_programar & ", ¿Desea seccionarla? ", MsgBoxStyle.YesNo)
                        'programar seccionamiento de competencia'
                        fecha_de_inicio = fecha_de_inicio.AddDays(-1)
                        GoTo salir_calc
                    End If
                End If




            End While
        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

salir_calc:
        Button11.Enabled = True

        txt_horas_a_ejecutar.Text = horas_a_programar
        txtdias_ejecutar.Text = dias_formacion
        comprueba_cruce()
        'ValidacionAmbienteConsulta()
    End Sub



    '*************************************************************************************************************************************************************
    '*************************************************************************************************************************************************************
    '**************************************************************  INSTRUCTOR  *********************************************************************************

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles btn_buscar_instructor.Click
        If txt_buscarinstructor.Text = "" Then
            datagrid = "instructores"
            llenagrid()
            GoTo salir

        End If

        buscar1()

salir:
    End Sub

    Private Sub RadioButton8_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton8.CheckedChanged
        btn_buscar_instructor.Enabled = False
        Label32.Text = "Nombre: "
    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged
        btn_buscar_instructor.Enabled = True
        Label32.Text = "Cedula: "
    End Sub
    Sub buscar1()

        sql = "Select * from  instructores where NUMERO_IDENTIFICACION_FUNCIONARIO=" & txt_buscarinstructor.Text
        conectado()
        datagrid = "instructores"
        llenagrid()



        If Dtginstructor.RowCount = 1 Then

            ret = MsgBox("El instructor no existe")

            GoTo sale
        Else
            selecciona1()

        End If
sale:
    End Sub
    Sub horasRestantes(ByVal nombreInstructor As String)

        Dim sql As String = "SELECT SUM(CAST(p.hora_programada AS FLOAT)) AS Total_Horas_Cumplidas, i.Horas AS Total_Horas_Asignadas FROM  programacion p JOIN  instructores i ON i.NOMBRE_FUNCIONARIO = p.instructor WHERE p.instructor = @Instructor AND YEAR(p.fecha_de_inicio) = YEAR(GETDATE()) GROUP BY i.NOMBRE_FUNCIONARIO, i.Horas;"

        conectado()

        Dim cmd As New SqlCommand(sql, cnn)
        cmd.Parameters.AddWithValue("@Instructor", nombreInstructor)

        Dim horasRestantes As String = ""
        Dim horasCumplidas As String = ""
        Dim totalHoras As String = ""

        reader = cmd.ExecuteReader
        If (reader.HasRows) Then
            reader.Read()

            totalHoras = reader("Total_Horas_Asignadas")

            tbtHorasCumplidas.ForeColor = Color.Black
            horasCumplidas = reader("Total_Horas_Cumplidas").ToString

            tbHorasRestantes.ForeColor = Color.Black
            horasRestantes = (Integer.Parse(totalHoras) - Integer.Parse(horasCumplidas)).ToString

        Else

            cerrar_conexion()
            tbHorasRestantes.ForeColor = Color.DarkRed
            tbtHorasCumplidas.ForeColor = Color.DarkRed
            horasRestantes = "Error"
            horasCumplidas = "Error"
           

           
        End If
        tbHorasRestantes.Text = horasRestantes
        tbtHorasCumplidas.Text = horasCumplidas

    End Sub

    Sub selecciona1()
        Dim FILA As Integer = Dtginstructor.CurrentRow.Index.ToString

        txtcedula.Text = Dtginstructor.Rows(FILA).Cells(0).Value.ToString()
        txtnombrefuncionario.Text = Dtginstructor.Rows(FILA).Cells(1).Value.ToString()
        horasRestantes(txtnombrefuncionario.Text)
        Cbvinculacion.Text = Dtginstructor.Rows(FILA).Cells(2).Value.ToString()
        Cbzona.Text = Dtginstructor.Rows(FILA).Cells(3).Value.ToString()
        txtemail.Text = Dtginstructor.Rows(FILA).Cells(4).Value.ToString()
        txttelefono.Text = Dtginstructor.Rows(FILA).Cells(5).Value.ToString()
        txtespecialidad.Text = Dtginstructor.Rows(FILA).Cells(6).Value.ToString()
        txtestado.Text = Dtginstructor.Rows(FILA).Cells(7).Value.ToString()
        Dtinicio.Text = Dtginstructor.Rows(FILA).Cells(8).Value.ToString()
        Dtfin.Text = Dtginstructor.Rows(FILA).Cells(9).Value.ToString()
        txthorascontrato.Text = Dtginstructor.Rows(FILA).Cells(10).Value.ToString()
        TXTDIRECCION.Text = Dtginstructor.Rows(FILA).Cells(11).Value.ToString()
        txttelefono2.Text = Dtginstructor.Rows(FILA).Cells(12).Value.ToString()

        txthorasdia.Text = Dtginstructor.Rows(FILA).Cells(14).Value.ToString()
        txtsalario.Text = Dtginstructor.Rows(FILA).Cells(15).Value.ToString()




        GroupBox7.Enabled = True

        ' DataGridView1.Enabled = False



    End Sub

    Private Sub txt_buscarinstructor_TextChanged(sender As Object, e As EventArgs) Handles txt_buscarinstructor.TextChanged
        If RadioButton8.Checked Then
            sql = "Select * from  instructores where NOMBRE_FUNCIONARIO LIKE '%" + txt_buscarinstructor.Text + "%'"
            conectado()
            datagrid = "instructores"

            llenagrid()

        End If
    End Sub

    Private Sub Dtginstructor_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Dtginstructor.CellClick
        selecciona1()
        Button35.Text = "Nuevo"
    End Sub



    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        TabControl1.TabPages.Item(0).Show()

        gbprogcompetencia.Visible = True
        GroupBox3.Visible = False

        limpiar_lbl()
        ComboBox4.Text = txtnombrefuncionario.Text


    End Sub
    '*************************************Actualizar competencia de programacion

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click



        If txtcompetencia_programacion.Text = "" Or txt_horas_a_ejecutar.Text = "" Or txt_horas_a_ejecutar.Text = "0" Or ComboBox4.Text = "" Then
            MsgBox("La programacion no esta completa, Revice e intente de nuevo", MsgBoxStyle.Critical)
        Else
            Try
                sql = "Select * from  instructores where NOMBRE_FUNCIONARIO = '" & ComboBox4.Text & "' and zona is null"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader
                If reader.Read Then
                    Dim rta_zona As Integer
                    rta_zona = MsgBox("El instructor no tiene Zona definida, ¿Desea Definir la Zona?", 4)
                    If rta_zona = 7 Then
                    Else
                        FmrZona.Show()
                        GoTo Salir
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            reader.Close()
            sql = "Select * from  programacion where id='" & lblidcompetencia.Text & "'"
            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader

            If reader.Read Then
                If txt_horas_a_ejecutar.Text < reader("Duracion") Then
                    Dim rta, cantidad As Integer
                    cantidad = reader("Duracion") - txt_horas_a_ejecutar.Text
                    rta = MsgBox("Esta Competencia no sejecutara completamente, ¿desea seccionar con " & cantidad & " Horas, para ejecucion posterior?", 4)
                    If rta = 7 Then

                        sql = "update  programacion set cesionada= 0 Where id = " & lblidcompetencia.Text & ""
                        conectado()

                        cmd = New SqlCommand(sql, cnn)
                        cmd.ExecuteNonQuery()
                        cerrar_conexion()

                    Else
                        '*Debe validar la cantidad de horas a ejecutar, la duracion debe ser entera*'
                        sql = "update programacion set cesionada= 1, duracion= " & txt_horas_a_ejecutar.Text & " Where id = " & lblidcompetencia.Text & ""
                        conectado()
                        ' MsgBox(sql)
                        cmd = New SqlCommand(sql, cnn)
                        cmd.ExecuteNonQuery()
                        cerrar_conexion()
                        maximo_id_programacion()
                        txthoras_programar.Text = txt_horas_a_ejecutar.Text
                        maximo_id_programacion()
                        sql = "insert into  programacion (id, ficha, competencia, iniciada, terminada, cesionada, duracion, curso, estado_de_registro, Aviso_terminacion) values ( " & maximo + 1 & ", '" & txtficha.Text & "', '" & txtcompetencia_programacion.Text & "', 0, 0, 0, " & cantidad & ", '" & txtcurso.Text & "', 'Sin registrar', 0)"
                        conectado()
                        cmd = New SqlCommand(sql, cnn)
                        cmd.ExecuteNonQuery()
                        cerrar_conexion()



                    End If
                End If
            End If

            Dim iniciada, finalizada As Boolean

            If Dtpfechafin.Value >= Now.Date Then
                iniciada = True
                finalizada = False
            Else
                iniciada = False
                finalizada = True

            End If
            Dim fechaini As String = Dtpfechadeinicio.Value.Date
            Dim fechafin As String = Dtpfechafin.Value.Date

            Try



                sql = "update  programacion set "
                sql += " iniciada= 1, terminada= 0, hora_programada='" & txt_horas_a_ejecutar.Text & "', "
                sql += " fecha_de_inicio= '" & fechaini & "', fecha_no_ejecutable= '" & dias_no_habil & "', fecha_de_terminacion= '" & fechafin & "', "
                sql += " instructor= '" & ComboBox4.Text & "',"
                sql += "hlunes_iniciada= '" & lbliniciolunes.Text & "', hlunes_terminada= '" & lblfinlunes.Text & "', hlunes= '" & lblHorasLunes.Text & "', ambiente_lunes= '" & lblAmbienteLunes.Text & "', "
                sql += "hmartes_iniciada= '" & lbliniciomartes.Text & "', hmartes_terminada= '" & lblFinMartes.Text & "', hmartes= '" & lblHorasMartes.Text & "', ambiente_martes= '" & lblAmbienteMartes.Text & "', "
                sql += "hmiercoles_iniciada= '" & lblInicioMiercoles.Text & "', hmiercoles_terminada= '" & lblFinMiercoles.Text & "', hmiercoles= '" & lblHorasMiercoles.Text & "', ambiente_miercoles= '" & lblAmbienteMiercoles.Text & "', "
                sql += "hjueves_iniciada= '" & lblInicioJueves.Text & "', hjueves_terminada= '" & lblFinJueves.Text & "', hjueves= '" & lblHorasJueves.Text & "', ambiente_jueves= '" & lblAmbienteJueves.Text & "', "
                sql += "hviernes_iniciada= '" & lblInicioViernes.Text & "', hviernes_terminada= '" & lblFinViernes.Text & "', hviernes= '" & lblHorasViernes.Text & "', ambiente_viernes= '" & lblAmbienteViernes.Text & "', "
                sql += "hsabado_iniciada= '" & lblInicioSabado.Text & "', hsabado_terminada= '" & lblFinSabado.Text & "', hsabado= '" & lblHorasSabado.Text & "', ambiente_sabado= '" & lblAmbienteSabado.Text & "', "
                sql += "hdomingo_iniciada= '" & lblInicioDomingo.Text & "', hdomingo_terminada= '" & lblFinDomingo.Text & "', hdomingo= '" & lblHorasDomingo.Text & "', ambiente_domingo= '" & lblAmbienteDomingo.Text & "', "
                sql += "fecha_programacion= '" & Now.Date & "', programado_por= '" & lblusuario.Text & "'"
                sql += " Where id = " & lblidcompetencia.Text & ""



                conectado()

                cmd = New SqlCommand(sql, cnn)
                cmd.ExecuteNonQuery()
                cerrar_conexion()

                If datagrid = "error" Then


                    sql = "update  programacion set "
                    sql += " estado_de_registro= 'Correjido', motivo_de_error= 'Correjido'"

                    sql += " Where id = " & lblidcompetencia.Text & ""
                    conectado()
                    cmd = New SqlCommand(sql, cnn)
                    cmd.ExecuteNonQuery()
                    cerrar_conexion()
                    MsgBox("Correccion de error exitosa")

                Else
                    MsgBox("Correcto, programacion exitosa")
                End If

                cuerpo = " <p><strong>Se&ntilde;or(a):</strong> <br />" & ComboBox4.Text & ".</p> <p> Se le ha programado la competencia " & txtcompetencia_programacion.Text & "Desde el dia " & Dtpfechadeinicio.Value.Date & ", Hasta " & Dtpfechafin.Value.Date & ", En el Programa de formacion " & txtcurso.Text & ", Con ficha: " & txtficha.Text
                cuerpo += " </p>"
                cuerpo += " <p> EN EL HORARIO:"

                If lblHorasLunes.Text <> "" Then
                    cuerpo += " <p>"
                    cuerpo += " <br /> LUNES: "
                    cuerpo += " <br /> *   HORA,  Desde: " & lbliniciolunes.Text & "; Hasta: " & lblfinlunes.Text
                    cuerpo += " <br /> *   AMBIENTE: " & lblAmbienteLunes.Text
                    cuerpo += " <br /> *   HORAS: " & lblHorasLunes.Text
                    cuerpo += " </p>"
                End If

                If lblHorasMartes.Text <> "" Then
                    cuerpo += " <p>"
                    cuerpo += " <br /> MARTES: "
                    cuerpo += " <br /> *   HORA,  Desde: " & lbliniciomartes.Text & "; Hasta: " & lblFinMartes.Text
                    cuerpo += " <br /> *   AMBIENTE: " & lblAmbienteMartes.Text
                    cuerpo += " <br /> *   HORAS: " & lblHorasMartes.Text
                    cuerpo += " </p>"
                End If

                If lblHorasMiercoles.Text <> "" Then
                    cuerpo += " <p>"
                    cuerpo += " <br />MIERCOLES: "
                    cuerpo += " <br /> *   HORA,  Desde: " & lblInicioMiercoles.Text & "; Hasta: " & lblFinMiercoles.Text
                    cuerpo += " <br /> *   AMBIENTE: " & lblAmbienteMiercoles.Text
                    cuerpo += " <br /> *   HORAS: " & lblHorasMiercoles.Text
                    cuerpo += " </p>"
                End If

                If lblHorasJueves.Text <> "" Then
                    cuerpo += " <p>"
                    cuerpo += " <br /> JUEVES: "
                    cuerpo += " <br /> *   HORA,  Desde: " & lblInicioJueves.Text & "; Hasta: " & lblFinJueves.Text
                    cuerpo += " <br /> *   AMBIENTE: " & lblAmbienteJueves.Text
                    cuerpo += " <br /> *   HORAS: " & lblHorasJueves.Text
                    cuerpo += " </p>"
                End If

                If lblHorasViernes.Text <> "" Then
                    cuerpo += " <p>"
                    cuerpo += " <br /> VIERNES: "
                    cuerpo += " <br /> *   HORA,  Desde: " & lblInicioViernes.Text & "; Hasta: " & lblFinViernes.Text
                    cuerpo += " <br /> *   AMBIENTE: " & lblAmbienteViernes.Text
                    cuerpo += " <br /> *   HORAS: " & lblHorasViernes.Text
                    cuerpo += " </p>"
                End If

                If lblHorasSabado.Text <> "" Then
                    cuerpo += " <p>"
                    cuerpo += " <br />SABADO: "
                    cuerpo += " <br /> *   HORA,  Desde: " & lblInicioSabado.Text & "; Hasta: " & lblFinSabado.Text
                    cuerpo += " <br /> *   AMBIENTE: " & lblAmbienteSabado.Text
                    cuerpo += " <br /> *   HORAS: " & lblHorasSabado.Text
                    cuerpo += " </p>"
                End If

                If lblHorasDomingo.Text <> "" Then
                    cuerpo += " <p>"
                    cuerpo += " <br />DOMINGO: "
                    cuerpo += " <br /> *   HORA,  Desde: " & lblInicioDomingo.Text & "; Hasta: " & lblFinDomingo.Text
                    cuerpo += " <br /> *   AMBIENTE: " & lblAmbienteDomingo.Text
                    cuerpo += " <br /> *   HORAS: " & lblHorasDomingo.Text
                    cuerpo += " </p>"
                End If


                cuerpo += " <p></p>"
                cuerpo += " <p></p>"

                cuerpo += " <p> Por favor verifique que la informaci&oacute;n sea correcta, de lo contrario acercarse en el menor tiempo posible a la coordinaci&oacute;n acad&eacute;mica para que sea corregida."

                cuerpo += "</p>"
                cuerpo += " <p></p>"
                cuerpo += " <p></p>"

                cuerpo += " <p> Cordialmente: </p>"

                cuerpo += " <p>    " & lblusuario.Text
                cuerpo += " <br />    Coordinador Academico"
                cuerpo += "</p>"

                enviaficha()

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
Salir:


        End If
    End Sub

    Sub enviaficha()
        sql = "Select * from  instructores where NOMBRE_FUNCIONARIO= '" & ComboBox4.Text & "'"
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




      

        asunto = "Programacion Docente"
        XLApp = CreateObject("Excel.Application")
        XLBook = XLApp.Workbooks.Open(My.Computer.FileSystem.CurrentDirectory & "\FICHA2.XLS")
        XLSheet = XLBook.Worksheets(1)
        XLSheet.Name = txtficha.Text
        XLApp.Visible = False
        crealibro()

        XLApp.ActiveWorkbook.Close()
        ' mataexcel()
        adjunto = libro_adjunto



        '***************************************************************************************************************************************************************
        ' enviar_correo()
        '***************************************************************************************************************************************************************
        Dim emisor As String = "cordinacionagroempresarial@gmail.com"
        Dim pass As String = "gtkpfeyahjkgjnyr"

        enviarCorreo(emisor, pass, cuerpo, asunto, para, adjunto)

        limpiar_lbl()
        ComboBox4.Text = ""
        gbprogcompetencia.Enabled = False
        txtdias_ejecutar.Text = ""
        Dtpfechadeinicio.Value = Now
        Dtpfechafin.Value = Now
        txt_horas_a_ejecutar.Text = ""
        txtcompetencia_programacion.Text = ""
        txthoras_programar.Text = ""
        ComboBox3.Text = ""
        cmbhorainicio.SelectedIndex = 0
        cmbhorafin.SelectedIndex = 0
        TextBox1.Text = ""
           

        'limpiar_lbl()
        Button11.Enabled = False
        sql = "Select * from  programacion where ficha=" & txtficha.Text
        conectado()
        datagrid = "programacion"
        llenagrid()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbhorainicio.SelectedIndexChanged
        Button14.Enabled = False
        Button16.Enabled = False

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbhorafin.SelectedIndexChanged
        Button14.Enabled = False
        Button16.Enabled = False
    End Sub


    Sub crealibro()

        XLSheet.Cells(8, 2).Value = txtficha.Text                   'COLOCA LA FICHA EN EL CAMPO DESIGNADO
        Dim fecha_inicio As Date = txtinicio.Text
        XLSheet.Range("K8").Value = fecha_inicio.Day
        XLSheet.Range("M8").Value = fecha_inicio.Month
        XLSheet.Range("O8").Value = fecha_inicio.Year

        Dim fecha_final As Date = txtfin.Text
        XLSheet.Range("V8").Value = fecha_final.Day
        XLSheet.Range("X8").Value = fecha_final.Month
        XLSheet.Range("Z8").Value = fecha_final.Year

        XLSheet.Range("AG8").Value = txtmatriculados.Text

        XLSheet.Range("AQ8").Value = txtactivos.Text

        XLSheet.Range("L10").Value = txtnivel.Text & " " & txtcurso.Text

        XLSheet.Range("AQ10").Value = txtcodprograma.Text & " -V" & txtversion.Text
        XLSheet.Range("H12").Value = txtmunicipio.Text
        XLSheet.Range("Z12").Value = txtlugar.Text
        XLSheet.Range("N17").Value = txtinstructor.Text
        XLSheet.Range("G19").Value = txtcodproyecto.Text & "- " & txtproyecto.Text
        'XLSheet.Range("A23:A24").EntireRow.Copy()
        ' XLSheet.Range("B25").EntireRow.Insert()

        Try

            Dim contador As Integer = 23
            Dim NUMERO_COMPETENCIA As Integer = 1

            Dim fila As Integer


            sql = "Select * from  programacion where ficha='" & txtficha.Text & "'"
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

            Dim carpeta As String = "C:\ProgramacionAcademico\" & ComboBox4.Text & "\" & txtficha.Text
            Dim RutaGuardado As String = "\" & ComboBox4.Text & Now.Day & Now.Month & Now.Year & Now.Hour & lblidcompetencia.Text & ".xls"
            Dim dir As System.IO.DirectoryInfo = New DirectoryInfo(carpeta)
            If dir.Exists Then

                XLBook.SaveAs(carpeta & RutaGuardado)
                libro_adjunto = carpeta & RutaGuardado
            Else
                dir.Create()
                XLBook.SaveAs(carpeta & RutaGuardado)
                libro_adjunto = carpeta & RutaGuardado

            End If
            'XLBook.Visible = True
        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub
    Sub Solocrealibro()

        XLSheet.Cells(8, 2).Value = txtficha.Text                   'COLOCA LA FICHA EN EL CAMPO DESIGNADO
        Dim fecha_inicio As Date = txtinicio.Text
        XLSheet.Range("K8").Value = fecha_inicio.Day
        XLSheet.Range("M8").Value = fecha_inicio.Month
        XLSheet.Range("O8").Value = fecha_inicio.Year

        Dim fecha_final As Date = txtfin.Text
        XLSheet.Range("V8").Value = fecha_final.Day
        XLSheet.Range("X8").Value = fecha_final.Month
        XLSheet.Range("Z8").Value = fecha_final.Year

        XLSheet.Range("AG8").Value = txtmatriculados.Text

        XLSheet.Range("AQ8").Value = txtactivos.Text

        XLSheet.Range("L10").Value = txtnivel.Text & " " & txtcurso.Text

        XLSheet.Range("AQ10").Value = txtcodprograma.Text & " -V" & txtversion.Text
        XLSheet.Range("H12").Value = txtmunicipio.Text
        XLSheet.Range("Z12").Value = txtlugar.Text
        XLSheet.Range("N17").Value = txtinstructor.Text
        XLSheet.Range("G19").Value = txtcodproyecto.Text & "- " & txtproyecto.Text
        'XLSheet.Range("A23:A24").EntireRow.Copy()
        ' XLSheet.Range("B25").EntireRow.Insert()

        Try

            Dim contador As Integer = 23
            Dim NUMERO_COMPETENCIA As Integer = 1

            Dim fila As Integer


            sql = "Select * from  programacion where ficha='" & txtficha.Text & "'"
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
            'XLBook.Visible = True
        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        XLApp = CreateObject("Excel.Application")
        XLBook = XLApp.Workbooks.Open(My.Computer.FileSystem.CurrentDirectory & "\FICHA2.XLS")
        XLSheet = XLBook.Worksheets(1)
        XLSheet.Name = txtficha.Text
        XLApp.Visible = True
        Solocrealibro()

    End Sub

    '*************************************Matar excel del sistema
    Sub mataexcel()
        XLApp.Quit()
        XLApp = Nothing
        Try

            Dim proc As System.Diagnostics.Process

            For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
                If proc.MainWindowTitle.Trim.Length = 0 Then
                    'proc.GetCurrentProcess.StartInfo
                    proc.Kill()
                End If
            Next
        Catch ex As Exception
            My.Computer.FileSystem.WriteAllText("C:\errores.log", Format(Now, "01/MM/yyy HH:mm") & " - " & ex.Message & vbCrLf, True)
        End Try
    End Sub
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click

        mataexcel()


    End Sub
    '*************************************Borrar competencia de programacion
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles btneliminarcompe_programacion.Click
        sql = "delete from  programacion where id='" & lblidcompetencia.Text & "'"
        conectado()
        cmd = New SqlCommand(sql, cnn)
        cmd.ExecuteNonQuery()
        cerrar_conexion()


        sql = "Select * from  programacion where ficha=" & txtficha.Text
        conectado()
        datagrid = "programacion"
        llenagrid()
        btneliminarcompetencia.Enabled = False
        btneliminarcompe_programacion.Enabled = False

    End Sub

    '*************************************Agregar curso

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If txtficha.Text = "" Or txtcurso.Text = "" Or txtcodprograma.Text = "" Or txtversion.Text = "" Or txtmatriculados.Text = "" Or txtactivos.Text = "" Or txtnivel.Text = "" Or txtinicio.Text = "" Or txtfin.Text = "" Or txtmunicipio.Text = "" Or txtlugar.Text = "" Then
            MsgBox("Todos los campos son obligatorios para registrar el pprograma de formacion", MsgBoxStyle.Critical)
        Else

            Try
                sql = "select * from grupos where ficha= '" & txtficha.Text & "'"
                conectado()
                cmd = New SqlCommand(sql, cnn)

                reader = cmd.ExecuteReader
                If reader.Read Then
                    MsgBox("El Programa de formacion con esta ficha Ya existe por favor verifique ", MsgBoxStyle.Critical)
                    cerrar_conexion()
                    GoTo salir
                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            sql = "insert into grupos (ficha, codigo_programa, Version, Nombre_curso, Nivel, Fecha_inicio, Fecha_terminacion, Aprendices_matriculados, Aprendices_activos, Lugar, Municipio, Instructor_responsable, Proyecto, codigo_proyecto) values ( " & txtficha.Text & ", '" & txtcodprograma.Text & "', '" & txtversion.Text & "', '" & txtcurso.Text & "', '" & txtnivel.Text & "', '" & txtinicio.Text & "', '" & txtfin.Text & "', " & txtmatriculados.Text & ", " & txtactivos.Text & ", '" & txtlugar.Text & "', '" & txtmunicipio.Text & "', '" & txtinstructor.Text & "', '" & txtproyecto.Text & "', '" & txtcodproyecto.Text & "')"
            conectado()
            cmd = New SqlCommand(sql, cnn)
            cmd.ExecuteNonQuery()
            cerrar_conexion()
            Dim ret As Integer


            ret = MsgBox("Programa de formacion agregado exitosamente, ¿Desea agregar las competencias de otro programa?", MsgBoxStyle.YesNo)
            txtficha.Enabled = False
            If ret = 6 Then
                GroupBox3.Visible = True
                rbtomartodo.Checked = True
                rbtodo()
            End If




        End If

salir:
        txtficha.Enabled = False
    End Sub


    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Dim respuesta As Integer
        respuesta = MsgBox("Esta seguro que eliminara el curso con ficha " & txtficha.Text, MsgBoxStyle.YesNo)
        If respuesta = 6 Then


            sql = "delete from  grupos where ficha= " & txtficha.Text
            conectado()
            cmd = New SqlCommand(sql, cnn)
            cmd.ExecuteNonQuery()
            cerrar_conexion()
            MsgBox("Curso eliminado Exitosamente")

            sql = "Select * from  grupos"
            conectado()
            datagrid = "grupos"
            llenagrid()

        End If
    End Sub

    Private Sub btneliminarcompetencia_Click(sender As Object, e As EventArgs) Handles btneliminarcompetencia.Click
        sql = "delete from  competencia where id= " & id_competencia.Text
        conectado()
        cmd = New SqlCommand(sql, cnn)
        cmd.ExecuteNonQuery()
        cerrar_conexion()
        MsgBox("Competencia eliminada Exitosamente")
        sql = "Select * from  competencia where ficha= " & txtficha.Text
        conectado()
        datagrid = "competencia"
        llenagrid()
        btneliminarcompe_programacion.Enabled = False
        btneliminarcompetencia.Enabled = False
    End Sub

    Private Sub Button17_Click_1(sender As Object, e As EventArgs) Handles Button17.Click

        XLApp = CreateObject("Excel.Application")
        XLBook = XLApp.Workbooks.Open(My.Computer.FileSystem.CurrentDirectory & "\Pdocente.XLS")
        XLSheet = XLBook.Worksheets(1)
        XLSheet.Name = "FUNCIONARIO"
        XLApp.Visible = True


        XLSheet.Range("AF5").Value = txtcedula.Text
        XLSheet.Range("H5").Value = txtnombrefuncionario.Text
        XLSheet.Range("AU5").Value = txttelefono.Text
        XLSheet.Range("BD5").Value = txttelefono2.Text
        XLSheet.Range("AU9").Value = txthorascontrato.Text
        XLSheet.Range("BA9").Value = txthorasdia.Text
        XLSheet.Range("BE9").Value = txtsalario.Text
        XLSheet.Range("L10").Value = Dtinicio.Text
        XLSheet.Range("AF10").Value = Dtfin.Text
        XLSheet.Range("AU10").Value = TXTDIRECCION.Text


        XLSheet.Range("F7").Value = txtespecialidad.Text
        XLSheet.Range("AF7").Value = txtemail.Text
        XLSheet.Range("L9").Value = Cbvinculacion.Text
        XLSheet.Range("AF9").Value = Cbzona.Text


        sql = "Select * from  programacion T1 JOIN grupos T2 ON T1.ficha=T2.ficha where instructor='" & txtnombrefuncionario.Text & "' and fecha_de_inicio >= '" & dtfechainicio.Value & "'"
        Clipboard.SetText(sql)
        conectado()
        cmd = New SqlCommand(sql, cnn)
        reader = cmd.ExecuteReader
        Dim contador As Integer = 15
        Dim fila As Integer
        While reader.Read
            fila = contador

            XLSheet.Range("B" & contador & ":B" & contador + 2).EntireRow.Copy()


            contador += 3
            XLSheet.Range("B" & contador).EntireRow.Insert()

            XLSheet.Range("R" & fila).Value = reader("competencia")
            XLSheet.Range("M" & fila).Value = reader("Ficha")
            XLSheet.Range("Q" & fila).Value = reader("id")
            XLSheet.Range("B" & fila).Value = reader("Curso")
            XLSheet.Range("AI" & fila).Value = reader("duracion")
            XLSheet.Range("AL" & fila).Value = reader("fecha_de_inicio")
            XLSheet.Range("BH" & fila).Value = reader("Municipio")
            XLSheet.Range("BK" & fila).Value = reader("Lugar")
            If reader("iniciada") Then
                XLSheet.Range("R" & fila).Interior.Color = Color.Green
                XLSheet.Range("M" & fila).Interior.Color = Color.Green
                XLSheet.Range("Q" & fila).Interior.Color = Color.Green
                XLSheet.Range("B" & fila).Interior.Color = Color.Green
                XLSheet.Range("AI" & fila).Interior.Color = Color.Green
                XLSheet.Range("AL" & fila).Interior.Color = Color.Green
            End If

            If reader("terminada") Then
                XLSheet.Range("R" & fila).Interior.Color = Color.Red
                XLSheet.Range("M" & fila).Interior.Color = Color.Red
                XLSheet.Range("Q" & fila).Interior.Color = Color.Red
                XLSheet.Range("B" & fila).Interior.Color = Color.Red
                XLSheet.Range("AI" & fila).Interior.Color = Color.Red
                XLSheet.Range("AL" & fila).Interior.Color = Color.Red
            End If



            XLSheet.Range("AO" & fila).Value = reader("hdomingo_iniciada")
            XLSheet.Range("AO" & fila + 1).Value = reader("hdomingo")
            XLSheet.Range("AO" & fila + 2).Value = reader("ambiente_domingo")
            XLSheet.Range("AP" & fila).Value = reader("hdomingo_terminada")



            XLSheet.Range("AQ" & fila).Value = reader("hlunes_iniciada")
            XLSheet.Range("AQ" & fila + 1).Value = reader("hlunes")
            XLSheet.Range("AQ" & fila + 2).Value = reader("ambiente_lunes")
            XLSheet.Range("AR" & fila).Value = reader("hlunes_terminada")

            XLSheet.Range("AS" & fila).Value = reader("hmartes_iniciada")
            XLSheet.Range("AS" & fila + 1).Value = reader("hmartes")
            XLSheet.Range("AS" & fila + 2).Value = reader("ambiente_martes")
            XLSheet.Range("AT" & fila).Value = reader("hmartes_terminada")

            XLSheet.Range("AU" & fila).Value = reader("hmiercoles_iniciada")
            XLSheet.Range("AU" & fila + 1).Value = reader("hmiercoles")
            XLSheet.Range("AU" & fila + 2).Value = reader("ambiente_miercoles")
            XLSheet.Range("AV" & fila).Value = reader("hmiercoles_terminada")

            XLSheet.Range("AW" & fila).Value = reader("hjueves_iniciada")
            XLSheet.Range("AW" & fila + 1).Value = reader("hjueves")
            XLSheet.Range("AW" & fila + 2).Value = reader("ambiente_jueves")
            XLSheet.Range("AX" & fila).Value = reader("hjueves_terminada")

            XLSheet.Range("AY" & fila).Value = reader("hviernes_iniciada")
            XLSheet.Range("AY" & fila + 1).Value = reader("hviernes")
            XLSheet.Range("AY" & fila + 2).Value = reader("ambiente_viernes")
            XLSheet.Range("AZ" & fila).Value = reader("hviernes_terminada")

            XLSheet.Range("BA" & fila).Value = reader("hsabado_iniciada")
            XLSheet.Range("BA" & fila + 1).Value = reader("hsabado")
            XLSheet.Range("BA" & fila + 2).Value = reader("ambiente_sabado")
            XLSheet.Range("BB" & fila).Value = reader("hsabado_terminada")

            XLSheet.Range("BD" & fila).Value = reader("hora_programada")
            XLSheet.Range("BE" & fila).Value = reader("fecha_de_terminacion")

        End While

        XLSheet.Range("B" & contador & ":B" & contador + 1).EntireRow.Delete()
        Dim i As Integer
        Dim suma As Integer = 0
        For i = 15 To contador - 3 Step 3
            suma += XLSheet.Range("AI" & i).Value
        Next
        XLSheet.Range("AI" & contador).Value = suma
        XLApp.Application.DisplayAlerts = False



        XLApp.Application.DisplayAlerts = False

        Dim carpeta As String = "C:\Academico\" & txtnombrefuncionario.Text & "\programacion"
        Dim RutaGuardado As String = "\" & Now.Day & Now.Month & Now.Year & Now.Hour & ".xls"
        Dim dir As System.IO.DirectoryInfo = New DirectoryInfo(carpeta)
        If dir.Exists Then

            XLBook.SaveAs(carpeta & RutaGuardado)
            libro_adjunto = carpeta & RutaGuardado
        Else
            dir.Create()
            XLBook.SaveAs(carpeta & RutaGuardado)
            libro_adjunto = carpeta & RutaGuardado

        End If






    End Sub

    Private Sub Button10_Click_1(sender As Object, e As EventArgs) Handles Button10.Click
        sql = "Select * from  programacion where iniciada= 1 and (estado_de_registro= 'Sin registrar' or estado_de_registro= 'Correjido')"
        conectado()
        datagrid = "programacion"
        llenagrid()
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        sql = "Select * from  programacion where iniciada= 1 and estado_de_registro= 'Error'"
        conectado()
        datagrid = "programacion"
        llenagrid()
        datagrid = "error"
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' MsgBox("Hola")
        Timer1.Enabled = False




        Try
            sql = "update  programacion set "
            sql += " iniciada= 0, terminada= 1 "
            sql += " Where fecha_de_terminacion > '01/01/1900' and fecha_de_terminacion < '" & Now.Date & "'"
            conectado()
            cmd = New SqlCommand(sql, cnn)
            cmd.ExecuteNonQuery()
            cerrar_conexion()
            Timer1.Enabled = False
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        If Now.Hour > "8" Then
            '-*********************Programacion de aviso de terminacion de competencia
            Dim fecha_mas_cinco_dias As Date

            fecha_mas_cinco_dias = Now.Date
            fecha_mas_cinco_dias = fecha_mas_cinco_dias.AddDays(6)





            sql = "Select * from  programacion where fecha_de_terminacion < '" & fecha_mas_cinco_dias & "' and Aviso_terminacion= 0"
            conectado()
            da = New SqlClient.SqlDataAdapter(sql, cnn)

            ds = New DataSet

            da.Fill(ds, "grupos")
            Dim dt As New DataTable
            dt = New DataTable
            dt = ds.Tables(0)
            cerrar_conexion()
            Timer1.Enabled = False
            Dim j As Integer
            Dim compe As String
            Dim ficha As String
            Dim curso As String
            Dim instructor As String
            Dim fecha_de_terminacion As String
            Dim id As String


            For j = 0 To dt.Rows.Count - 1

                compe = dt.Rows(j).Item("competencia")
                ficha = dt.Rows(j).Item("ficha")
                curso = dt.Rows(j).Item("curso")
                fecha_de_terminacion = dt.Rows(j).Item("fecha_de_terminacion")
                id = dt.Rows(j).Item("id")

                instructor = dt.Rows(j).Item("instructor")
                sql = "Select * from  instructores where NOMBRE_FUNCIONARIO='" & instructor & "'"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader
                If reader.Read Then
                    para = reader("Correo")
                End If

                reader.Close()
                cerrar_conexion()

                Timer1.Enabled = False
                cuerpo = "Señor " & instructor & ", la coordinacion academica del Centro Agroempresarial y Acuícola le informa que el dia: " & fecha_de_terminacion & " se termina la competencia"
                cuerpo += " " & compe & ", que le fue programada a usted en el curso " & curso & ", con ficha " & ficha & ", Recuerde que una vez termine la competencia tiene tres días máximos para evaluarla"
                cuerpo += vbCrLf
                cuerpo += vbCrLf

                cuerpo += vbCrLf
                cuerpo += vbCrLf

                cuerpo += vbCrLf & "Cordialmente:"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "    JAVIER CARRILLO PINTO"
                cuerpo += vbCrLf & "    Coordinador Academico"
                asunto = "Terminacion de competencia"
                adjunto = ""
                enviar_correosinconfirmar()

                sql = "update  programacion set "
                sql += " Aviso_terminacion= 1, Fecha_aviso= " & "'" & Now.Date & "'"
                sql += "  where id= " & id
                conectado()
                cmd = New SqlCommand(sql, cnn)
                cmd.ExecuteNonQuery()

                cerrar_conexion()
                Timer1.Enabled = False

            Next



            Try


            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            ' Timer1.Enabled = True
        End If
exit1:
    End Sub



    Private Sub NuevoCursoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NuevoCursoToolStripMenuItem.Click
        GroupBox2.Enabled = True
        txtficha.Enabled = True
        gbprogcompetencia.Visible = False
        GroupBox3.Visible = False


    End Sub



    Private Sub gbprogcompetencia_VisibleChanged(sender As Object, e As EventArgs) Handles gbprogcompetencia.VisibleChanged
        adaptar()
    End Sub

    Sub adaptar()
        If datagrid = "competencia" Then
            DataGridView1.Columns("competencia").Width = DataGridView1.Size.Width - 500
        End If

        If gbprogcompetencia.Visible = True Then
            DataGridView1.Height = Me.Height - 600
            ' MsgBox(gbprogcompetencia.Location.X.ToString & "-" & gbprogcompetencia.Location.Y + gbprogcompetencia.Height)
            DataGridView1.Location = New System.Drawing.Point(27, gbprogcompetencia.Location.Y + gbprogcompetencia.Height)

        ElseIf GroupBox3.Visible = True Then
            DataGridView1.Height = Me.Height - 400

            DataGridView1.Location = New System.Drawing.Point(27, 390)

        Else
            DataGridView1.Height = Me.Height - 400

            DataGridView1.Location = New System.Drawing.Point(27, 229)
        End If
    End Sub

    Private Sub GroupBox3_VisibleChanged(sender As Object, e As EventArgs) Handles GroupBox3.VisibleChanged

        adaptar()
    End Sub

    Private Sub CopiaMasivaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopiaMasivaToolStripMenuItem.Click
        MsgBox("Primero seleccione los cursos a los que se les asignara la competencia y deje sin seleccionar el curso origen", MsgBoxStyle.Information)
        Dim filas, codigo_programa, version, i As Integer
        filas = DataGridView1.RowCount - 1

        codigo_programa = DataGridView1.Rows(0).Cells(1).Value
        version = DataGridView1.Rows(0).Cells(2).Value

        For i = 0 To filas - 1

            If codigo_programa = DataGridView1.Rows(i).Cells(1).Value And version = DataGridView1.Rows(i).Cells(2).Value Then

            Else
                MsgBox("Para hacer copia masiva todos los programas deben tener el mismo codigo y version", MsgBoxStyle.Critical)
                GoTo salir
            End If
        Next



        Dim dtCol = New DataGridViewCheckBoxColumn()
        dtCol.Name = "Seleccionar"
        DataGridView1.Columns.Insert(0, dtCol)
        datagrid = "compia_masiva"

salir:

    End Sub



    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        DataGridView1.Columns.Remove("Seleccionar")

    End Sub




    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        If lblidcompetencia.Text = "ID" Then
            MsgBox("Debe seleccional la competencia a terminar por sistema")
        Else
            Try
                sql = "update  programacion set "
                sql += " iniciada= 0, terminada= 1, evaluado= 1, fecha_de_inicio= '" & Now.Date & "', fecha_de_terminacion= '" & Now.Date & "', "
                sql += "hlunes_iniciada= 'sist', hlunes_terminada= 'sist', hlunes= 'sist', ambiente_lunes= 'sist', "
                sql += "hmartes_iniciada= 'sist', hmartes_terminada='sist', hmartes= 'sist', ambiente_martes= 'sist', "
                sql += "hmiercoles_iniciada= 'sist', hmiercoles_terminada= 'sist', hmiercoles='sist', ambiente_miercoles= 'sist', "
                sql += "hjueves_iniciada= 'sist', hjueves_terminada= 'sist', hjueves= 'sist', ambiente_jueves= 'sist', "
                sql += "hviernes_iniciada= 'sist', hviernes_terminada= 'sist', hviernes= 'sist', ambiente_viernes='sist', "
                sql += "hsabado_iniciada= 'sist', hsabado_terminada= 'sist', hsabado= 'sist', ambiente_sabado= 'sist', "
                sql += "hdomingo_iniciada= 'sist', hdomingo_terminada= 'sist', hdomingo= 'sist', ambiente_domingo='sist', "
                sql += "fecha_programacion=  '" & Now.Date & "', programado_por= 'sist', aviso_terminacion= 1"
                sql += " Where id = " & lblidcompetencia.Text & ""
                conectado()
                lblidcompetencia.Text = "ID"
                cmd = New SqlCommand(sql, cnn)
                Clipboard.SetText(sql)

                cmd.ExecuteNonQuery()
                MsgBox("Competencia terminada exitosamente")
                ToolStripButton1.Enabled = False
                cerrar_conexion()
                limpiar_lbl()
                sql = "Select * from  programacion where ficha=" & txtficha.Text
                conectado()
                datagrid = "programacion"
                llenagrid()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub lblidcompetencia_TextChanged(sender As Object, e As EventArgs) Handles lblidcompetencia.TextChanged
        If lblidcompetencia.Text = "ID" Then
        Else
            ToolStripButton1.Enabled = True
        End If

    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Dim filas, i As Integer
        filas = DataGridView1.RowCount - 1


        For i = 0 To filas

            If DataGridView1.Rows(i).Cells(0).Value Then
                DataGridView1.Rows(i).ReadOnly = True
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Gray
            End If
        Next
        datagrid = "copia_masiva_origen"
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        DataGridView1.Columns.Remove("Seleccionar")
    End Sub

    Private Sub VERMENUToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VERMENUToolStripMenuItem.Click
        If ToolStrip1.Visible Then
            ToolStrip1.Visible = False
        Else
            ToolStrip1.Visible = True
        End If

    End Sub



    Private Sub Form1_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged

        TabControl1.Width = Me.Width - 150
        TabControl1.Height = Me.Height - 50
        DataGridView1.Width = Me.Width - 300
        DataGridView2.Width = Me.Width - 300

        adaptar()
    End Sub


    Private Sub CambiarContraseñaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CambiarContraseñaToolStripMenuItem.Click
        LoginForm2.Show()
    End Sub

    Private Sub Label47_Click(sender As Object, e As EventArgs) Handles Label47.Click

    End Sub

    Public Function IsValidEmail(ByVal email As String) As Boolean
        If email = String.Empty Then Return False
        ' Compruebo si el formato de la dirección es correcto.
        Dim re As Regex = New Regex("^[\w._%-]+@[\w.-]+\.[a-zA-Z]{2,4}$")
        Dim m As Match = re.Match(email)
        Return (m.Captures.Count <> 0)
    End Function

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        If Dtfin.Value.Date <= Dtinicio.Value.Date Then
            MsgBox("La fecha de finalizacion del contrato debe ser mayor que la de inicio")
            GoTo salir
        End If

        If Dtfin.Value.Date <= Now.Date Then
            MsgBox("Las fechas del contrato deben ser del año en curso")
            GoTo salir
        End If

        Dim bln As Boolean = IsValidEmail(txtemail.Text)
        If bln = False Then
            MessageBox.Show("Verifique Email. Formato Incorrecto", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtemail.Focus()
            GoTo salir
        End If
        Try

            sql = "select NUMERO_IDENTIFICACION_FUNCIONARIO from instructores where NUMERO_IDENTIFICACION_FUNCIONARIO='" & txtcedula.Text & "'"
            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader
            If reader.Read Then
                MessageBox.Show("Ya existe un Instructor registrado con este documento", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                reader.Close()
                GoTo salir
            Else
                reader.Close()
                sql = "insert into  instructores (NUMERO_IDENTIFICACION_FUNCIONARIO, NOMBRE_FUNCIONARIO, TIPO_VINCULACION, Zona, Correo, Telefono, Especialidad, Estado, Fecha_de_inicio, Fecha_de_finalizacion, Horas, Direccion, Telefono2, Contrato_meses, contrato_horas_dia, asignacion, Pass) values ('" & txtcedula.Text & "', '" & txtnombrefuncionario.Text & "', '" & Cbvinculacion.Text & "', '" & Cbzona.Text & "', '" & txtemail.Text & "', " & txttelefono.Text & ", '" & txtespecialidad.Text & "', '" & txtestado.Text & "', '" & Dtinicio.Value.Date & "', '" & Dtfin.Value.Date & "', " & txthorascontrato.Text & ", '" & TXTDIRECCION.Text & "', '" & txttelefono2.Text & "', '" & txtduracion.Text & "', " & txthorasdia.Text & ", " & txtsalario.Text & ", " & txtcedula.Text & ")"
                conectado() '***********************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
                cmd = New SqlCommand(sql, cnn)
                cmd.ExecuteNonQuery()
                cerrar_conexion()
                MsgBox(" Instructor Agregado con exito ")
                inicializar()
            End If
        Catch ex As Exception
            reader.Close()
            MsgBox(ex.ToString)
        End Try
salir:

    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Try
            If txthorascontrato.Text = "" Then
                txthorascontrato.Text = 0
            End If
            If txthorasdia.Text = "" Then
                txthorasdia.Text = 0
            End If
            If txtsalario.Text = "" Then
                txtsalario.Text = 0
            End If
            If txthorascontrato.Text = "" Then
                txthorascontrato.Text = 0
            End If

            sql = "update  instructores set NOMBRE_FUNCIONARIO= '" & txtnombrefuncionario.Text & "', TIPO_VINCULACION= '" & Cbvinculacion.Text & "', Zona= '" & Cbzona.Text & "', Correo= '" & txtemail.Text & "', Telefono= '" & txttelefono.Text & "', Especialidad= '" & txtespecialidad.Text & "', Estado= '" & txtestado.Text & "', Fecha_de_inicio= '" & Dtinicio.Value & "', Fecha_de_finalizacion= '" & Dtfin.Value & "', Horas= " & txthorascontrato.Text & ", Direccion= '" & TXTDIRECCION.Text & "', Telefono2= '" & txttelefono2.Text & "', Contrato_meses= '" & txtduracion.Text & "', contrato_horas_dia= " & txthorasdia.Text & ", asignacion= " & txtsalario.Text & ", Pass= '" & txtcedula.Text & "' Where NUMERO_IDENTIFICACION_FUNCIONARIO = '" & txtcedula.Text & "'"
            conectado() '***********************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
            cmd = New SqlCommand(sql, cnn)
            cmd.ExecuteNonQuery()
            cerrar_conexion()
            MsgBox(" Instructor Actualizado con exito ")
            inicializar()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If txtficha.Text = "" Or txtcurso.Text = "" Or txtcodprograma.Text = "" Or txtversion.Text = "" Or txtmatriculados.Text = "" Or txtactivos.Text = "" Or txtnivel.Text = "" Or txtinicio.Text = "" Or txtfin.Text = "" Or txtmunicipio.Text = "" Or txtlugar.Text = "" Then
            MsgBox("Todos los campos son obligatorios para actualizar el pprograma de formacion", MsgBoxStyle.Critical)
        Else
            sql = "update  grupos set "
            sql += "codigo_programa= '" & txtcodprograma.Text & "', Version= '" & txtversion.Text & "', Nombre_curso= '" & txtcurso.Text & "', Nivel= '" & txtnivel.Text & "', Fecha_inicio= '" & txtinicio.Text & "', Fecha_terminacion= '" & txtfin.Text & "', Aprendices_matriculados= " & txtmatriculados.Text & ", Aprendices_activos= " & txtactivos.Text & ", Lugar= '" & txtlugar.Text & "', Municipio= '" & txtmunicipio.Text & "', Instructor_responsable= '" & txtinstructor.Text & "', Proyecto= '" & txtproyecto.Text & "', codigo_proyecto= '" & txtcodproyecto.Text & "'  where ficha= " & txtficha.Text & ""
            conectado()
            cmd = New SqlCommand(sql, cnn)
            cmd.ExecuteNonQuery()
            cerrar_conexion()
            Dim ret As Integer


            ret = MsgBox("Programa de formacion Actualizado exitosamente")

        End If

        txtficha.Enabled = False

    End Sub


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles txtbuscacompe.TextChanged
        vista_competencia()


    End Sub


    Private Sub TextBox1_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        lectura()
        If reader.Read Then
        Else
            MsgBox("El instructor no existe")
            ComboBox4.Text = ""
            GoTo salir
        End If
        lectura()
        While reader.Read
            ComboBox4.Items.Add(reader("NOMBRE_FUNCIONARIO"))
            If TextBox1.Text = "" Then
                ComboBox4.Text = ""
            Else
                ComboBox4.SelectedIndex = 0

            End If

        End While


salir:
    End Sub
    Sub lectura()
        If TextBox1.Text = "" Then
            sql = "Select * from  instructores ORDER BY NOMBRE_FUNCIONARIO ASC"
        Else
            sql = "Select * from  instructores where NOMBRE_FUNCIONARIO LIKE '%" + TextBox1.Text + "%'"
        End If
        ComboBox4.Items.Clear()
        conectado()
        cmd = New SqlCommand(sql, cnn)
        reader = cmd.ExecuteReader
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        vista_de_curso()
    End Sub

    Sub vista_de_curso()
        If CheckBox3.Checked And CheckBox4.Checked Then
            sql = "Select top 10 * from  grupos"
            conectado()
            datagrid = "grupos"
            llenagrid()
        ElseIf CheckBox3.Checked And CheckBox4.Checked = 0 Then

            sql = "Select top 10 * from  grupos where Fecha_terminacion <= '" & Now.Date & "'"

            conectado()
            datagrid = "grupos"
            llenagrid()
        ElseIf CheckBox3.Checked = 0 And CheckBox4.Checked Then
            sql = "Select top 10 * from  grupos where Fecha_terminacion > '" & Now.Date & "'"
            conectado()
            datagrid = "grupos"
            llenagrid()
        ElseIf CheckBox3.Checked = 0 And CheckBox4.Checked = 0 Then
            sql = "Select * from  grupos where Fecha_terminacion > '" & Now.Date & "' and Fecha_terminacion < '" & Now.Date & "'"
            conectado()
            datagrid = "grupos"
            llenagrid()
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        vista_de_curso()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles chiniciada.CheckedChanged
        vista_competencia()

    End Sub
    Sub vista_competencia()
        Try

            If ComboBox6.Text = "TODOS LOS AMBIENTES" Then
                sql = "Select * from  programacion where competencia LIKE '%" + txtbuscacompe.Text + "%' and iniciada= " & chiniciada.CheckState & " and terminada= " & chterminada.CheckState & " and cesionada= " & chseccionada.CheckState & ""

            Else
                sql = "Select * from  programacion where competencia LIKE '%" + txtbuscacompe.Text + "%' and iniciada= " & chiniciada.CheckState & " and terminada= " & chterminada.CheckState & " and cesionada= " & chseccionada.CheckState & " and " & ComboBox7.Text & "= '" & ComboBox6.Text & "'"
            End If
            'MsgBox(sql)
            conectado()
            llenagridcompe()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try




    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click

    End Sub

    Private Sub chterminada_CheckedChanged(sender As Object, e As EventArgs) Handles chterminada.CheckedChanged
        vista_competencia()
    End Sub

    Private Sub chseccionada_CheckedChanged(sender As Object, e As EventArgs) Handles chseccionada.CheckedChanged
        vista_competencia()
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        vista_competencia()
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        vista_competencia()
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

    End Sub

    Private Sub TabPage5_Click(sender As Object, e As EventArgs) Handles TabPage5.Click

    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        Dim maximo As Integer
        sql = "Select * from  ambientes order by id DESC"
        conectado()
        cmd = New SqlCommand(sql, cnn)
        reader = cmd.ExecuteReader

        If reader.Read Then
            maximo = reader("ID")
        End If

        sql = "insert into  ambientes (ID, ambiente, Municipio) values(" & maximo + 1 & ", '" & TextBox12.Text & "', '" & ComboBox8.Text & "')"
        conectado()
        cmd = New SqlCommand(sql, cnn)
        cmd.ExecuteNonQuery()
        cerrar_conexion()
        limpiambiente()
        muestra_todos_los_ambientes()
        Button29.Visible = False
        Habilita()
        carga_todos_los_ambientes()
    End Sub
    Sub limpiambiente()
        TextBox2.Text = ""
        TextBox12.Text = ""
        ComboBox8.SelectedItem = ComboBox8.Items.Item(0)
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        limpiambiente()
        desabilita()
        Button29.Visible = True
    End Sub
    Sub desabilita()
        Button26.Enabled = False
        Button27.Enabled = False
        Button28.Enabled = False
    End Sub
    Sub Habilita()
        Button26.Enabled = True
        Button27.Enabled = True
        Button28.Enabled = True
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        Dim resp As Integer = MsgBox("Esta seguro de modificar la informacion del ambiente?", MsgBoxStyle.OkCancel)
        If resp = 1 Then
            Try
                sql = "update dbo ambientes set"
                sql += "ambiente= '" & TextBox12.Text & "', Municipio= '" & ComboBox8.Text & "'where ID= " & TextBox2.Text & " "
                conectado()
                cmd = New SqlCommand(sql, cnn)
                cmd.ExecuteNonQuery()
                cerrar_conexion()
            Catch ex As Exception
                MsgBox("El ambiente se ha modificado")
            End Try
            limpiambiente()
            muestra_todos_los_ambientes()
            Habilita()
        End If

    End Sub

    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        Dim FILA As Integer = DataGridView3.CurrentRow.Index.ToString
        TextBox2.Text = DataGridView3.Rows(FILA).Cells("ID").Value.ToString()
        TextBox12.Text = DataGridView3.Rows(FILA).Cells("ambiente").Value.ToString()
        ComboBox8.Text = DataGridView3.Rows(FILA).Cells("Municipio").Value.ToString()
        Label62.Text = DataGridView3.Rows(FILA).Cells("ambiente").Value.ToString()
        Button29.Visible = False
        Habilita()
    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        sql = "Select * from  ambientes where ambiente LIKE '%" + TextBox4.Text + "%'"
        grid_ambiente()

    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        sql = "Select * from  ambientes where Municipio = '" & ComboBox9.Text & "'"
        grid_ambiente()
    End Sub

    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click
        TabControl1.SelectedTab = TabControl1.TabPages.Item(0)

        ComboBox3.Text = Label62.Text

    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        TabControl1.SelectedTab = TabControl1.TabPages.Item(4)
        TextBox4.Select()

    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        sql = "SELECT * from programacion where fecha_de_registro_sofia is not null"
        conectado()



        da = New SqlClient.SqlDataAdapter(sql, cnn)
        cb = New SqlClient.SqlCommandBuilder(da)
        ds = New DataSet
        da.Fill(ds, "compe")



        reader = cmd.ExecuteReader

        While reader.Read

        End While

    End Sub


    Private Sub TabPage7_Click(sender As Object, e As EventArgs) Handles TabPage7.Click

    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If Me.TabControl1.SelectedIndex = 6 Then
            If dtplaneacion.RowCount = 0 Then


            End If

        End If

    End Sub



    Sub llenagridcompe_planeacion()


        da = New SqlClient.SqlDataAdapter(sql, cnn)
        cb = New SqlClient.SqlCommandBuilder(da)
        ds = New DataSet
        da.Fill(ds, "compe")
        dtplaneacion.DataSource = ds
        dtplaneacion.DataMember = "compe"
        'DataGridView1.Columns("Habilitado").Visible = False
        ' DataGridView1.Columns("estado votacion").Visible = False
        ' DataGridView1.Columns("voto").Visible = False

        dtplaneacion.Columns("id").Width = 30
        dtplaneacion.Columns("curso").Width = 200
        dtplaneacion.Columns("competencia").Width = 600

    End Sub



    Private Sub rbsininiciar_CheckedChanged(sender As Object, e As EventArgs) Handles rbsininiciar.CheckedChanged
        If rbsininiciar.Checked Then
            sql = "Select * from  programacion join grupos on programacion.ficha= grupos.ficha  where iniciada=0 and terminada= 0 and cesionada= 0 and grupos.Fecha_terminacion > '" & dtpterminacion.Value.Date & "' order by programacion.ficha asc"
            conectado()
            Me.llenagridcompe_planeacion()
            Button33.Enabled = True
        End If

    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked Then
            sql = "Select * from  programacion where iniciada=1 order by ficha asc"
            conectado()
            Me.llenagridcompe_planeacion()
        End If

    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged

        If RadioButton5.Checked Then
            sql = "Select * from  programacion where terminada= 1 order by ficha asc"
            conectado()
            Me.llenagridcompe_planeacion()
        End If
    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged

        If RadioButton6.Checked Then
            sql = "Select * from  programacion where cesionada=1 order by ficha asc"
            conectado()
            Me.llenagridcompe_planeacion()
        End If
    End Sub


    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        XLApp = CreateObject("Excel.Application")
        XLBook = XLApp.Workbooks.Open(My.Computer.FileSystem.CurrentDirectory & "\planeacion.XLS")
        XLSheet = XLBook.Worksheets(1)

        XLSheet.Name = "Planeacion"
        XLSheet2 = XLBook.Worksheets(2)
        XLSheet2.Name = "filtro"
        XLApp.Visible = True
        crealibro_planeacion()


    End Sub

    Dim carpeta As String
    Dim upload As File


    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click
        seguimiento_E_P.Show()
        seguimiento_E_P.lblestatus_user.Text = lbliduserLogin.Text
        seguimiento_E_P.lblnombreusuario.Text = lblusuario.Text
    End Sub


    Private Sub Button35_Click(sender As Object, e As EventArgs)
        'My.Computer.FileSystem.CreateDirectory("C:\ReportesExcel\Reportes")
        ' If Directory.Exists("C:\ReportesExcel\Reportes") Then
        Dim carpeta As String = "Katherine Diaz"

        Dim ruta As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

        Dim fi As New IO.FileInfo(ruta)
        Try

            Directory.CreateDirectory(Path.Combine(ruta, carpeta))

            If Not Directory.Exists(Path.Combine(ruta, carpeta)) Then
                Directory.CreateDirectory(Path.Combine(ruta, carpeta))



            Else

                ' MsgBox("ya existe una carpeta con este nombre")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        ' Else
        ' MsgBox("bien")
        ' End If

    End Sub
    '--------------------------------------------------------------------------*******************tab aprendiz********************----------------------------------------------------------------
    Private Sub browseButton_Click(sender As Object, e As EventArgs) Handles browseButton.Click
        OpenFileDialog1.Filter = "Excel Worksheets|*.xls"
        OpenFileDialog1.ShowDialog()


        txtexaminar.Text = OpenFileDialog1.FileName
        btnactualizar.Enabled = True

    End Sub

    Private Sub btnactualizar_Click(sender As Object, e As EventArgs) Handles btnactualizar.Click
        Try
            XLApp = CreateObject("Excel.Application")
            If txtexaminar.Text = "" Then
                Exit Sub
            End If
            XLBook = XLApp.Workbooks.Open(txtexaminar.Text)
            XLSheet = XLBook.Worksheets(1)
            'XLSheet.Name = txtficha.Text
            'XLApp.Visible = True
            Dim i As Integer = 2
            Dim fich As Integer = MsgBox("Seguro cargar la ficha: " & XLSheet.Range("C3").Value.ToString & " ?", MsgBoxStyle.YesNo)
            If fich = 7 Then
                Exit Sub
            End If
            ' sql = "select * from grupos where ficha= " & XLSheet.Range("C3").Value.ToString
            ' conectado()
            ' cmd = New SqlCommand(sql, cnn)
            ' reader = cmd.ExecuteReader
            ' If reader.Read Then
            ' Else
            '  MsgBox("La ficha *(" & XLSheet.Range("C3").Value.ToString & ")* no existe en la base de datos")
            ' Exit Sub
            ' End If

            btnactualizar.Enabled = False
atras1:
            If XLSheet.Range("A" & i).Value.ToString <> "Tipo de Documento" Then
                i += 1
                GoTo atras1
            End If

            Dim doc As String = XLSheet.Range("A" & i).Value.ToString
            While XLSheet.Range("A" & i).Value.ToString <> ""
                i = i + 1
atras:
                If XLSheet.Range("A" & i).Value = vbNullString Then
                    Exit While
                End If
                If XLSheet.Range("B" & i).Value.ToString = doc Then
                    i += 1
                    GoTo atras
                Else
                    doc = XLSheet.Range("B" & i).Value.ToString
                End If

                sql = "select * from aprendiz where documento= '" & XLSheet.Range("B" & i).Value.ToString & "'"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader
                If reader.Read Then
                    sql = "UPDATE aprendiz set [documento]='" & doc & "',"
                    sql += "[nombre]='" & XLSheet.Range("C" & i).Value.ToString & "',"
                    sql += "[apellido]='" & XLSheet.Range("D" & i).Value.ToString & "',"
                    sql += "[ficha]='" & XLSheet.Range("C3").Value.ToString & "',"
                    sql += "[estado]='" & XLSheet.Range("E" & i).Value.ToString & "',"
                    sql += "[Tipo_documento]='" & XLSheet.Range("A" & i).Value.ToString & "'"
                    sql += "where [documento]='" & XLSheet.Range("A" & i).Value.ToString & "'"
                    conectado()
                    cmd = New SqlCommand(sql, cnn)
                    cmd.ExecuteNonQuery()
                Else
                    Try
                        sql = "insert into aprendiz (documento, nombre, apellido, ficha, Estado, Tipo_documento) "
                        sql += "Values ('" & doc & "', '" & XLSheet.Range("C" & i).Value.ToString & "', '" & XLSheet.Range("D" & i).Value.ToString & "', " & XLSheet.Range("C3").Value.ToString & ", '" & XLSheet.Range("E" & i).Value.ToString & "','" & XLSheet.Range("A" & i).Value.ToString & "')"
                        conectado()
                        cmd = New SqlCommand(sql, cnn)
                        cmd.ExecuteNonQuery()

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End If
            End While
            rbfichaaprendiz.Checked = True
            txtfichaaprendiz.Text = XLSheet.Range("C3").Value.ToString

            MsgBox("Datos guardados con exito")
            ' Cierro el libro
            XLBook.Close()

            XLApp.Quit()
            XLApp = Nothing

            ' Aunque no se recomienda, obligamos a que se
            ' lleve a cabo la recolección de elementos
            ' no utilizados.
            GC.Collect()

            ' Detenemos el proceso actual hasta que finalice
            ' el método Collect
            GC.WaitForPendingFinalizers()
        Catch ex As Exception
            If ex.ToString <> "" Then
                Exit Sub

            End If


        End Try



    End Sub

    Sub mostrar_aprendices_ficha()
        sql = "select * from aprendiz "
        If txtfichaaprendiz.Text <> "" Then
            sql += "where ficha like %" & txtfichaaprendiz.Text & "%"
        End If
        llena()

    End Sub
    Sub llena()
        Try
            conectado()
            da = New SqlDataAdapter(sql, cnn)
            cb = New SqlCommandBuilder(da)
            ds = New DataSet
            da.Fill(ds, "tabla")
            dtapreendices.DataSource = ds
            dtapreendices.DataMember = "tabla"
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub rbfichaaprendiz_CheckedChanged(sender As Object, e As EventArgs) Handles rbfichaaprendiz.CheckedChanged
        If rbfichaaprendiz.Checked Then
            btnbuscarcursos.Text = "Buscar por ficha"
        End If
    End Sub

    Private Sub rbnombrecurso_CheckedChanged(sender As Object, e As EventArgs) Handles rbnombrecurso.CheckedChanged
        If rbnombrecurso.Checked Then
            btnbuscarcursos.Text = "Buscar por nombre"
        End If
    End Sub

    Private Sub btnbuscarcursos_Click(sender As Object, e As EventArgs) Handles btnbuscarcursos.Click
        If btnbuscarcursos.Text = "Buscar por ficha" Then
            mostrar_aprendices_ficha()
        Else
            sql = "select * from aprendiz "
            If txtfichaaprendiz.Text <> "" Then
                sql += "where ficha like %" & txtfichaaprendiz.Text & "%"
            End If
            llena()
        End If
    End Sub

    Private Sub txtfichaaprendiz_TextChanged(sender As Object, e As EventArgs) Handles txtfichaaprendiz.TextChanged

        If rbfichaaprendiz.Checked Then

            sql = "select CONCAT(ficha,' | ',Nombre_curso) as combo, ficha from grupos where ficha like '%" & txtfichaaprendiz.Text & "%'"
        Else
            sql = "select CONCAT(ficha,' | ',Nombre_curso) as combo , ficha from grupos where nombre_curso like '%" & txtfichaaprendiz.Text & "%'"
        End If

        Try
            llenarcombos(cmbgrupos, "combo", "ficha")
        Catch ex As Exception
            ex.ToString()
        End Try




    End Sub

    Private Sub cmbgrupos_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbgrupos.SelectedIndexChanged

    End Sub

    '  Private Sub cmbgrupos_SelectedValueChanged(sender As Object, e As EventArgs) Handles cmbgrupos.SelectedValueChanged
    '     MsgBox("Hola")
    '    MsgBox(cmbgrupos.SelectedValue)
    '   If cmbgrupos.SelectedValue.ToString <> "" Then
    '
    '       sql = "select * from aprendiz where ficha =" & cmbgrupos.SelectedValue.ToString
    '      llena()
    ' End If



    'End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles txtnombreaprendiz.TextChanged
        If cmbgrupos.SelectedValue.ToString <> "System.Data.DataRowView" Then

            sql = "select * from aprendiz where nombre like '%" & txtnombreaprendiz.Text & "%'"
            llena()
        End If

    End Sub
    Sub limpiarInstructore()
        txtcedula.Text = ""
        txtnombrefuncionario.Text = ""
        Cbvinculacion.Text = ""
        Cbzona.Text = ""
        txtemail.Text = ""
        txttelefono.Text = ""
        txtespecialidad.Text = ""
        txtestado.Text = ""
        Dtinicio.Value = Now
        Dtfin.Value = Now
        txthorascontrato.Text = ""
        TXTDIRECCION.Text = ""
        txttelefono2.Text = ""
        txtduracion.Text = ""
        txthorasdia.Text = ""
        txtsalario.Text = ""
        txtcedula.Text = ""
    End Sub
    Private Sub Button35_Click_1(sender As Object, e As EventArgs) Handles Button35.Click
        If Button35.Text = "Nuevo" Then
            limpiarInstructore()
            GroupBox7.Enabled = True
            Button35.Text = "Cancelar"
        Else
            GroupBox7.Enabled = False
            limpiarInstructore()
            Button35.Text = "Nuevo"
        End If


    End Sub
    Sub validarAmbiente()
        Dim horaF, horaI, horaCI, horaCT As DateTime
        sql = "SELECT TOP 1 [hlunes_iniciada],[hlunes_terminada],[fecha_de_inicio],[fecha_de_terminacion] FROM programacion WHERE [ambiente_lunes]='" & ComboBox3.Text & "' order by [fecha_programacion] desc"
        conectado()
        cmd = New SqlCommand(sql, cnn)
        reader = cmd.ExecuteReader
        If reader.Read Then

            horaI = reader("hlunes_iniciada")
            horaF = reader("hlunes_terminada")
            horaCI = cmbhorainicio.Text
            horaCT = cmbhorafin.Text

            If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es: Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & " en el horario: Hora Inicio: " & horaI & " Hora Fin" & horaF, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                reader.Close()
                Exit Sub
            Else
                If horaI > horaCI And horaCT > horaF Then
                    MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas." & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    reader.Close()
                    Exit Sub
                    'MessageBox.Show(horaI & " > " & horaCI & " And " & horaCT & " > " & horaF, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    GoTo seguir
                    reader.Close()
                    Exit Sub
                End If


            End If
        Else
            GoTo seguir
            reader.Close()
        End If
seguir:
    End Sub
    Sub ValidacionAmbienteConsulta()
        Try
            ' sql = "SELECT TOP 1  [hlunes_iniciada],[hlunes_terminada],[ambiente_lunes],[hmartes_iniciada],[hmartes_terminada],[ambiente_martes],[hmiercoles_iniciada],[hmiercoles_terminada],[ambiente_miercoles],[hjueves_iniciada],[hjueves_terminada]"
            ' sql += ",[ambiente_jueves],[hviernes_iniciada],[hviernes_terminada],[ambiente_viernes],[hsabado_iniciada],[hsabado_terminada],[ambiente_sabado],[hdomingo_iniciada],[hdomingo_terminada],[ambiente_domingo] FROM [programacion] "
            If Chlunes.Checked Then
                sql = "SELECT TOP 1 [curso] ,[ficha],[hlunes_iniciada],[hlunes_terminada],[fecha_de_inicio],[fecha_de_terminacion] FROM programacion WHERE [ambiente_lunes]='" & ComboBox3.Text & "' order by [fecha_programacion] desc"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                If reader.Read Then
                    horaI = reader("hlunes_iniciada")
                    horaF = reader("hlunes_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text

                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & "En dia Lunes: Hora Inicio: " & horaI & " Hora Fin" & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas del dia lunes" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                        Else
                            reader.Close()
                        End If


                    End If
                Else
                    reader.Close()
                End If
            End If

            If ChMartes.Checked Then
                sql = "SELECT TOP 1 [curso] ,[ficha],[hmartes_iniciada],[hmartes_terminada],[fecha_de_inicio],[fecha_de_terminacion]  FROM [programacion] WHERE [ambiente_martes]='" & ComboBox3.Text & "' order by [fecha_programacion] desc"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                If reader.Read Then
                    horaI = reader("hmartes_iniciada")
                    horaF = reader("hmartes_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & " En el dia Martes:" & vbNewLine & "En el siguiente Horario: Hora Inicio: " & horaI & " Hora Fin" & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas del dia martes " & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                        Else

                            reader.Close()

                        End If


                    End If
                Else

                    reader.Close()
                End If


            End If

            If ChMiercoles.Checked Then
                sql = "SELECT TOP 1 [curso] ,[ficha], [hmiercoles_iniciada],[hmiercoles_terminada],[fecha_de_inicio],[fecha_de_terminacion] FROM [programacion] WHERE [ambiente_miercoles]='" & ComboBox3.Text & "' order by [fecha_programacion] desc"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                If reader.Read Then
                    horaI = reader("hmiercoles_iniciada")
                    horaF = reader("hmiercoles_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & " En el dia Miercoles: " & vbNewLine & "En el siguiente Horario: Hora Inicio: " & horaI & " Hora Fin" & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                        Else

                            reader.Close()

                        End If
                    End If
                Else

                    reader.Close()
                End If
            End If

            If ChJueves.Checked Then
                sql = "SELECT TOP 1 [curso] ,[ficha], [hjueves_iniciada],[hjueves_terminada],[fecha_de_inicio],[fecha_de_terminacion]  FROM [programacion] WHERE [ambiente_jueves]='" & ComboBox3.Text & "' order by [fecha_programacion] desc"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                If reader.Read Then
                    horaI = reader("hjueves_iniciada")
                    horaF = reader("hjueves_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & " En el dia Jueves: " & vbNewLine & "En el siguiente Horario: Hora Inicio: " & horaI & " Hora Fin" & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                        Else

                            reader.Close()

                        End If


                    End If
                Else

                    reader.Close()
                End If
            End If

            If ChViernes.Checked Then
                sql = "SELECT TOP 1 [curso] ,[ficha], [hviernes_iniciada],[hviernes_terminada],[fecha_de_inicio],[fecha_de_terminacion] FROM [programacion] WHERE [ambiente_viernes]='" & ComboBox3.Text & "' order by [fecha_programacion] desc"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                If reader.Read Then
                    horaI = reader("hviernes_iniciada")
                    horaF = reader("hviernes_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & " En el dia Viernes: " & vbNewLine & "En el siguiente Horario: Hora Inicio: " & horaI & " Hora Fin" & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                        Else

                            reader.Close()

                        End If


                    End If
                Else

                    reader.Close()
                End If
            End If

            If ChSabado.Checked Then
                sql = "SELECT TOP 1 [curso] ,[ficha], [hsabado_iniciada],[hsabado_terminada],[fecha_de_inicio],[fecha_de_terminacion]FROM [programacion] WHERE [ambiente_sabado]='" & ComboBox3.Text & "' order by [fecha_programacion] desc"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                If reader.Read Then
                    horaI = reader("hsabado_iniciada")
                    horaF = reader("hsabado_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & " En el dia Sabado:" & vbNewLine & "En el siguiente Horario: Hora Inicio: " & horaI & " Hora Fin" & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                        Else

                            reader.Close()

                        End If


                    End If
                Else

                    reader.Close()
                End If

            End If

            If ChDomingo.Checked Then
                sql = "SELECT TOP 1 [curso] ,[ficha],[hdomingo_iniciada],[hdomingo_terminada],[fecha_de_inicio],[fecha_de_terminacion]  FROM [programacion] WHERE [ambiente_domingo]='" & ComboBox3.Text & "' order by [fecha_programacion] desc"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                If reader.Read Then
                    horaI = reader("hdomingo_iniciada")
                    horaF = reader("hdomingo_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & "El Dia Domingo en el Siguiente horario:" & vbNewLine & "Hora Inicio: " & horaI & " Hora Fin" & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                        Else

                            reader.Close()
                            Exit Sub
                        End If


                    End If


                Else

                    reader.Close()
                End If
            End If

        Catch ex As Exception
            reader.Close()
            MsgBox(ex.ToString)
        End Try
seguir:

    End Sub


    '******************************************************************************comprueba ambiente*********************************************************************
    Function comprueba_ambiente()
        Try

            ' sql = "SELECT TOP 1  [hlunes_iniciada],[hlunes_terminada],[ambiente_lunes],[hmartes_iniciada],[hmartes_terminada],[ambiente_martes],[hmiercoles_iniciada],[hmiercoles_terminada],[ambiente_miercoles],[hjueves_iniciada],[hjueves_terminada]"
            ' sql += ",[ambiente_jueves],[hviernes_iniciada],[hviernes_terminada],[ambiente_viernes],[hsabado_iniciada],[hsabado_terminada],[ambiente_sabado],[hdomingo_iniciada],[hdomingo_terminada],[ambiente_domingo] FROM [programacion] "
            If lblAmbienteLunes.Text <> "" Then
                sql = "SELECT  [curso] ,[ficha],[hlunes_iniciada],[hlunes_terminada],[fecha_de_inicio],[fecha_de_terminacion] FROM programacion WHERE [ambiente_lunes]='" & lblAmbienteLunes.Text & "' and (('" & Dtpfechadeinicio.Value.Date & "' > fecha_de_inicio and '" & Dtpfechadeinicio.Value.Date & "'  < fecha_de_terminacion) or ('" & Dtpfechafin.Value.Date & "' > fecha_de_inicio and '" & Dtpfechafin.Value.Date & "'  < fecha_de_terminacion) or  ( fecha_de_inicio > '" & Dtpfechadeinicio.Value.Date & "'  and fecha_de_terminacion < '" & Dtpfechafin.Value.Date & "' ) ) order by [fecha_programacion] desc"
                Clipboard.SetText(sql)
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                While reader.Read
                    horaI = reader("hlunes_iniciada")
                    horaF = reader("hlunes_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text

                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde: " & reader("fecha_de_inicio") & " || Hasta: " & reader("fecha_de_terminacion") & vbNewLine & "En dia Lunes" & vbNewLine & "En el siguiente Horario: Hora Inicio: " & horaI & " - Hora Fin" & horaF & vbNewLine & "Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                        Return 1
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas del dia lunes" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & "Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                            Return 1
                            GoTo salir
                        Else



                        End If


                    End If

                End While
            End If
            If lblAmbienteMartes.Text <> "" Then
                sql = "SELECT  [curso] ,[ficha],[hmartes_iniciada],[hmartes_terminada],[fecha_de_inicio],[fecha_de_terminacion]  FROM [programacion] WHERE [ambiente_martes]='" & lblAmbienteMartes.Text & "' and '" & Dtpfechadeinicio.Value.Date & "' > fecha_de_inicio and '" & Dtpfechadeinicio.Value.Date & "'  < fecha_de_terminacion  order by [fecha_programacion] desc"
                Clipboard.SetText(sql)
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                While reader.Read
                    horaI = reader("hmartes_iniciada")
                    horaF = reader("hmartes_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & "En el dia Martes" & vbNewLine & "En el siguiente Horario: Hora Inicio: " & horaI & " - Hora Fin" & horaF & vbNewLine & "Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                        Return 1
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas del dia martes " & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & "Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                            Return 1
                            GoTo salir
                        Else

                        End If


                    End If


                End While


            End If

            If lblAmbienteMiercoles.Text <> "" Then
                sql = "SELECT TOP 1 [curso] ,[ficha], [hmiercoles_iniciada],[hmiercoles_terminada],[fecha_de_inicio],[fecha_de_terminacion] FROM [programacion] WHERE [ambiente_miercoles]='" & lblAmbienteMiercoles.Text & "'  and '" & Dtpfechadeinicio.Value.Date & "' > fecha_de_inicio and '" & Dtpfechadeinicio.Value.Date & "'  < fecha_de_terminacion order by [fecha_programacion] desc"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                While reader.Read
                    horaI = reader("hmiercoles_iniciada")
                    horaF = reader("hmiercoles_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & "En el dia Miercoles " & vbNewLine & "En el siguiente Horario: Hora Inicio: " & horaI & " - Hora Fin" & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                        Return 1
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & "Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                            Return 1
                            GoTo salir
                        Else

                        End If
                    End If
                End While
            End If

            If lblAmbienteJueves.Text <> "" Then
                sql = "SELECT TOP 1 [curso] ,[ficha], [hjueves_iniciada],[hjueves_terminada],[fecha_de_inicio],[fecha_de_terminacion]  FROM [programacion] WHERE [ambiente_jueves]='" & lblAmbienteJueves.Text & "' and '" & Dtpfechadeinicio.Value.Date & "' > fecha_de_inicio and '" & Dtpfechadeinicio.Value.Date & "'  < fecha_de_terminacion  order by [fecha_programacion] desc"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                While reader.Read
                    horaI = reader("hjueves_iniciada")
                    horaF = reader("hjueves_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & " En el dia Jueves: " & vbNewLine & "En el siguiente Horario: Hora Inicio: " & horaI & " - Hora Fin" & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                        Return 1
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & "Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                            Return 1
                            GoTo salir
                        Else

                        End If


                    End If
                End While

            End If

            If lblAmbienteViernes.Text <> "" Then
                sql = "SELECT TOP 1 [curso] ,[ficha], [hviernes_iniciada],[hviernes_terminada],[fecha_de_inicio],[fecha_de_terminacion] FROM [programacion] WHERE [ambiente_viernes]='" & lblAmbienteViernes.Text & "' and '" & Dtpfechadeinicio.Value.Date & "' > fecha_de_inicio and '" & Dtpfechadeinicio.Value.Date & "'  < fecha_de_terminacion  order by [fecha_programacion] desc"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                While reader.Read
                    horaI = reader("hviernes_iniciada")
                    horaF = reader("hviernes_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & "El dia Viernes: " & vbNewLine & "En el siguiente Horario: Hora Inicio: " & horaI & " Hora Fin" & horaF & vbNewLine & "Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                        Return 1
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & "Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                            Return 1
                            GoTo salir
                        Else

                        End If


                    End If
                End While
            End If

            If lblAmbienteSabado.Text <> "" Then
                sql = "SELECT TOP 1 [curso] ,[ficha], [hsabado_iniciada],[hsabado_terminada],[fecha_de_inicio],[fecha_de_terminacion]FROM [programacion] WHERE [ambiente_sabado]='" & lblAmbienteSabado.Text & "' and '" & Dtpfechadeinicio.Value.Date & "' > fecha_de_inicio and '" & Dtpfechadeinicio.Value.Date & "'  < fecha_de_terminacion  order by [fecha_programacion] desc"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                While reader.Read
                    horaI = reader("hsabado_iniciada")
                    horaF = reader("hsabado_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & " En el dia Sabado:" & vbNewLine & "En el siguiente Horario: Hora Inicio: " & horaI & " Hora Fin" & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                        Return 1
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                            Return 1
                            GoTo salir
                        Else



                        End If


                    End If
                End While

            End If

            If lblAmbienteDomingo.Text <> "" Then
                sql = "SELECT TOP 1 [curso] ,[ficha],[hdomingo_iniciada],[hdomingo_terminada],[fecha_de_inicio],[fecha_de_terminacion]  FROM [programacion] WHERE [ambiente_domingo]='" & lblAmbienteDomingo.Text & "' and '" & Dtpfechadeinicio.Value.Date & "' > fecha_de_inicio and '" & Dtpfechadeinicio.Value.Date & "'  < fecha_de_terminacion  order by [fecha_programacion] desc"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                While reader.Read
                    horaI = reader("hdomingo_iniciada")
                    horaF = reader("hdomingo_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & "El Dia Domingo en el Siguiente horario:" & vbNewLine & "Hora Inicio: " & horaI & " Hora Fin" & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                        Return 1
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                            Return 1
                            GoTo salir
                        Else



                        End If


                    End If


                End While
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
            reader.Close()
            Return 1

            MsgBox(ex.ToString)
        End Try
salir:
        reader.Close()
    End Function
    '*********************************************************************************************************************************************************************



    Function cruce_ambiente()
        Try

            ' sql = "SELECT TOP 1  [hlunes_iniciada],[hlunes_terminada],[ambiente_lunes],[hmartes_iniciada],[hmartes_terminada],[ambiente_martes],[hmiercoles_iniciada],[hmiercoles_terminada],[ambiente_miercoles],[hjueves_iniciada],[hjueves_terminada]"
            ' sql += ",[ambiente_jueves],[hviernes_iniciada],[hviernes_terminada],[ambiente_viernes],[hsabado_iniciada],[hsabado_terminada],[ambiente_sabado],[hdomingo_iniciada],[hdomingo_terminada],[ambiente_domingo] FROM [programacion] "
            If Chlunes.Checked Then
                sql = "SELECT  [curso] ,[ficha],[hlunes_iniciada],[hlunes_terminada],[fecha_de_inicio],[fecha_de_terminacion] FROM programacion WHERE [ambiente_lunes]='" & ComboBox3.Text & "' and '" & Dtpfechadeinicio.Value.Date & "' > fecha_de_inicio and '" & Dtpfechadeinicio.Value.Date & "'  < fecha_de_terminacion order by [fecha_programacion] desc"
                Clipboard.SetText(sql)
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                While reader.Read
                    horaI = reader("hlunes_iniciada")
                    horaF = reader("hlunes_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text

                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde: " & reader("fecha_de_inicio") & " || Hasta: " & reader("fecha_de_terminacion") & vbNewLine & "En dia Lunes" & vbNewLine & "En el siguiente Horario: Hora Inicio: " & horaI & " - Hora Fin" & horaF & vbNewLine & "Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                        Return 1
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas del dia lunes" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & "Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                            Return 1
                            GoTo salir
                        Else



                        End If


                    End If

                End While
            End If
            If ChMartes.Checked Then
                sql = "SELECT  [curso] ,[ficha],[hmartes_iniciada],[hmartes_terminada],[fecha_de_inicio],[fecha_de_terminacion]  FROM [programacion] WHERE [ambiente_martes]='" & ComboBox3.Text & "' and '" & Dtpfechadeinicio.Value.Date & "' > fecha_de_inicio and '" & Dtpfechadeinicio.Value.Date & "'  < fecha_de_terminacion  order by [fecha_programacion] desc"
                Clipboard.SetText(sql)
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                While reader.Read
                    horaI = reader("hmartes_iniciada")
                    horaF = reader("hmartes_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & "En el dia Martes" & vbNewLine & "En el siguiente Horario: Hora Inicio: " & horaI & " - Hora Fin" & horaF & vbNewLine & "Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                        Return 1
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas del dia martes " & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & "Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                            Return 1
                            GoTo salir
                        Else

                        End If


                    End If


                End While


            End If

            If ChMiercoles.Checked Then
                sql = "SELECT TOP 1 [curso] ,[ficha], [hmiercoles_iniciada],[hmiercoles_terminada],[fecha_de_inicio],[fecha_de_terminacion] FROM [programacion] WHERE [ambiente_miercoles]='" & ComboBox3.Text & "'  and '" & Dtpfechadeinicio.Value.Date & "' > fecha_de_inicio and '" & Dtpfechadeinicio.Value.Date & "'  < fecha_de_terminacion order by [fecha_programacion] desc"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                While reader.Read
                    horaI = reader("hmiercoles_iniciada")
                    horaF = reader("hmiercoles_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & "En el dia Miercoles " & vbNewLine & "En el siguiente Horario: Hora Inicio: " & horaI & " - Hora Fin" & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                        Return 1
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & "Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                            Return 1
                            GoTo salir
                        Else

                        End If
                    End If
                End While
            End If

            If ChJueves.Checked Then
                sql = "SELECT TOP 1 [curso] ,[ficha], [hjueves_iniciada],[hjueves_terminada],[fecha_de_inicio],[fecha_de_terminacion]  FROM [programacion] WHERE [ambiente_jueves]='" & ComboBox3.Text & "' and '" & Dtpfechadeinicio.Value.Date & "' > fecha_de_inicio and '" & Dtpfechadeinicio.Value.Date & "'  < fecha_de_terminacion  order by [fecha_programacion] desc"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                While reader.Read
                    horaI = reader("hjueves_iniciada")
                    horaF = reader("hjueves_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & " En el dia Jueves: " & vbNewLine & "En el siguiente Horario: Hora Inicio: " & horaI & " - Hora Fin" & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                        Return 1
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & "Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                            Return 1
                            GoTo salir
                        Else

                        End If


                    End If
                End While

            End If

            If ChViernes.Checked Then
                sql = "SELECT TOP 1 [curso] ,[ficha], [hviernes_iniciada],[hviernes_terminada],[fecha_de_inicio],[fecha_de_terminacion] FROM [programacion] WHERE [ambiente_viernes]='" & ComboBox3.Text & "' and '" & Dtpfechadeinicio.Value.Date & "' > fecha_de_inicio and '" & Dtpfechadeinicio.Value.Date & "'  < fecha_de_terminacion  order by [fecha_programacion] desc"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                While reader.Read
                    horaI = reader("hviernes_iniciada")
                    horaF = reader("hviernes_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & "El dia Viernes: " & vbNewLine & "En el siguiente Horario: Hora Inicio: " & horaI & " Hora Fin" & horaF & vbNewLine & "Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                        Return 1
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & "Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                            Return 1
                            GoTo salir
                        Else

                        End If


                    End If
                End While
            End If

            If ChSabado.Checked Then
                sql = "SELECT TOP 1 [curso] ,[ficha], [hsabado_iniciada],[hsabado_terminada],[fecha_de_inicio],[fecha_de_terminacion]FROM [programacion] WHERE [ambiente_sabado]='" & ComboBox3.Text & "' and '" & Dtpfechadeinicio.Value.Date & "' > fecha_de_inicio and '" & Dtpfechadeinicio.Value.Date & "'  < fecha_de_terminacion  order by [fecha_programacion] desc"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                While reader.Read
                    horaI = reader("hsabado_iniciada")
                    horaF = reader("hsabado_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & " En el dia Sabado:" & vbNewLine & "En el siguiente Horario: Hora Inicio: " & horaI & " Hora Fin" & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                        Return 1
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                            Return 1
                            GoTo salir
                        Else



                        End If


                    End If
                End While

            End If

            If ChDomingo.Checked Then
                sql = "SELECT TOP 1 [curso] ,[ficha],[hdomingo_iniciada],[hdomingo_terminada],[fecha_de_inicio],[fecha_de_terminacion]  FROM [programacion] WHERE [ambiente_domingo]='" & ComboBox3.Text & "' and '" & Dtpfechadeinicio.Value.Date & "' > fecha_de_inicio and '" & Dtpfechadeinicio.Value.Date & "'  < fecha_de_terminacion  order by [fecha_programacion] desc"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                While reader.Read
                    horaI = reader("hdomingo_iniciada")
                    horaF = reader("hdomingo_terminada")
                    horaCI = cmbhorainicio.Text
                    horaCT = cmbhorafin.Text
                    If horaCI < horaF And horaCI > horaI And horaF < horaCT And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCT > horaI And horaCT < horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Or horaCI = horaI And horaCT = horaF And Dtpfechadeinicio.Value.Date < reader("fecha_de_terminacion") Then
                        MessageBox.Show("El ambiente  no está disponible para dar formacion en este Horario, el horario actual es:" & vbNewLine & "Desde " & reader("fecha_de_inicio") & " || Hasta " & reader("fecha_de_terminacion") & vbNewLine & "El Dia Domingo en el Siguiente horario:" & vbNewLine & "Hora Inicio: " & horaI & " Hora Fin" & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                        Return 1
                    Else
                        If horaI > horaCI And horaCT > horaF Then
                            MessageBox.Show("Las horas a programar no cumplen con los criterios, se está ejecutando formación en la mitad de estas horas" & vbNewLine & "La hora Inicial es: " & horaI & vbNewLine & "La hora final es: " & horaF & vbNewLine & " Con la ficha: " & reader("ficha") & vbNewLine & "Del programa" & reader("curso"), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            reader.Close()
                            Return 1
                            GoTo salir
                        Else



                        End If


                    End If


                End While
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
            reader.Close()
            Return 1

            MsgBox(ex.ToString)
        End Try
salir:
        reader.Close()
    End Function

    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button37_Click(sender As Object, e As EventArgs)
        GraficaAnmbiente.Show()
    End Sub


    Private Sub DiagramarAmbienteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DiagramarAmbienteToolStripMenuItem.Click
        GraficaAnmbiente.Show()
    End Sub

    Private Sub btnsentxls_Click(sender As Object, e As EventArgs) Handles btnsentxls.Click


        cuerpo = " <p><strong>Se&ntilde;or(a):</strong> <br />" & ComboBox4.Text & ".</p> <p> Se le ha enviado una copia de de la programaci&oacute;n del Programa de formaci&oacute;n " & txtcurso.Text & ", Con ficha: " & txtficha.Text
        cuerpo += " </p>"
        cuerpo += " <p></p>"
        cuerpo += " <p></p>"

        cuerpo += " <p> Por favor verifique que la informaci&oacute;n sea correcta, de lo contrario acercarse en el menor tiempo posible a la coordinaci&oacute;n acad&eacute;mica para que sea corregida."

        cuerpo += "</p>"
        cuerpo += " <p></p>"
        cuerpo += " <p></p>"

        cuerpo += " <p> Cordialmente: </p>"

        cuerpo += " <p>    " & lblusuario.Text
        cuerpo += " <br />    Coordinador Academico"
        cuerpo += "</p>"
        If ComboBox4.Text <> "" Then
            enviaficha()
        End If
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click

    End Sub

    Private Sub RegistrarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RegistrarToolStripMenuItem.Click
        Form3.lblusuariosofia.Text = lblusuario.Text
        Form3.lbluser.Text = lbliduserLogin.Text
        Form3.Show()

    End Sub

  
    Private Sub btnfraccionar_Click(sender As Object, e As EventArgs) Handles btnfraccionar.Click
        If txthoras_programar.Text <> "" Then
            Dim horas As Integer
            Dim seccion As Integer = txthoras_programar.Text
            horas = InputBox("Cantidad de horas a fraccionar", "Fraccionar Competencia", 0)

            If horas > seccion Then
                MsgBox("La fraccion no puede ser mayor que las horas de la competencia")
                Exit Sub
            End If


            If horas > 0 Then
                seccion = seccion - horas
                sql = "update programacion set cesionada= 1, duracion= " & horas & " Where id = " & lblidcompetencia.Text & ""
                Try
                    conectado()
                    ' MsgBox(sql)
                    cmd = New SqlCommand(sql, cnn)
                    cmd.ExecuteNonQuery()
                    cerrar_conexion()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
                maximo_id_programacion()
                sql = "insert into  programacion (id, ficha, competencia, iniciada, terminada, cesionada, duracion, curso, estado_de_registro, Aviso_terminacion) values ( " & maximo + 1 & ", '" & txtficha.Text & "', '" & txtcompetencia_programacion.Text & "', 0, 0, 0, " & seccion & ", '" & txtcurso.Text & "', 'Sin registrar', 0)"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                cmd.ExecuteNonQuery()
                cerrar_conexion()
                MsgBox("competencia fraxionada Exitosamente")
                sql = "Select * from  programacion where ficha=" & txtficha.Text
                conectado()
                datagrid = "programacion"
                llenagrid()
            End If
        End If

    End Sub

    Private Sub Button37_Click_1(sender As Object, e As EventArgs)
        sql = "Select id, competencia, ficha "
    End Sub

   
    Private Sub Button37_Click_2(sender As Object, e As EventArgs) Handles Button37.Click

        Dim rpt As Integer = MessageBox.Show("¿segur que está que esta evaluada la competencia " & txtcompetencia_programacion.Text & "?", "Advertencia", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
        If rpt = 6 Then
            sql = " UPDATE programacion set Evaluado = 1 Where id = " & lblidcompetencia.Text & ""

            conectado()
            cmd = New SqlCommand(sql, cnn)
            cmd.ExecuteNonQuery()
            cerrar_conexion()
            MsgBox("competencia evaluada Exitosamente")
            sql = "Select * from  programacion where ficha=" & txtficha.Text
            conectado()
            datagrid = "programacion"
            llenagrid()

        Else
            Exit Sub
        End If



      
    End Sub

    Private Sub RegistroHorariosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RegistroHorariosToolStripMenuItem.Click

        Dim fecha As Date = InputBox("Fecha de registro", "FRegistro", Now.Date.Date)

        Dim text_dia As String
        text_dia = diasemana_minuscula(Weekday(fecha, FirstDayOfWeek.Monday))

        XLApp = CreateObject("Excel.Application")
        XLBook = XLApp.Workbooks.Open(My.Computer.FileSystem.CurrentDirectory & "\registrohorariopordia.xlsx")
        XLSheet = XLBook.Worksheets(1)
        XLSheet.Name = "CUMPLIMIENTO"
        XLApp.Visible = True
        XLSheet.Range("B3").Value = fecha.Year & "-" & fecha.Month & "-" & fecha.Day
        XLSheet.Range("B4").Value = "Fonseca"

        sql = "select t1.id, t1.curso, t1.ficha, t1.competencia, t3.NUMERO_IDENTIFICACION_FUNCIONARIO, t1.instructor, t1.ambiente_" & text_dia & ", t1.h" & text_dia & "_iniciada, t1.h" & text_dia & "_terminada "
        sql += "from Academicsoft.dbo.programacion t1 join Academicsoft.dbo.ambientes t2  "
        sql += "on t1.ambiente_" & text_dia & "=t2.ambiente "
        sql += "join Academicsoft.dbo.instructores t3 "
        sql += "on t1.instructor = t3.NOMBRE_FUNCIONARIO "
        sql += "where t2.Municipio = 'Fonseca' and h" & text_dia & "_iniciada is not null and '" & XLSheet.Range("B3").Value & "' BETWEEN t1.fecha_de_inicio AND fecha_de_terminacion "
        sql += "order by h" & text_dia & "_iniciada asc"

        Dim contador As Integer = 7
        Clipboard.SetText(sql)
        Try
            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader
            contador = 7
            Dim fila As Integer
            While reader.Read
                fila = contador
                XLSheet.Range("A" & contador & ":A" & contador).EntireRow.Copy()
                contador += 1
                XLSheet.Range("A" & contador).EntireRow.Insert()
                XLSheet.Range("A" & fila).Value = reader("ficha")
                XLSheet.Range("B" & fila).Value = reader("curso")
                XLSheet.Range("D" & fila).Value = reader("ambiente_" & text_dia)
                XLSheet.Range("E" & fila).Value = reader("h" & text_dia & "_iniciada") & ":00"
                XLSheet.Range("F" & fila).Value = reader("h" & text_dia & "_terminada") & ":00"
                XLSheet.Range("G" & fila).Value = reader("NUMERO_IDENTIFICACION_FUNCIONARIO")
                XLSheet.Range("H" & fila).Value = reader("instructor")
            End While
            cerrar_conexion()
            reader.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


        contador = 7
        FILA = contador
        'MsgBox(XLSheet.Range("A" & FILA).Value)
        While XLSheet.Cells(FILA, 1).value > 1
            FILA = contador
            sql = "SELECT[fecha],convert(char(8), hora, 108) as hora ,[usuario]  ,[Accion] from registro where usuario= '" & XLSheet.Range("G" & FILA).Value & "' and fecha = '" & XLSheet.Range("B3").Value & "'"
            Clipboard.SetText(sql)
            Try
                conectado1()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader
                While reader.Read

                    If reader("Accion") = "ENTRADA" Then
                        XLSheet.Range("I" & FILA).Value = XLSheet.Range("I" & FILA).Value & reader("hora") & vbCrLf
                    End If
                    If reader("Accion") = "SALIDA" Then
                        XLSheet.Range("J" & FILA).Value = XLSheet.Range("J" & FILA).Value & reader("hora") & vbCrLf
                    End If

                End While
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            contador = contador + 1
        End While


        XLApp.Application.DisplayAlerts = False

        Dim carpeta As String = "C:\ProgramacionAcademico\CUMPLIMIENTO\" & fecha.Date
        Dim RutaGuardado As String = "\" & fecha.Day & fecha.Month & fecha.Year & fecha.Hour & fecha.Minute & fecha.Second & ".xlsX"
        Dim dir As System.IO.DirectoryInfo = New DirectoryInfo(carpeta)
        If dir.Exists Then

            XLBook.SaveAs(carpeta & RutaGuardado)
            libro_adjunto = carpeta & RutaGuardado
        Else
            dir.Create()
            XLBook.SaveAs(carpeta & RutaGuardado)
            libro_adjunto = carpeta & RutaGuardado

        End If



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

    Private Sub GenerarCarpetasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenerarCarpetasToolStripMenuItem.Click
        Dim SourceFile, DestinationFile
        SourceFile = My.Computer.FileSystem.CurrentDirectory & "\AYUDA PARA CARGUE DE JUICIOS EVALUATIVOS.pdf"
       


        sql = "SELECT * from grupos where Fecha_terminacion>= '" & Now.Date.Date & "'"
        Clipboard.SetText(sql)
        Try
            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader
            While reader.Read

                Dim carpeta As String = "C:\ProgramacionAcademico\CARPETAS\" & reader("ficha") & "-" & reader("Nombre_curso")
                Dim dir As System.IO.DirectoryInfo = New DirectoryInfo(carpeta)
                If dir.Exists Then
                    DestinationFile = carpeta & "\AYUDA PARA CARGUE DE JUICIOS EVALUATIVOS.pdf"
                    FileCopy(SourceFile, DestinationFile) ' Copy source to target. 

                Else
                    dir.Create()
                    DestinationFile = carpeta & "\AYUDA PARA CARGUE DE JUICIOS EVALUATIVOS.pdf"
                    FileCopy(SourceFile, DestinationFile) ' Copy source to target. 

                End If

            End While
            MsgBox("Carpetas creadas con exito")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
End Class