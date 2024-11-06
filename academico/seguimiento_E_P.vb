Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Globalization
Public Class seguimiento_E_P
    '/*Diligenciamiento de Empresas*/'
    Private Sub btnnuevogEmpresa_Click(sender As Object, e As EventArgs) Handles btnnuevogEmpresa.Click
        If btnnuevogEmpresa.Text = "Nuevo" Then
            txtnit.Text = ""
            txtrazonS.Text = ""
            txtdireccionE.Text = ""
            txttelefonoE.Text = ""
            cmbmunicipioE.Enabled = False
            txtdireccionE.Enabled = True
            txttelefonoE.Enabled = True
            txtnit.Enabled = True
            txtrazonS.Enabled = True
            cmbmunicipioE.DataSource = Nothing
            cmbmunicipioE.Items.Clear()
            btnnuevogEmpresa.Text = "Guardar"
        Else
            Try
                If txtnit.Text = "" Or txtrazonS.Text = "" Or txtdireccionE.Text = "" Then
                    MsgBox("Existen campos vacíos")
                    Exit Sub
                Else
                    sql = "SELECT cedula_NIT from empresa  where cedula_NIT ='" & txtnit.Text & "'"
                    conectado()
                    cmd = New SqlCommand(sql, cnn)
                    reader = cmd.ExecuteReader
                    If reader.Read Then

                        MessageBox.Show("Esta empresa ya existe", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                        Exit Sub
                    Else
                        reader.Close()
                        sql = "INSERT INTO empresa "
                        sql += "VALUES('" & txtnit.Text & "','" & UCase(txtrazonS.Text) & "','" & cmbmunicipioE.SelectedValue.ToString & "','" & txtdireccionE.Text & "','" & txttelefonoE.Text & "')"
                        Agregar()
                        LlenarEmpresas()
                        LimpiarTextEmpresa()

                    End If

                End If


                cerrar_conexion()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If

        

    End Sub
    Private Sub cmbdptE_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbdptE.SelectedIndexChanged
        If cmbdptE.SelectedValue.ToString <> "System.Data.DataRowView" Then
            cmbmunicipioE.Enabled = True
            sql = "select t1.* from municipios t1 inner join departamento t2 on t1.departamento=t2.iddepartamento where t2.iddepartamento='" & cmbdptE.SelectedValue.ToString & "'"
            llenarcombos(cmbmunicipioE, "Municipio", "id")
        End If

    End Sub
    Private Sub txtempresaE_TextChanged(sender As Object, e As EventArgs) Handles txtempresaE.TextChanged
        If txtempresaE.Text = "" Then
            cmbempresaE.DataSource = Nothing
            cmbempresaE.Items.Clear()
        Else
            sql = "SELECT cedula_NIT,razonSocial FROM empresa where razonSocial LIKE '%" + txtempresaE.Text + "%'"
            llenarcombos(cmbempresaE, "razonSocial", "cedula_NIT")
            cmbempresaE.Enabled = True
        End If
        cerrar_conexion()
    End Sub
    Sub LimpiarTextEmpresa()
        txtnit.Text = ""
        txtrazonS.Text = ""
        txtdireccionE.Text = ""
        lblempresa.Text = ""
        txttelefonoE.Text = ""
        cmbmunicipioE.Enabled = False
        txtdireccionE.Enabled = False
        txttelefonoE.Enabled = False
        txtnit.Enabled = False
        txtrazonS.Enabled = False
        cmbmunicipioE.DataSource = Nothing
        cmbmunicipioE.Items.Clear()
        ' btnnuevogEmpresa.Text = "Nuevo"
    End Sub
    Sub LlenarAprendicesEmpresas()
        Dim empresa As String
        empresa = datagridEmpresa.CurrentRow.Cells(0).Value.ToString
        ' MsgBox(empresa)
        sql = "select TOP 80 CONCAT(t2.nombre,' ',t2.apellido)as aprendiz,fechaInicio,fechaFin,cargoAprendiz from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT where t3.cedula_NIT='" & empresa & "'"
        datagridEmpA.DataSource = ListarDatos()
        datagridEmpA.AllowUserToAddRows = False



        If datagridEmpA.RowCount > 0 Then
            datagridEmpA.Rows(0).Selected = True
        Else
            datagridEmpA.DataSource = DBNull.Value
        End If

        sql = "SELECT  TOP 80 [nombre],[cargoJefe] ,[telefonoJefe]   ,[empresa] ,[email],t2.[razonSocial] FROM [jefeInmediato] t1 inner join empresa t2 on t1.empresa=t2.[cedula_NIT] where t2.[cedula_NIT]='" & empresa & "'"
        datagridE.DataSource = ListarDatos()
        datagridE.Columns(3).Visible = False
        datagridE.Columns(4).Visible = False


        If datagridE.RowCount > 0 Then
            datagridE.Rows(0).Selected = True
            'MsgBox("entra")
        Else
            datagridAE.DataSource = DBNull.Value
        End If
    End Sub
    Private Sub datagridEmpresa_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagridEmpresa.CellClick
        LlenarAprendicesEmpresas()
        Try
            Dim fila As Integer
            fila = datagridEmpresa.CurrentRow.Index.ToString
            lblempresa.Text = datagridEmpresa.Rows(fila).Cells(0).Value.ToString
            txtnit.Text = datagridEmpresa.Rows(fila).Cells(0).Value.ToString
            txtrazonS.Text = datagridEmpresa.Rows(fila).Cells(1).Value.ToString
            cmbdptE.Text = datagridEmpresa.Rows(fila).Cells(6).Value.ToString
            cmbmunicipioE.Text = datagridEmpresa.Rows(fila).Cells(5).Value.ToString
            txtdireccionE.Text = datagridEmpresa.Rows(fila).Cells(3).Value.ToString
            txttelefonoE.Text = datagridEmpresa.Rows(fila).Cells(4).Value.ToString



        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub btneditarEmpresa_Click(sender As Object, e As EventArgs) Handles btneditarEmpresa.Click
        If btneditarEmpresa.Text = "Habilitar Campos" Then
            txtnit.Enabled = True
            txtrazonS.Enabled = True
            txtdireccionE.Enabled = True
            txttelefonoE.Enabled = True
            cmbmunicipioE.Enabled = True
            cmbdptE.Enabled = True
            btneditarEmpresa.Text = "Editar"
        Else
            Try
                If lblempresa.Text = "" Then
                    MsgBox("Debe seleccionar una empresa para editar")
                Else
                    If lblempresa.Text = "" Or txtnit.Text = "" Then
                        MsgBox("Existen campos obligatorios vacíos")
                    Else
                        sql = "UPDATE [empresa] set "
                        sql += "[cedula_NIT]='" & txtnit.Text & "',"
                        sql += "[razonSocial]='" & txtrazonS.Text & "',"
                        sql += "[municipio]='" & cmbmunicipioE.SelectedValue.ToString & "',"
                        sql += "[direccion]='" & txtdireccionE.Text & "',"
                        sql += "[telefono]='" & txttelefonoE.Text & "'"
                        sql += "where [cedula_NIT]='" & lblempresa.Text & "'"
                        Agregar()
                        LlenarEmpresas()
                        LimpiarTextEmpresa()

                    End If


                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If

       
    End Sub
    Private Sub cmbdepartamentoaprendiz_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbdepartamentoaprendiz.SelectedIndexChanged

        If cmbdepartamentoaprendiz.SelectedValue.ToString <> "System.Data.DataRowView" Then
            cmbmunicipioaprendiz.Enabled = True
            sql = "select t1.* from municipios t1 inner join departamento t2 on t1.departamento=t2.iddepartamento where t2.iddepartamento='" & cmbdepartamentoaprendiz.SelectedValue.ToString & "'"
            llenarcombos(cmbmunicipioaprendiz, "Municipio", "id")
        End If



    End Sub
    Private Sub btnagregarNA_Click(sender As Object, e As EventArgs) Handles btnagregarNA.Click
        Try
            If btnagregarNA.Text = "Nuevo" Then
                txtdocumento.Text = ""
                txtdocumento.Enabled = True
                txtnombreaprendiz.Text = ""
                txtnombreaprendiz.Enabled = True
                txtapellidoaprendiz.Text = ""
                txtapellidoaprendiz.Enabled = True
                txtemailaprendiz.Text = ""
                txtemailaprendiz.Enabled = True
                txttelefonoaprendiz.Text = ""
                cmbestadoformacion.Enabled = True
                txtficha.Text = ""
                txtficha.Enabled = True
                txtdireccionaprendiz.Text = ""
                txtdireccionaprendiz.Enabled = True
                txttelefonoaprendiz.Enabled = True
                txtemailaprendiz.Enabled = True
                cmbmunicipioaprendiz.Text = ""
                cmbmunicipioaprendiz.DataSource = Nothing
                cmbmunicipioaprendiz.Items.Clear()
                cmbdepartamentoaprendiz.Enabled = True
                cmbtipodoc.Enabled = True
                txtprogramaA.Text = ""
                cmbmunicipioaprendiz.Enabled = False
                btnagregarNA.Text = "Guardar"
            Else
                If txtdocumento.Text = "" Or txtnombreaprendiz.Text = "" Or
                 txtapellidoaprendiz.Text = "" Or
                 txtficha.Text = "" Or cmbdepartamentoaprendiz.Text = "" Then
                    MsgBox("Existen campos vacíos")
                Else
                    Dim bln As Boolean = IsValidEmail(txtemailaprendiz.Text)
                    If bln = False Then
                        MessageBox.Show("Verifique Email Formato Incorrecto", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        txtemailaprendiz.Focus()
                        Exit Sub
                    Else
                        sql = "SELECT documento from aprendiz  where documento ='" & txtdocumento.Text & "'"
                        conectado()
                        cmd = New SqlCommand(sql, cnn)
                        reader = cmd.ExecuteReader
                        If reader.Read Then
                            MsgBox("Este aprendiz ya existe")
                            reader.Close()
                        Else
                            reader.Close()
                            sql = "SELECT ficha, Nombre_curso from [grupos]  where ficha ='" & txtficha.Text & "'"
                            conectado()
                            cmd = New SqlCommand(sql, cnn)
                            reader = cmd.ExecuteReader
                            If reader.Read Then
                                reader.Close()
                                cerrar_conexion()
                                sql = "INSERT INTO aprendiz (documento,nombre,apellido,correo, telefono, ficha, direccion,municipio,Estado,Tipo_documento)"
                                sql += "VALUES('" & txtdocumento.Text & "', '" & UCase(txtnombreaprendiz.Text) & "', '" & UCase(txtapellidoaprendiz.Text) & "', '" & txtemailaprendiz.Text & "', '" & txttelefonoaprendiz.Text & "'," & txtficha.Text & ","
                                sql += "'" & txtdireccionaprendiz.Text & "','" & cmbmunicipioaprendiz.SelectedValue.ToString & "','" & cmbestadoformacion.Text & "','" & cmbtipodoc.Text & "')"
                                Agregar()
                                LlenarDataGridSAprendices()
                                limpiarAprendiz()

                                Agregarcolor()
                            Else
                                MessageBox.Show("la ficha agregada no existe, verifique que esté bien escrita", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                reader.Close()
                            End If
                        End If
                    End If



                End If
            End If

            





            cerrar_conexion()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
    Sub limpiarAprendiz()
        txtdocumento.Text = ""
        txtnombreaprendiz.Text = ""
        txtapellidoaprendiz.Text = ""
        txtemailaprendiz.Text = ""
        txttelefonoaprendiz.Text = ""
        txtficha.Text = ""
        txtdireccionaprendiz.Text = ""
        cmbmunicipioaprendiz.Text = ""
        txtprogramaA.Text = ""
        txtficha.Enabled = False
        txtdocumento.Enabled = False
        cmbtipodoc.Enabled = False
        cmbestadoformacion.Enabled = False
        cmbdepartamentoaprendiz.Enabled = False
        txtapellidoaprendiz.Enabled = False
        txtdireccionaprendiz.Enabled = False
        txttelefonoaprendiz.Enabled = False
        txtemailaprendiz.Enabled = False
        cmbmunicipioaprendiz.DataSource = Nothing
        cmbmunicipioaprendiz.Items.Clear()
        cmbmunicipioaprendiz.Enabled = False

    End Sub
    Private Sub btnbuscarAD_Click(sender As Object, e As EventArgs) Handles btnbuscarAD.Click
        If txtdocumento.Text = "" Then
            ' sql = "select TOP 50 [documento],[nombre],[apellido],[telefono],correo,[ficha],[Estado],t2.Municipio,t1.municipio, [Tipo_documento],direccion,t3.nombreDepartamento from aprendiz t1 inner join municipios t2 on t1.municipio=t2.Id inner join departamento t3 on t2.departamento=t3.iddepartamento"
            sql = "select TOP 80 [documento],[nombre],[apellido],[telefono],correo,[ficha],[Estado],t1.municipio, [Tipo_documento],direccion from aprendiz t1"
            datagridaprendices.DataSource = ListarDatos()
            Agregarcolor()
            txtdocumento.Enabled = True
        Else
            ' sql = "select TOP 50 [documento],[nombre],[apellido],[telefono],correo,[ficha],[Estado],t2.Municipio,t1.municipio, [Tipo_documento],direccion,t3.nombreDepartamento from aprendiz t1 inner join municipios t2 on t1.municipio=t2.Id inner join departamento t3 on t2.departamento=t3.iddepartamento where [documento] like '%" & txtdocumento.Text & "%'"
            sql = "select TOP 80 [documento],[nombre],[apellido],[telefono],correo,[ficha],[Estado],t1.municipio, [Tipo_documento],direccion from aprendiz t1  where [documento] like '%" & txtdocumento.Text & "%'"
            datagridaprendices.DataSource = ListarDatos()
            Agregarcolor()
            datagridaprendices.Columns(7).Visible = False
            ' datagridaprendices.Columns(6).Visible = False
            ' datagridaprendices.Columns(8).Visible = False
            ' datagridaprendices.Columns(9).Visible = False

        End If
        cerrar_conexion()
    End Sub
    Private Sub datagridaprendices_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagridaprendices.CellClick
        Try
            Dim fila As Integer
            fila = datagridaprendices.CurrentRow.Index.ToString
            txtdocumento.Text = datagridaprendices.Rows(fila).Cells(0).Value.ToString
            iddocumento.Text = datagridaprendices.Rows(fila).Cells(0).Value.ToString
            txtnombreaprendiz.Text = datagridaprendices.Rows(fila).Cells(1).Value.ToString
            txtapellidoaprendiz.Text = datagridaprendices.Rows(fila).Cells(2).Value.ToString
            txttelefonoaprendiz.Text = datagridaprendices.Rows(fila).Cells(3).Value.ToString
            txtemailaprendiz.Text = datagridaprendices.Rows(fila).Cells(4).Value.ToString
            txtficha.Text = datagridaprendices.Rows(fila).Cells(5).Value.ToString
            cmbestadoformacion.Text = datagridaprendices.Rows(fila).Cells(6).Value.ToString

            'If datagridaprendices.Rows(fila).Cells(7).Value.Equals(DBNull.Value) Then
            'cmbmunicipioaprendiz.Text = ""
            'cmbmunicipioaprendiz.DataSource = Nothing
            ' cmbmunicipioaprendiz.Items.Clear()
            ' Else
            ' cmbmunicipioaprendiz.Text = datagridaprendices.Rows(fila).Cells(7).Value.ToString
            ' End If

            cmbtipodoc.Text = datagridaprendices.Rows(fila).Cells(8).Value.ToString
            txtdireccionaprendiz.Text = datagridaprendices.Rows(fila).Cells(9).Value.ToString


            ' cmbdepartamentoaprendiz.Text = datagridaprendices.Rows(fila).Cells(11).Value.ToString

            sql = "select TOP 80 CONCAT(t2.nombre,' ',t2.apellido)as aprendiz, t3.razonSocial,fechaInicio,fechaFin,instructor,cargoAprendiz,t1.idseguimiento,t5.nombreModalidad,t2.documento,t6.nombreEstado from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT inner join modalidad t5 on t1.modalidad=t5.idmodalida inner join estado t6 on t1.Estado=t6.idestado"
            sql += " where t1.aprendiz= '" & txtdocumento.Text & "'"
            datagridseguimiento.DataSource = ListarDatos()
            datagridseguimiento.Columns(4).Visible = False
            datagridseguimiento.Columns(5).Visible = False
            datagridseguimiento.Columns(6).Visible = False
            datagridseguimiento.Columns(8).Visible = False

            If datagridseguimiento.RowCount > 0 Then
                datagridseguimiento.Rows(0).Selected = True
            Else
                ' MsgBox("")
            End If


            cerrar_conexion()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub txtempresaSA_TextChanged(sender As Object, e As EventArgs) Handles txtempresaSA.TextChanged
        If txtempresaSA.Text = "" Then
            cmbempresaSA.DataSource = Nothing
            cmbempresaSA.Items.Clear()
        Else
            sql = "SELECT cedula_NIT,razonSocial FROM empresa where razonSocial LIKE '%" + txtempresaSA.Text + "%'"
            llenarcombos(cmbempresaSA, "razonSocial", "cedula_NIT")


        End If
        'cerrar_conexion()
    End Sub
    Public Function IsValidEmail(ByVal email As String) As Boolean
        If email = String.Empty Then Return False
        ' Compruebo si el formato de la dirección es correcto.
        Dim re As Regex = New Regex("^[\w._%-]+@[\w.-]+\.[a-zA-Z]{2,4}$")
        Dim m As Match = re.Match(email)
        Return (m.Captures.Count <> 0)
    End Function
    Private Sub btnnuevoSeguimiennto_Click(sender As Object, e As EventArgs) Handles btnnuevoSeguimiennto.Click
        If btnnuevoSeguimiennto.Text = "Nuevo" Then
            txtinstructorSA.Text = ""
            txtinstructorSA.Enabled = True
            cmbinstructorSA.Enabled = True

            txtempresaSA.Text = ""
            txtempresaSA.Enabled = True
            cmbempresaSA.Enabled = True
            cmbestadoSA.Enabled = True
            cmbmodalidadSA.Enabled = True
            'cmbjefeISA.Enabled = True
            txtcargoA.Enabled = True
            txtcargoA.Text = ""
            ' txtjefeISA.Text = ""
            cmbmodalidadS.Enabled = True
            btnnuevoSeguimiennto.Text = "Guardar"
        Else
            Try

                If txtdocumento.Text = "" Then
                    MessageBox.Show("Debe Seleccionar un Aprendiz para hacer el seguimiento.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

                Else
                    If txtempresaSA.Text = "" Or cmbempresaSA.Text = "" Then
                        MessageBox.Show("Existen valores obligatorios vacios, favor de llenar todos los campos", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        If datetimeinicio.Value > datetimefin.Value Then
                            MessageBox.Show("La fecha de inicio de etapa practica no puede ser mayor a la fecha final", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else

                            Dim respuesta = MessageBox.Show("Está seguro que desea registrar este Seguimiento?", "Informacion", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                            If respuesta = 6 Then
                                sql = "select  TOP 1 documento,CONCAT(t2.nombre,' ',t2.apellido)as aprendiz, t3.razonSocial,fechaInicio,fechaFin,instructor,cargoAprendiz,t1.idseguimiento,t5.nombreModalidad as Modalidad,t7.nombreEstado, t6.Nombre_curso,t2.ficha,fechaAsignacion from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT  inner join modalidad t5 on t1.modalidad=t5.idmodalida join grupos t6 on t2.ficha=t6.ficha inner join estado t7 on t1.estado=t7.idestado where t2.documento='" & txtdocumento.Text & "' order by fechaAsignacion desc"
                                conectado()
                                cmd = New SqlCommand(sql, cnn)
                                reader = cmd.ExecuteReader
                                If reader.Read Then
                                    ' MsgBox(reader("fechaAsignacion"))
                                    'MsgBox(reader("fechaFin"))
                                    If reader("fechaFin").ToString <= Now.Date.ToString Or reader("fechaAsignacion") <= reader("fechaFin").ToString Then

                                        MessageBox.Show("Este aprendiz ya tiene una asignacion", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                        reader.Close()
                                        Exit Sub
                                    Else
                                        reader.Close()
                                        GoTo seguir
                                    End If

                                Else
seguir:
                                    If txtemailaprendiz.Text = "" Then
                                        MessageBox.Show("Este aprendiz no tiene correo", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                        Exit Sub
                                    Else
                                        sql = "select Estado from aprendiz   where documento ='" & txtdocumento.Text & "'"
                                        conectado()
                                        cmd = New SqlCommand(sql, cnn)
                                        reader = cmd.ExecuteReader
                                        If reader.Read Then
                                            ' MsgBox(reader("Estado").ToString)
                                            If reader("Estado").ToString = "" Or reader("Estado").ToString = DBNull.Value.ToString Or reader("Estado").ToString = "CANCELADO" Or reader("Estado").ToString = "TRASLADADO" Or reader("Estado").ToString = "CONDICIONADO" Or reader("Estado").ToString = "RETIRO VOLUNTARIO" Then
                                                MessageBox.Show("Este aprendiz no tiene un estado aceptado por el sistema", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                reader.Close()
                                                Exit Sub
                                            Else
                                                reader.Close()
                                                ' Now & " " & String.Format("{0:HH:mm:ss}", DateTime.Now)
                                                sql = "INSERT INTO seguimiento (instructor, empresa, cargoAprendiz, fechaInicio, fechaFin, usuario, modalidad, fechaAsignacion ,estado, aprendiz,hora ) "
                                                sql += "VALUES ('" & cmbinstructorSA.Text & "', '" & cmbempresaSA.SelectedValue.ToString & "','" & txtcargoA.Text & "','" & datetimeinicio.Value.ToShortDateString & "','" & datetimefin.Value.ToShortDateString & "','" & lblestatus_user.Text & "', " & cmbmodalidadSA.SelectedValue.ToString & ",'" & DateTimeAsignacion.Value.Date & "', '" & cmbestadoSA.SelectedValue.ToString & "','" & txtdocumento.Text & "','" & Now.ToString("HH:mm:ss") & "' )"
                                                Agregar()
                                                LlenarSeguimientosA()
                                                LimpiarSeguimiento()
                                                cerrar_conexion()

                                                Enviar()

                                            End If
                                        Else
                                            reader.Close()

                                        End If
                                    End If


                                End If

                            Else
                                Exit Sub
                            End If
                           
                        End If
                    End If
                End If


                cerrar_conexion()
            Catch ex As Exception
                reader.Close()
                MsgBox(ex.ToString)
            End Try
        End If

        


    End Sub
    Sub Enviar()

        Try
            reader.Close()
            Dim consulta, aprendiz, instructor As String
            consulta = "select TOP 1  Tipo_documento,documento,concat(t1.nombre,' ',apellido) as aprendiz,t1.ficha as fichaAprendiz, t3.Nivel,t2.fechaInicio,t2.fechaFin, t3.Nombre_curso as programa,t1.telefono as telefonoAprendiz,t1.correo AS CorreoAprendiz, t5.cedula_NIT,t5.razonSocial,t6.Municipio as nombreMuni,t5.direccion as direccionE,t5.telefono as telefonoE, t7.nombreDepartamento as departamentoEmpresa, t4.NUMERO_IDENTIFICACION_FUNCIONARIO,t4.NOMBRE_FUNCIONARIO,t4.Correo as CorreoInstructor,t4.Telefono as telefonoI, [idseguimiento] from aprendiz t1 inner join seguimiento t2 on t1.documento=t2.aprendiz inner join grupos t3 on t3.ficha=t1.ficha inner join instructores t4 on t2.instructor=t4.NOMBRE_FUNCIONARIO inner join empresa t5 on t2.empresa=t5.cedula_NIT inner join municipios t6 on t5.municipio=t6.Id inner join departamento t7 on t6.departamento=t7.iddepartamento where t2.aprendiz='" & txtdocumento.Text & "' ORDER BY fechaAsignacion DESC"
            conectado()
            cmd = New SqlCommand(consulta, cnn)
            reader = cmd.ExecuteReader

            If reader.Read Then
                ' MsgBox("")
                If reader("CorreoAprendiz").ToString = "" Or reader("CorreoAprendiz").ToString = DBNull.Value.ToString Then

                    instructor = reader("CorreoInstructor")
                    para = instructor '"khattherine@gmail.com" 
                Else
                    aprendiz = reader("CorreoAprendiz")
                    instructor = reader("CorreoInstructor") ' "khattherine@gmail.com"
                    para = aprendiz & ";" & instructor '"khattherine@gmail.com" 
                End If


                cuerpo = "<HTML><BODY><h3 >Se ha registrado un seguimiento de etapa productiva.</h3>" & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "<b><p style='font-family:Engravers MT;font-size:38'>Informacion del Aprendiz:</p></b><br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Aprendiz: " & reader("aprendiz") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Documento: " & reader("documento") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Telefono: " & reader("telefonoAprendiz") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Ficha: " & reader("fichaAprendiz") & "<br>"
                cuerpo += vbCrLf
                ficha_aprendiz = reader("fichaAprendiz")
                cuerpo += vbCrLf & "Nivel de programa de Formacion: " & reader("Nivel") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Programa de Formacion: " & reader("programa") & "<br>"
                cuerpo += vbCrLf & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "<b><p style='font-family:Engravers MT;font-size:30'>Informacion de la Empresa:</p></b>" & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "NIT: " & reader("cedula_NIT") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Empresa: " & reader("razonSocial") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Municipio: " & reader("nombreMuni") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Direccion: " & reader("direccionE") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Telefono: " & reader("telefonoE") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "<br>"
                cuerpo += vbCrLf & "<b><p style='font-family:Engravers MT;font-size:30'>Informacion del instructor responsable del seguimiento:</p></b>" & "<br>"
                cuerpo += vbCrLf

                cuerpo += vbCrLf & "Instructor: " & reader("NOMBRE_FUNCIONARIO") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Telefono: " & reader("telefonoI") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Correo: " & instructor & "<br>"
                cuerpo += vbCrLf

                ' cuerpo += vbCrLf & "Se ha asignado seguimiento "

                cuerpo += vbCrLf
                cuerpo += vbCrLf & "<br>"

                cuerpo += vbCrLf & "<b><p style='font-family:Calibri;font-size:30'>Cordialmente:</p></b>"

                cuerpo += vbCrLf & " <b> <p style='font-family:Arial Rounded MT Bold'>ERIKA CECILIA PITRES BOLAÑOS</p></b>"
                cuerpo += vbCrLf & "<b> <p style='font-family:Calibri '>Coordinador Academico</p></b></BODY></HTML>"

                'adjunto = "c:\cartas_seguimiento\" & ficha_aprendiz & "\" & reader("aprendiz") & ".doc"
                asunto = "Seguimiento Aprendiz"

                enviar_correoseg()
                print_carta_seguimiento()
                If reader("CorreoAprendiz").ToString = "" Or reader("CorreoAprendiz").ToString = DBNull.Value.ToString Then


                    Try
                        sql = "UPDATE seguimiento set "
                        sql += "[confirmacionEA]='0',"
                        sql += "[confirmacionEI]='1'"
                        sql += "where [idseguimiento]=" & reader("idseguimiento") & ""
                        Agregar()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                Else

                    Try
                        sql = "UPDATE seguimiento set "
                        sql += "[confirmacionEA]='1',"
                        sql += "[confirmacionEI]='1'"
                        sql += "where [idseguimiento]=" & reader("idseguimiento") & ""
                        Agregar()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End If





                reader.Close()
                cerrar_conexion()
            Else

                AvisoAsignacionInstructor(txtdocumento.Text)
                reader.Close()
                cerrar_conexion()
            End If


        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
       
    End Sub
    Sub AvisoAsignacionInstructor(aprendiz As String)
        reader.Close()
        Dim query, coordinador As String
        query = "select TOP 1  Tipo_documento,documento,concat(t1.nombre,' ',apellido) as aprendiz,t1.ficha as fichaAprendiz, t3.Nivel,t2.fechaInicio,t2.fechaFin, t3.Nombre_curso as programa,t1.telefono as telefonoAprendiz,t1.correo AS CorreoAprendiz, t5.cedula_NIT,t5.razonSocial,t6.Municipio as nombreMuni,t5.direccion as direccionE,t5.telefono as telefonoE, t7.nombreDepartamento as departamentoEmpresa,t2.instructor from aprendiz t1 inner join seguimiento t2 on t1.documento=t2.aprendiz inner join grupos t3 on t3.ficha=t1.ficha  inner join empresa t5 on t2.empresa=t5.cedula_NIT inner join municipios t6 on t5.municipio=t6.Id inner join departamento t7 on t6.departamento=t7.iddepartamento where t2.aprendiz='" & aprendiz & "' ORDER BY fechaAsignacion DESC"
        conectado()
        cmd = New SqlCommand(query, cnn)
        reader = cmd.ExecuteReader
        Dim null As String = "NULL"
        If reader.Read Then
            If reader("instructor").ToString = "" Or reader("instructor").ToString = null Or reader("instructor").ToString = DBNull.Value.ToString Then
                coordinador = "ecpitrebo@misena.edu.co" '"javiercarrillo@misena.edu.co""khattherine@gmail.com" '
                para = coordinador
                'style='color:#80BFFF'"
                cuerpo = "<HTML><BODY><h3 >Estimado Coordinador, el presente correo es para informarle que el siguiente aprendiz aún no se le ha asignado un instructor para el seguimiento de la etapa productiva, por favor diligenciar la asignacion del instructor.</h3><br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "<b><p style='font-family:Engravers MT;font-size:30'>Informacion del Aprendiz:</p></b><br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Aprendiz: " & reader("aprendiz") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Documento: " & reader("documento") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Telefono: " & reader("telefonoAprendiz") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "<Ficha: " & reader("fichaAprendiz") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Nivel de programa de Formacion: " & reader("Nivel") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Programa de Formacion: " & reader("programa")
                cuerpo += vbCrLf & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "<b><p style='font-family:Engravers MT;font-size:30'>Informacion de la Empresa:</p></b>" & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "NIT: " & reader("cedula_NIT") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Empresa: " & reader("razonSocial") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Municipio: " & reader("nombreMuni") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Direccion: " & reader("direccionE") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Telefono: " & reader("telefonoE") & "<br>"

                cuerpo += vbCrLf & "<br>"

                cuerpo += vbCrLf & "<p style='font-family:Engravers MT;font-size:30'>Cordialmente:</p>" & "<br>"

                cuerpo += vbCrLf & "<b> <p style='font-family:Arial Rounded MT Bold'>  Sistema Automatico de Seguimiento de Aprendices</p></b></BODY></HTML>"



                asunto = "Informacion de Seguimiento de Aprendices en etapa practica"
                enviar_correoseg()

                reader.Close()
            End If
            reader.Close()
        Else
            MsgBox("no se envio el correo")
            reader.Close()
        End If
        reader.Close()

        cerrar_conexion()
    End Sub
    Sub print_carta_seguimiento()
        Dim msword As New Word.Application
        Dim documento As Word.Document

        Dim carpeta As String = "c:\cartas_seguimiento\" & ficha_aprendiz



        Dim dir As System.IO.DirectoryInfo = New DirectoryInfo(carpeta)
        If dir.Exists Then
            FileCopy(My.Computer.FileSystem.CurrentDirectory & "\carta_asignacion.doc", carpeta & "\" & reader("aprendiz") & ".doc")
        Else
            dir.Create()
            FileCopy(My.Computer.FileSystem.CurrentDirectory & "\carta_asignacion.doc", carpeta & "\" & reader("aprendiz") & ".doc")
        End If



        documento = msword.Documents.Open(carpeta & "\" & reader("aprendiz") & ".doc")

        documento.Bookmarks.Item("fecha").Range.Text = Now
        documento.Bookmarks.Item("aprendiz").Range.Text = UCase(reader("aprendiz"))
        documento.Bookmarks.Item("correo_instructor").Range.Text = reader("CorreoInstructor")
        documento.Bookmarks.Item("documento").Range.Text = reader("documento")
        documento.Bookmarks.Item("empresa").Range.Text = reader("razonSocial")

        documento.Bookmarks.Item("ficha").Range.Text = reader("fichaAprendiz")
        documento.Bookmarks.Item("instructor").Range.Text = reader("NOMBRE_FUNCIONARIO")
        documento.Bookmarks.Item("nivel").Range.Text = reader("Nivel")
        documento.Bookmarks.Item("programa").Range.Text = reader("programa")
        If reader("Tipo_documento") = "CC" Then
            documento.Bookmarks.Item("tipodocumento").Range.Text = "cedula de ciudadania"
            'reader.Close()
        ElseIf reader("Tipo_documento") = "TI" Then
            documento.Bookmarks.Item("tipodocumento").Range.Text = "tarjeta de identidad"
            'reader.Close()
        End If

        ' documento.Bookmarks.Item("tiupousuario").Range.Text = DataGridView1.Rows(i).Cells("tipo_usuario").Value
        ' documento.Bookmarks.Item("estado").Range.Text = DataGridView1.Rows(i).Cells("Estado").Value
        ' msword.Visible = True
        documento.Save()
        'reader.Close()
        documento.Close()
    End Sub
    Sub EnviaMail(ruta As String)

        Dim m_Outlook As Outlook.Application
        Dim objMail As Outlook.MailItem

        Dim HTML As String
        HTML = "Texto que se mostrará en el cuerpo del correo"

        Try
            m_Outlook = New Outlook.Application
            objMail = m_Outlook.CreateItem(Outlook.OlItemType.olMailItem)

            Dim consulta, aprendiz, instructor As String
            consulta = "select TOP 1  documento,concat(t1.nombre,' ',apellido) as aprendiz,t1.ficha as fichaAprendiz,t3.Nombre_curso,t1.telefono as telefonoAprendiz,t1.correo AS CorreoAprendiz, t5.cedula_NIT,t5.razonSocial,t6.Municipio,t7.nombreDepartamento, t4.NUMERO_IDENTIFICACION_FUNCIONARIO,t4.NOMBRE_FUNCIONARIO,t4.Correo as CorreoInstructor from aprendiz t1 inner join seguimiento t2 on t1.documento=t2.aprendiz inner join grupos t3 on t3.ficha=t1.ficha inner join instructores t4 on t2.instructor=t4.NOMBRE_FUNCIONARIO inner join empresa t5 on t2.empresa=t5.cedula_NIT inner join municipios t6 on t5.municipio=t6.Id inner join departamento t7 on t6.departamento=t7.iddepartamento where t2.aprendiz='" & txtdocumento.Text & "' ORDER BY fechaAsignacion DESC"
            conectado()
            cmd = New SqlCommand(consulta, cnn)
            reader = cmd.ExecuteReader

            If reader.Read Then
                aprendiz = reader("CorreoAprendiz")
                instructor = reader("CorreoInstructor")
                objMail.To = aprendiz & ";" & instructor

                reader.Close()
            End If

            objMail.Subject = "Asunto del Correo"
            objMail.HTMLBody = HTML
            objMail.Importance = Outlook.OlImportance.olImportanceHigh
            objMail.Attachments.Add(ruta)

            Me.Cursor = Cursors.Default
            objMail.Display()

        Catch ex As Exception
            Me.Cursor = Cursors.Default

        Finally
            m_Outlook = Nothing
        End Try
    End Sub
    Sub LimpiarSeguimiento()
        txtinstructorSA.Text = ""
        txtinstructorSA.Enabled = False
        cmbinstructorSA.Enabled = False
        txtempresaSA.Text = ""
        txtempresaSA.Enabled = False
        cmbempresaSA.Enabled = False
        cmbestadoSA.Enabled = False
        cmbmodalidadSA.Enabled = False
        lblconfirmacion.Text = ""
        txtcargoA.Text = ""
        txtcargoA.Enabled = False
        btneditarSeguimiento.Text = "Habilitar Campos"



    End Sub
    Sub LlenarSeguimientoDetalleAprendiz()
        sql = "select TOP 80  CONCAT(t1.nombre,' ',t1.apellido) as nombre,documento,t1.ficha,t3.Nombre_curso,t1.Estado,t1.municipio from  aprendiz t1 inner join grupos t3 on t1.ficha=t3.ficha "
        datagridAS.DataSource = ListarDatos()
        datagridAS.Columns(4).Visible = False
        datagridAS.Columns(5).Visible = False
        AgregarcolorS()
        ' sql = "select TOP 80 CONCAT(t2.nombre,' ',t2.apellido)as aprendiz,documento, t2.ficha,t3.Nombre_curso,t2.estado from  aprendiz t2 inner join grupos t3 on t2.ficha=t3.ficha "

        ' datagridAS.DataSource = ListarDatos()
        ' datagridAS.Columns(4).Visible = False
        ' datagridAS.Rows(3).DefaultCellStyle.BackColor = Color.Green
        'datagridAS.Rows(3).Clone()
        'AgregarcolorS()
        cerrar_conexion()
    End Sub
    Private Sub datagridAS_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagridAS.CellClick
        Dim fila As Integer
        Dim query As String
        fila = datagridAS.CurrentRow.Cells(1).Value.ToString
        query = "select COUNt(t1.aprendiz)as seguimientos from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT  inner join modalidad t5 on t1.modalidad=t5.idmodalida join grupos t6 on t2.ficha=t6.ficha  where  t2.documento='" & fila & "'"
        conectado()
        cmd = New SqlCommand(query, cnn)
        reader = cmd.ExecuteReader
        If reader.Read Then
            If reader("seguimientos") > 0 Then

                sql = "select  TOP 80 documento,CONCAT(t2.nombre,' ',t2.apellido)as aprendiz, t3.razonSocial,fechaInicio,fechaFin,instructor,cargoAprendiz,t1.idseguimiento,t5.nombreModalidad as Modalidad,t7.nombreEstado, t6.Nombre_curso,t2.ficha from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT  inner join modalidad t5 on t1.modalidad=t5.idmodalida join grupos t6 on t2.ficha=t6.ficha inner join estado t7 on t1.estado=t7.idestado where t2.documento='" & fila & "'"
                reader.Close()
                detallesSAS.DataSource = ListarDatos()

                detallesSAS.Columns(0).Visible = False
                detallesSAS.Columns(7).Visible = False

                detallesSAS.ClearSelection()
                detallesSAS.CurrentCell = Nothing
                reader.Close()
                cerrar_conexion()
            Else
                detallesSAS.DataSource = DBNull.Value
                cerrar_conexion()
            End If
            reader.Close()
            cerrar_conexion()
        Else
            reader.Close()
            cerrar_conexion()
        End If

        cerrar_conexion()

    End Sub
    Sub ValidacionAsignacion()
        Try
            sql = "select  TOP 1 documento,CONCAT(t2.nombre,' ',t2.apellido)as aprendiz, t3.razonSocial,fechaInicio,fechaFin,instructor,cargoAprendiz,t1.idseguimiento,t5.nombreModalidad as Modalidad,t7.nombreEstado, t6.Nombre_curso,t2.ficha,fechaAsignacion from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT  inner join modalidad t5 on t1.modalidad=t5.idmodalida join grupos t6 on t2.ficha=t6.ficha inner join estado t7 on t1.estado=t7.idestado where t2.documento='" & txtdocumento.Text & "' order by fechaAsignacion desc"
            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader
            If reader.Read Then
                If reader("fechaAsignacion").ToString > Now.Date And reader("instructor").ToString <> "" Then
                    MessageBox.Show("Este aprendiz ya tiene una asignacion", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    reader.Close()
                End If

            Else
                reader.Close()
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub seguimiento_E_P_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sql = "SELECT * FROM [departamento]"
        llenarcombos(cmbdepartamentoaprendiz, "nombreDepartamento", "iddepartamento")
        llenarcombos(cmbdptE, "nombreDepartamento", "iddepartamento")
        sql = "SELECT * FROM modalidad"
        llenarcombos(cmbmodalidadSA, "nombreModalidad", "idmodalida")
        sql = "SELECT * FROM modalidad"
        llenarcombos(cmbmodalidadS, "nombreModalidad", "idmodalida")
        sql = "SELECT * FROM estado"
        llenarcombos(cmbestadoS, "nombreEstado", "idestado")
        ' sql = "SELECT * FROM estado"

        ' llenarcombos(cmbestadoEvaluacion, "nombreEstado", "idestado")

        sql = "SELECT * FROM estado"
        llenarcombos(cmbestadoSA, "nombreEstado", "idestado")
        '/*Datagrids*/'
        LlenarEmpresas()
        LlenarDataGridSAprendices()
        Agregarcolor()
        LlenarSeguimientosA()
        LlenarSeguimientoDetalleAprendiz()

        DataJefes()

        '
        ' LlenarJefes()
        cmbtipodoc.SelectedIndex = 0
        cmbestadoformacion.SelectedIndex = 0
        cmbestadoEvaluacion.SelectedIndex = 0
        '*********************************************************
        datagridaprendices.ClearSelection()
        datagridaprendices.CurrentCell = Nothing
        datagridseguimiento.ClearSelection()
        datagridseguimiento.CurrentCell = Nothing
        '*********************************************************

        '*********************************************************

        datagridEmpresa.ClearSelection()
        datagridEmpresa.CurrentCell = Nothing
        '*********************************************************
        '
        'datagridaprendices.ClearSelection()
        'datagridEvidencias.ClearSelection()
        'detallesSAS.ClearSelection()
        ' datagridEmpresa.ClearSelection()
        ' detallesSAS.ClearSelection()
        'datagridEmpA.ClearSelection()
        'datagridAE.ClearSelection()
        'datagridE.ClearSelection()


        If lbltipousuario.Text = "inspector" Then
            TabPage1.Parent = Nothing
            TabPage2.Parent = Nothing
            TabPage3.Parent = Nothing
            TabPage4.Parent = Nothing
            llenarAprendcesSeguimientoInspeccion()
        End If


    End Sub
    Sub llenarAprendcesSeguimientoInspeccion()
        sql = "select    ROW_NUMBER() OVER(ORDER BY idseguimiento ASC) AS 'No', t2.documento as 'Documento',CONCAT(t2.nombre,' ',t2.apellido)as 'Nombre Aprendiz',t2.telefono as 'Telefono Aprendiz',t2.correo AS 'Correo Aprendiz',t8.ficha as 'Ficha',t8.Nombre_curso as 'Nombre Curso', t3.razonSocial as 'Empresa',t5.nombreModalidad as 'Modalidad',fechaInicio as 'Fecha de Inicio',fechaFin as 'Fecha Fin',fechaAsignacion as 'Fecha Asignación',instructor as 'Instructor',t9.Correo as 'Correo Instructor',t9.telefono as 'Telefono Instructor',cargoAprendiz as 'Cargo del Aprendiz',t6.nombreEstado as 'Estado' from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT inner join modalidad t5 on t1.modalidad=t5.idmodalida inner join estado t6 on t1.Estado=t6.idestado inner join grupos t8 on t2.ficha=t8.ficha inner join instructores t9 on t1.instructor=t9.NOMBRE_FUNCIONARIO"
        dtinspeccionA.DataSource = ListarDatos()
        dtinspeccionA.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dtinspeccionA.ColumnHeadersDefaultCellStyle.Font = New Drawing.Font("Microsoft Sans Serif", 8.2, FontStyle.Bold)
        dtinspeccionA.Columns(0).Width = 30
        dtinspeccionA.Columns(1).Width = 76
        dtinspeccionA.Columns(2).Width = 170
        dtinspeccionA.Columns(3).Width = 79
        dtinspeccionA.Columns(4).Width = 150
        dtinspeccionA.Columns(5).Width = 58
        dtinspeccionA.Columns(6).Width = 105
        dtinspeccionA.Columns(7).Width = 150
        dtinspeccionA.Columns(8).Width = 75
        dtinspeccionA.Columns(9).Width = 75
        dtinspeccionA.Columns(10).Width = 75
        dtinspeccionA.Columns(11).Width = 72
        dtinspeccionA.Columns(12).Width = 220
        dtinspeccionA.Columns(13).Width = 150
        AgregarcolorInspeccion()
        dtinspeccionA.ClearSelection()
        dtinspeccionA.CurrentCell = Nothing
    End Sub
    Sub DataJefes()

        sql = "SELECT  TOP 80 [nombre],[cargoJefe] ,[telefonoJefe]   ,[empresa] ,[email],t2.[razonSocial] FROM [jefeInmediato] t1 inner join empresa t2 on t1.empresa=t2.[cedula_NIT] "
        datagridE.DataSource = ListarDatos()
        datagridE.Columns(3).Visible = False
        datagridE.Columns(4).Visible = False

    End Sub
    Private Sub btnnuevogE_Click(sender As Object, e As EventArgs) Handles btnnuevojE.Click

        If btnnuevojE.Text = "Nuevo" Then
            txtnombreJI.Text = ""
            txtcargoJ.Text = ""
            txttelefonoJ.Text = ""
            txtcorreoJ.Text = ""
            txtempresaE.Text = ""
            cmbempresaE.DataSource = Nothing
            cmbempresaE.Items.Clear()
            cmbempresaE.Enabled = False
            btnnuevojE.Text = "Guardar"
        Else
            Try
                If txtnombreJI.Text = "" Or
                    txtcargoJ.Text = "" Or
                    txttelefonoJ.Text = "" Or
                    txtcorreoJ.Text = "" Or
                     txtempresaE.Text = "" Or IsNothing(cmbempresaE.SelectedValue) Then
                    MsgBox("Existen campos vacíos")
                    Exit Sub
                Else
                    sql = "SELECT email from jefeInmediato  where email ='" & txtcorreoJ.Text & "'"
                    conectado()
                    cmd = New SqlCommand(sql, cnn)
                    reader = cmd.ExecuteReader
                    If reader.Read Then

                        MessageBox.Show("Esta Jefe Inmediato ya existe", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        reader.Close()
                        Exit Sub
                    Else
                        reader.Close()
                        sql = "INSERT INTO jefeInmediato "
                        sql += "VALUES('" & txtcargoJ.Text & "', '" & txttelefonoJ.Text & "', '" & txtnombreJI.Text & "', '" & cmbempresaE.SelectedValue.ToString & "', '" & txtcorreoJ.Text & "')"
                        Agregar()
                        LimpiarEmpleado()
                    End If


                End If

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If

        

    End Sub
    Sub LimpiarEmpleado()
        txtnombreJI.Text = ""
        txtcargoJ.Text = ""
        txttelefonoJ.Text = ""
        txtcorreoJ.Text = ""
        txtempresaE.Text = ""
        lbljefe.Text = ""
        ' btnnuevojE.Text = "Nuevo"
        cmbempresaE.DataSource = Nothing
        cmbempresaE.Items.Clear()
        cmbempresaE.Enabled = False

    End Sub
    Private Sub btneditarE_Click(sender As Object, e As EventArgs) Handles btneditarE.Click

        If btneditarE.Text = "Habilitar Campos" Then
            txtnombreJI.Enabled = True
            txtcargoJ.Enabled = True
            txttelefonoJ.Enabled = True
            txtcorreoJ.Enabled = True
            txtempresaE.Enabled = True
            cmbempresaE.Enabled = True
           
            btneditarE.Text = "Guardar"
        Else
            Try
                If lbljefe.Text = "" Then
                    MsgBox("Debe seleccionar un Empleado para editar")
                Else
                    If txtnombreJI.Text = "" Or txtcorreoJ.Text = "" Then
                        MsgBox("Existen campos obligatorios vacíos")
                    Else
                        sql = "UPDATE [jefeInmediato] set "
                        sql += "[cargoJefe]='" & txtcargoJ.Text & "',"
                        sql += "[telefonoJefe]='" & txttelefonoJ.Text & "',"
                        sql += "[nombre]='" & txtnombreJI.Text & "',"
                        sql += "[empresa]='" & cmbempresaE.SelectedValue.ToString & "',"
                        sql += "[email]='" & txtcorreoJ.Text & "'"
                        sql += "where [email]='" & lbljefe.Text & "'"
                        Agregar()
                        LlenarJefes()
                        LimpiarEmpleado()

                    End If


                End If

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If

       
    End Sub
    Private Sub txtinstructorSA_TextChanged(sender As Object, e As EventArgs) Handles txtinstructorSA.TextChanged
        If txtinstructorSA.Text = "" Then
            cmbinstructorSA.DataSource = Nothing
            cmbinstructorSA.Items.Clear()

        Else
            sql = "SELECT  TOP 80 NUMERO_IDENTIFICACION_FUNCIONARIO,NOMBRE_FUNCIONARIO FROM instructores where NOMBRE_FUNCIONARIO LIKE '%" + txtinstructorSA.Text + "%'"
            llenarcombos(cmbinstructorSA, "NOMBRE_FUNCIONARIO", "NOMBRE_FUNCIONARIO")

        End If

    End Sub
    Private Sub txtInstructorS_TextChanged(sender As Object, e As EventArgs) Handles txtInstructorS.TextChanged
        If txtInstructorS.Text = "" Then
            cmbInstructorS.DataSource = Nothing
            cmbInstructorS.Items.Clear()
            cmbInstructorS.Enabled = False
        Else
            sql = "SELECT  TOP 80 NUMERO_IDENTIFICACION_FUNCIONARIO,NOMBRE_FUNCIONARIO FROM instructores where NOMBRE_FUNCIONARIO LIKE '%" + txtInstructorS.Text + "%'"
            llenarcombos(cmbInstructorS, "NOMBRE_FUNCIONARIO", "NOMBRE_FUNCIONARIO")
            cmbInstructorS.Enabled = True
        End If
    End Sub
    Private Sub txtempresaS_TextChanged(sender As Object, e As EventArgs) Handles txtempresaS.TextChanged
        If txtempresaS.Text = "" Then
            cmbempresaS.DataSource = Nothing
            cmbempresaS.Items.Clear()
        Else
            sql = "SELECT  TOP 80 cedula_NIT,razonSocial FROM empresa where razonSocial LIKE '%" + txtempresaS.Text + "%'"
            llenarcombos(cmbempresaS, "razonSocial", "cedula_NIT")
            cmbempresaS.Enabled = True
        End If
    End Sub
    'sql = "SELECT documento,concat(nombre,' ',apellido)as nombrea FROM aprendiz where concat(nombre,' ',apellido) LIKE '%" + txtaprendiz.Text + "%'"
    'llenarcombos(cmbaprendiz, "nombrea", "documento")
    'cmbaprendiz.Enabled = True
    Private Sub btnbuscarFA_Click(sender As Object, e As EventArgs) Handles btnbuscarFA.Click
        If txtficha.Text = "" Then
            txtficha.Enabled = True
        Else
            sql = "SELECT ficha, Nombre_curso from [grupos]  where ficha ='" & txtficha.Text & "'"
            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader
            If reader.Read Then
                txtprogramaA.Text = reader("Nombre_curso")
                reader.Close()
            Else
                reader.Close()
            End If

            cerrar_conexion()
        End If

        ' MsgBox("entro " & txtficha.Text)
       
    End Sub
    Private Sub btneditarA_Click(sender As Object, e As EventArgs) Handles btneditarA.Click



        If btneditarA.Text = "Habilitar Campos" Then
            txtdocumento.Enabled = True
            txtnombreaprendiz.Enabled = True
            txtapellidoaprendiz.Enabled = True
            txtemailaprendiz.Enabled = True
            txtficha.Enabled = True
            txtdireccionaprendiz.Enabled = True
            txttelefonoaprendiz.Enabled = True
            txtemailaprendiz.Enabled = True
            cmbdepartamentoaprendiz.Enabled = True
            ' 
            cmbtipodoc.Enabled = True
            cmbmunicipioaprendiz.Enabled = False
            cmbestadoformacion.Enabled = True
            btneditarA.Text = "Editar"
        Else
            If iddocumento.Text = "" Then
                MsgBox("Debe seleccionar un aprendiz para editarlo")
            Else
                If txtdocumento.Text = "" Or txtnombreaprendiz.Text = "" Or
             txtapellidoaprendiz.Text = "" Or
             txtficha.Text = "" Or cmbmunicipioaprendiz.Text = "" Then
                    MessageBox.Show("Existen campos vacíos", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

                Else
                    Dim bln As Boolean = IsValidEmail(txtemailaprendiz.Text)
                    If bln = False Then
                        MessageBox.Show("Verifique Email Formato Incorrecto", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        txtemailaprendiz.Focus()
                        Exit Sub
                    Else
                        sql = "SELECT ficha, Nombre_curso from [grupos]  where ficha ='" & txtficha.Text & "'"
                        conectado()
                        cmd = New SqlCommand(sql, cnn)
                        reader = cmd.ExecuteReader
                        If reader.Read Then
                            sql = "UPDATE aprendiz set [documento]='" & txtdocumento.Text & "',"
                            sql += "[nombre]='" & txtnombreaprendiz.Text & "',"
                            sql += "[apellido]='" & txtapellidoaprendiz.Text & "',"
                            sql += "[correo]='" & txtemailaprendiz.Text & "',"
                            sql += "[telefono]='" & txttelefonoaprendiz.Text & "',"
                            sql += "[ficha]='" & txtficha.Text & "',"
                            sql += "[direccion]='" & txtdireccionaprendiz.Text & "',"
                            sql += "[municipio]='" & cmbmunicipioaprendiz.SelectedValue.ToString & "',"
                            sql += "[estado]='" & cmbestadoformacion.Text & "',"
                            sql += "[Tipo_documento]='" & cmbtipodoc.Text & "'"
                            sql += "where [documento]='" & iddocumento.Text & "'"
                            Agregar()
                            LlenarDataGridSAprendices()
                            LlenarSeguimientosA()
                            Agregarcolor()
                            limpiarAprendiz()
                        Else

                            MessageBox.Show("la ficha agregada no existe, verifique que esté bien escrita", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If

                    End If




                End If


            End If
            'dataSet.Clear()             ' Elimina la información antigua.
            'dataAdapter.Fill(DataSet)   ' Recarga la nueva información.
            ' DataGrid1.ResetBindings()   ' Vuelve a mostrar los datos.
        End If
    End Sub
    Private Sub datagridseguimiento_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagridseguimiento.CellClick

        seguimiento()
    End Sub
    Sub seguimiento()
        Try
            Dim fila As Integer
            fila = datagridseguimiento.CurrentRow.Index.ToString


            ' txtaprendiz.Text = datagridseguimiento.Rows(fila).Cells(0).Value.ToString
            txtempresaSA.Text = datagridseguimiento.Rows(fila).Cells(1).Value.ToString
            cmbempresaSA.Text = datagridseguimiento.Rows(fila).Cells(1).Value.ToString
            If datagridseguimiento.Rows(fila).Cells(2).Value.Equals(DBNull.Value) Then
            Else
                datetimeinicio.Value = datagridseguimiento.Rows(fila).Cells(2).Value.ToString
            End If
            If datagridseguimiento.Rows(fila).Cells(3).Value.Equals(DBNull.Value) Then

            Else
                datetimefin.Value = datagridseguimiento.Rows(fila).Cells(3).Value.ToString
            End If
            txtinstructorSA.Text = datagridseguimiento.Rows(fila).Cells(4).Value.ToString
            cmbinstructorSA.Text = datagridseguimiento.Rows(fila).Cells(4).Value.ToString
            txtcargoA.Text = datagridseguimiento.Rows(fila).Cells(5).Value.ToString
            lblconfirmacion.Text = datagridseguimiento.Rows(fila).Cells(6).Value.ToString
            cmbmodalidadSA.Text = datagridseguimiento.Rows(fila).Cells(7).Value.ToString
            cmbestadoSA.Text = datagridseguimiento.Rows(fila).Cells(9).Value.ToString
            'txtjefeISA.Text = datagridseguimiento.Rows(fila).Cells(8).Value.ToString
            'cmbjefeISA.Text = datagridseguimiento.Rows(fila).Cells(8).Value.ToString


        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub txtnombreaprendiz_TextChanged(sender As Object, e As EventArgs) Handles txtnombreaprendiz.TextChanged
        If txtnombreaprendiz.Text = "" Then
            txtnombreaprendiz.Enabled = False

        End If
    End Sub
    Private Sub btneditarSeguimiento_Click(sender As Object, e As EventArgs) Handles btneditarSeguimiento.Click
        Try
            If btneditarSeguimiento.Text = "Habilitar Campos" Then
                cmbinstructorSA.Enabled = True
                txtinstructorSA.Enabled = True
                txtempresaSA.Enabled = True
                txtcargoA.Enabled = True
                datetimeinicio.Enabled = True
                datetimefin.Enabled = True
                cmbmodalidadSA.Enabled = True
                cmbempresaSA.Enabled = True
                lblestatus_user.Enabled = True
                cmbestadoSA.Enabled = True
                btneditarSeguimiento.Text = "Editar"
            Else
                If lblconfirmacion.Text = "" Then
                    MsgBox("Debe seleccionar un seguimiento para editarlo")
                Else
                    If lblconfirmacion.Text = "" Then
                        MsgBox("Existen campos vacíos")
                    Else
                        sql = "select instructor from seguimiento   where idseguimiento ='" & lblconfirmacion.Text & "'"
                        conectado()
                        cmd = New SqlCommand(sql, cnn)
                        reader = cmd.ExecuteReader
                        If reader.Read Then
                            If reader("instructor") <> cmbinstructorSA.Text.ToString Then
                                reader.Close()
                                updateS()
                                Try
                                    reader.Close()
                                    Dim consulta, aprendiz, instructor As String
                                    consulta = "select   Tipo_documento,documento,concat(t1.nombre,' ',apellido) as aprendiz,t1.ficha as fichaAprendiz, t3.Nivel,t2.fechaInicio,t2.fechaFin, t3.Nombre_curso as programa,t1.telefono as telefonoAprendiz,t1.correo AS CorreoAprendiz, t5.cedula_NIT,t5.razonSocial,t6.Municipio as nombreMuni,t5.direccion as direccionE,t5.telefono as telefonoE, t7.nombreDepartamento as departamentoEmpresa, t4.NUMERO_IDENTIFICACION_FUNCIONARIO,t4.NOMBRE_FUNCIONARIO,t4.Correo as CorreoInstructor,t4.Telefono as telefonoI, [idseguimiento] from aprendiz t1 inner join seguimiento t2 on t1.documento=t2.aprendiz inner join grupos t3 on t3.ficha=t1.ficha inner join instructores t4 on t2.instructor=t4.NOMBRE_FUNCIONARIO inner join empresa t5 on t2.empresa=t5.cedula_NIT inner join municipios t6 on t5.municipio=t6.Id inner join departamento t7 on t6.departamento=t7.iddepartamento where idseguimiento ='" & lblconfirmacion.Text & "'"
                                    conectado()
                                    cmd = New SqlCommand(consulta, cnn)
                                    reader = cmd.ExecuteReader

                                    If reader.Read Then
                                        ' MsgBox("")
                                        If reader("CorreoAprendiz").ToString = "" Or reader("CorreoAprendiz").ToString = DBNull.Value.ToString Then

                                            instructor = reader("CorreoInstructor")
                                            para = instructor '"khattherine@gmail.com" 
                                        Else
                                            aprendiz = reader("CorreoAprendiz")
                                            instructor = reader("CorreoInstructor") ' "khattherine@gmail.com"
                                            para = aprendiz & ";" & instructor '"khattherine@gmail.com" 
                                        End If


                                        cuerpo = "<HTML><BODY><h3 >Se ha registrado un seguimiento de etapa productiva.</h3>" & "<br>"
                                        cuerpo += vbCrLf
                                        cuerpo += vbCrLf
                                        cuerpo += vbCrLf & "<b><p style='font-family:Engravers MT;font-size:38'>Informacion del Aprendiz:</p></b><br>"
                                        cuerpo += vbCrLf
                                        cuerpo += vbCrLf & "Aprendiz: " & reader("aprendiz") & "<br>"
                                        cuerpo += vbCrLf
                                        cuerpo += vbCrLf & "Documento: " & reader("documento") & "<br>"
                                        cuerpo += vbCrLf
                                        cuerpo += vbCrLf & "Telefono: " & reader("telefonoAprendiz") & "<br>"
                                        cuerpo += vbCrLf
                                        cuerpo += vbCrLf & "Ficha: " & reader("fichaAprendiz") & "<br>"
                                        cuerpo += vbCrLf
                                        ficha_aprendiz = reader("fichaAprendiz")
                                        cuerpo += vbCrLf & "Nivel de programa de Formacion: " & reader("Nivel") & "<br>"
                                        cuerpo += vbCrLf
                                        cuerpo += vbCrLf & "Programa de Formacion: " & reader("programa") & "<br>"
                                        cuerpo += vbCrLf & "<br>"
                                        cuerpo += vbCrLf
                                        cuerpo += vbCrLf & "<b><p style='font-family:Engravers MT;font-size:30'>Informacion de la Empresa:</p></b>" & "<br>"
                                        cuerpo += vbCrLf
                                        cuerpo += vbCrLf & "NIT: " & reader("cedula_NIT") & "<br>"
                                        cuerpo += vbCrLf
                                        cuerpo += vbCrLf & "Empresa: " & reader("razonSocial") & "<br>"
                                        cuerpo += vbCrLf
                                        cuerpo += vbCrLf & "Municipio: " & reader("nombreMuni") & "<br>"
                                        cuerpo += vbCrLf
                                        cuerpo += vbCrLf & "Direccion: " & reader("direccionE") & "<br>"
                                        cuerpo += vbCrLf
                                        cuerpo += vbCrLf & "Telefono: " & reader("telefonoE") & "<br>"
                                        cuerpo += vbCrLf
                                        cuerpo += vbCrLf & "<br>"
                                        cuerpo += vbCrLf & "<b><p style='font-family:Engravers MT;font-size:30'>Informacion del instructor responsable del seguimiento:</p></b>" & "<br>"
                                        cuerpo += vbCrLf

                                        cuerpo += vbCrLf & "Instructor: " & reader("NOMBRE_FUNCIONARIO") & "<br>"
                                        cuerpo += vbCrLf
                                        cuerpo += vbCrLf & "Telefono: " & reader("telefonoI") & "<br>"
                                        cuerpo += vbCrLf
                                        cuerpo += vbCrLf & "Correo: " & instructor & "<br>"
                                        cuerpo += vbCrLf

                                        ' cuerpo += vbCrLf & "Se ha asignado seguimiento "

                                        cuerpo += vbCrLf
                                        cuerpo += vbCrLf & "<br>"

                                        cuerpo += vbCrLf & "<b><p style='font-family:Calibri;font-size:30'>Cordialmente:</p></b>"

                                        cuerpo += vbCrLf & " <b> <p style='font-family:Arial Rounded MT Bold'>JAVIER CARRILLO PINTO</p></b>"
                                        cuerpo += vbCrLf & "<b> <p style='font-family:Calibri '>Coordinador Academico</p></b></BODY></HTML>"

                                        'adjunto = "c:\cartas_seguimiento\" & ficha_aprendiz & "\" & reader("aprendiz") & ".doc"
                                        asunto = "Seguimiento Aprendiz"

                                        enviar_correoseg()
                                        print_carta_seguimiento()
                                        If reader("CorreoAprendiz").ToString = "" Or reader("CorreoAprendiz").ToString = DBNull.Value.ToString Then


                                            Try
                                                sql = "UPDATE seguimiento set "
                                                sql += "[confirmacionEA]='0',"
                                                sql += "[confirmacionEI]='1'"
                                                sql += "where [idseguimiento]=" & reader("idseguimiento") & ""
                                                Agregar()
                                            Catch ex As Exception
                                                MsgBox(ex.ToString)
                                            End Try
                                        Else

                                            Try
                                                sql = "UPDATE seguimiento set "
                                                sql += "[confirmacionEA]='1',"
                                                sql += "[confirmacionEI]='1'"
                                                sql += "where [idseguimiento]=" & reader("idseguimiento") & ""
                                                Agregar()
                                            Catch ex As Exception
                                                MsgBox(ex.ToString)
                                            End Try
                                        End If





                                        reader.Close()
                                        cerrar_conexion()
                                    Else

                                        AvisoAsignacionInstructor(txtdocumento.Text)
                                        reader.Close()
                                        cerrar_conexion()
                                    End If


                                Catch ex As Exception
                                    MsgBox(ex.ToString)
                                End Try

                            Else
                                reader.Close()
                                updateS()
                            End If
                        Else
                            reader.Close()
                            updateS()
                        End If


                    End If


                End If
            End If
            cerrar_conexion()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
    Sub updateS()
        sql = "UPDATE seguimiento set "
        sql += "[instructor]='" & cmbinstructorSA.Text.ToString & "',"
        sql += "[empresa]='" & cmbempresaSA.SelectedValue.ToString & "',"
        sql += "[fechaInicio]='" & datetimeinicio.Value & "',"
        sql += "[fechaFin]='" & datetimefin.Value & "',"
        sql += "[modalidad]='" & cmbmodalidadSA.SelectedValue.ToString & "',"
        sql += "[usuario]='" & lblestatus_user.Text & "',"
        sql += "[estado]='" & cmbestadoSA.SelectedValue.ToString & "'"
        sql += "where [idseguimiento]='" & lblconfirmacion.Text & "'"
        Agregar()
        LlenarSeguimientosA()
        LimpiarSeguimiento()
    End Sub
    Private Sub redireccionNuevoJ_Click(sender As Object, e As EventArgs)
        TabControl1.SelectedIndex = 3
    End Sub
    Private Sub redireccionNuevaE_Click(sender As Object, e As EventArgs) Handles redireccionNuevaE.Click
        TabControl1.SelectedIndex = 3
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs)
        MessageBox.Show("Este usuario no puede realizar la siguiente accion", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'Form1.Show()
        'Form1.TabControl1.SelectedIndex = 1
    End Sub
    Private Sub btnredireccionE_Click(sender As Object, e As EventArgs) Handles btnredireccionE.Click
        TabControl1.SelectedIndex = 3
    End Sub
    Private Sub btndetalleSA_Click(sender As Object, e As EventArgs) Handles btndetalleSA.Click
        TabControl1.SelectedIndex = 1
        LlenarDetalleSeguimientosA()
        txtaprendizSe.Text = datagridseguimiento.CurrentRow.Cells(0).Value.ToString
        txtInstructorS.Text = datagridseguimiento.CurrentRow.Cells(4).Value.ToString
        txtempresaS.Text = datagridseguimiento.CurrentRow.Cells(1).Value.ToString
        txtcargoAS.Text = datagridseguimiento.CurrentRow.Cells(5).Value.ToString
        If datagridseguimiento.CurrentRow.Cells(2).Value.Equals(DBNull.Value) Then
        Else
            datetimeInicioS.Value = datagridseguimiento.CurrentRow.Cells(2).Value.ToString
        End If
        If datagridseguimiento.CurrentRow.Cells(3).Value.Equals(DBNull.Value) Then
        Else
            datetimefinS.Value = datagridseguimiento.CurrentRow.Cells(3).Value.ToString
        End If



        cmbmodalidadS.Text = datagridseguimiento.CurrentRow.Cells(7).Value.ToString
        'txtjefeIS.Text = datagridseguimiento.CurrentRow.Cells(8).Value.ToString

        'cmbestadoS.Text = datagridE.CurrentRow.Cells(4).Value.ToString

    End Sub
    Sub LlenarDetalleSeguimientosA()
        sql = "select  TOP 80  documento,CONCAT(t2.nombre,' ',t2.apellido)as aprendiz, t3.razonSocial,fechaInicio,fechaFin,instructor,cargoAprendiz,t1.idseguimiento,t5.nombreModalidad as Modalidad,t7.nombreEstado, t6.Nombre_curso,t2.ficha from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT  inner join modalidad t5 on t1.modalidad=t5.idmodalida join grupos t6 on t2.ficha=t6.ficha inner join estado t7 on t1.estado=t7.idestado where t2.documento='" & datagridseguimiento.CurrentRow.Cells(8).Value.ToString & "'"
        detallesSAS.DataSource = ListarDatos()

        detallesSAS.Columns(0).Visible = False
        detallesSAS.Columns(7).Visible = False


        ' datagridseguimiento.Columns(4).Visible = False
    End Sub
    Sub LlenarDetalleSeguimientosAs()
        sql = "select  TOP 80 documento,CONCAT(t2.nombre,' ',t2.apellido)as aprendiz, t3.razonSocial,fechaInicio,fechaFin,instructor,cargoAprendiz,t1.idseguimiento,t5.nombreModalidad as Modalidad,t7.nombreEstado, t6.Nombre_curso,t2.ficha from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT  inner join modalidad t5 on t1.modalidad=t5.idmodalida join grupos t6 on t2.ficha=t6.ficha inner join estado t7 on t1.estado=t7.idestado where t2.documento='" & detallesSAS.CurrentRow.Cells(0).Value.ToString & "'"
        detallesSAS.DataSource = ListarDatos()

        detallesSAS.Columns(0).Visible = False
        detallesSAS.Columns(7).Visible = False


        ' datagridseguimiento.Columns(4).Visible = False
    End Sub
    Sub llenarAprendicesJefes()
        Try
            Dim jefeInmediato As String
            jefeInmediato = datagridE.CurrentRow.Cells(4).Value.ToString


            sql = "select TOP 80 CONCAT(t2.nombre,' ',t2.apellido)as aprendiz,t2.ficha,t5.Nombre_curso as Programa from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT inner join jefeInmediato t4 on t1.jefeInmediato=t4.email inner join grupos t5 on t2.ficha=t5.ficha where t1.jefeInmediato='" & jefeInmediato & "'"
            ' MsgBox(sql)
            datagridAE.DataSource = ListarDatos()
            datagridAE.AllowUserToAddRows = False





        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try





    End Sub
    Private Sub datagridE_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagridE.CellClick

        JefesInmediatos()
    End Sub
    Sub JefesInmediatos()

        Try
            If datagridE.CurrentRow Is Nothing Then

            Else
                Dim fila As Integer = datagridE.CurrentRow.Index.ToString
                'datagridE.Rows(0).Selected = True

                'fila = 

                lbljefe.Text = datagridE.Rows(fila).Cells(0).Value.ToString
                txtnombreJI.Text = datagridE.Rows(fila).Cells(0).Value.ToString
                txtcargoJ.Text = datagridE.Rows(fila).Cells(1).Value.ToString
                txttelefonoJ.Text = datagridE.Rows(fila).Cells(2).Value.ToString
                txtcorreoJ.Text = datagridE.Rows(fila).Cells(4).Value.ToString
                txtempresaE.Text = datagridE.Rows(fila).Cells(5).Value.ToString
                cmbempresaE.Text = datagridE.Rows(fila).Cells(5).Value.ToString
                llenarAprendicesJefes()
            End If
        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try
    End Sub
    Private Sub datagridseguimiento_DataSourceChanged(sender As Object, e As EventArgs) Handles datagridseguimiento.DataSourceChanged

        If datagridseguimiento.RowCount > 0 Then
            seguimiento()
            ' LimpiarSeguimiento()
        Else
            LimpiarSeguimiento()
        End If

    End Sub
    Private Sub datagridE_DataSourceChanged(sender As Object, e As EventArgs) Handles datagridE.DataSourceChanged
        LimpiarEmpleado()
        If datagridE.RowCount > 0 Then


            JefesInmediatos()
        Else
            datagridAE.DataSource = DBNull.Value
        End If
    End Sub
    Private Sub txtaprendizSe_TextChanged(sender As Object, e As EventArgs) Handles txtaprendizSe.TextChanged
        If txtaprendizSe.Text <> "" Then
            sql = "select TOP 80 CONCAT(t2.nombre,' ',t2.apellido)as aprendiz,documento, t2.ficha,t3.Nombre_curso,t2.estado from  aprendiz t2 inner join grupos t3 on t2.ficha=t3.ficha where CONCAT(t2.nombre,' ',t2.apellido) like '%" & txtaprendizSe.Text & "%'"
            datagridAS.DataSource = ListarDatos()
            AgregarcolorS()
            datagridAS.ClearSelection()
            datagridAS.CurrentCell = Nothing
            cerrar_conexion()
            ' txtaprendizSe.Enabled = True
        Else
            sql = "select TOP 80 CONCAT(t2.nombre,' ',t2.apellido)as aprendiz,documento, t2.ficha,t3.Nombre_curso,t2.estado from  aprendiz t2 inner join grupos t3 on t2.ficha=t3.ficha"
            datagridAS.DataSource = ListarDatos()
            AgregarcolorS()
            datagridAS.ClearSelection()
            datagridAS.CurrentCell = Nothing
            cerrar_conexion()
            'txtaprendizSe.Enabled = False
        End If
        cerrar_conexion()
    End Sub
    Sub Agregarcolor()
        Dim i As Integer
        For i = 0 To datagridaprendices.Rows.Count - 1
            'MsgBox(datagridaprendices.Rows.Count)
            ' MsgBox(datagridaprendices.Rows(i).Cells(6).Value.ToString())
            If datagridaprendices.Rows(i).Cells(6).Value.ToString = "CANCELADO" Or datagridaprendices.Rows(i).Cells(6).Value.ToString = "CONDICIONADO" Or datagridaprendices.Rows(i).Cells(6).Value.ToString = "TRASLADADO" Or datagridaprendices.Rows(i).Cells(6).Value.ToString = "RETIRO VOLUNTARIO" Then
                datagridaprendices.Rows(i).DefaultCellStyle.BackColor = Color.Red
            ElseIf (datagridaprendices.Rows(i).Cells(6).Value.ToString = "POR CERTIFICAR") Then
                datagridaprendices.Rows(i).DefaultCellStyle.BackColor = Color.YellowGreen
            ElseIf (datagridaprendices.Rows(i).Cells(6).Value.ToString = "CERTIFICADO") Then
                datagridaprendices.Rows(i).DefaultCellStyle.BackColor = Color.Green
            ElseIf (datagridaprendices.Rows(i).Cells(6).Value.ToString = "") Then
                datagridaprendices.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
            End If

            '    datagridaprendices.Rows(i).DefaultCellStyle.BackColor = Color.Red

            '    datagridaprendices.Rows(i).DefaultCellStyle.BackColor = Color.Green
        Next



    End Sub
    Sub LlenarDataGridSAprendices()


        ' sql = "select TOP 80 [documento],[nombre],[apellido],[telefono],correo,[ficha],[Estado],t2.Municipio,t1.municipio, [Tipo_documento],direccion,t3.nombreDepartamento from aprendiz t1 inner join municipios t2 on t1.municipio=t2.Id inner join departamento t3 on t2.departamento=t3.iddepartamento"
        sql = "select TOP 80 [documento],[nombre],[apellido],[telefono],correo,[ficha],[Estado],t1.municipio, [Tipo_documento],direccion from aprendiz t1 "
        datagridaprendices.DataSource = ListarDatos()
        datagridaprendices.Columns(7).Visible = False

        'datagridaprendices.Columns(6).Visible = False
        ' datagridaprendices.Columns(8).Visible = False
        'datagridaprendices.Columns(9).Visible = False
        cerrar_conexion()
    End Sub
    Sub LlenarSeguimientosA()
        sql = "select TOP 80 CONCAT(t2.nombre,' ',t2.apellido)as aprendiz, t3.razonSocial,fechaInicio,fechaFin,instructor,cargoAprendiz,t1.idseguimiento,t5.nombreModalidad,t2.documento,t6.nombreEstado from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT inner join modalidad t5 on t1.modalidad=t5.idmodalida inner join estado t6 on t1.Estado=t6.idestado"
        datagridseguimiento.DataSource = ListarDatos()
        datagridseguimiento.Columns(4).Visible = False
        datagridseguimiento.Columns(5).Visible = False
        datagridseguimiento.Columns(6).Visible = False
        datagridseguimiento.Columns(8).Visible = False

        cerrar_conexion()

        ' datagridseguimiento.Columns(4).Visible = False
    End Sub
    Sub LlenarEmpresas()
        sql = "SELECT  TOP 80 t1.*,t2.Municipio,t3.nombreDepartamento FROM empresa t1 inner join municipios t2 on t1.municipio=t2.Id inner join departamento t3 on t2.departamento=t3.iddepartamento"
        datagridEmpresa.DataSource = ListarDatos()
        datagridEmpresa.Columns(5).Visible = False
        datagridEmpresa.Columns(6).Visible = False
        cerrar_conexion()
    End Sub
    Sub LlenarJefes()
        sql = "SELECT  TOP 80 [nombre],[cargoJefe] ,[telefonoJefe]   ,[empresa] ,[email],t2.[razonSocial] FROM [jefeInmediato] t1 inner join empresa t2 on t1.empresa=t2.[cedula_NIT]"
        datagridE.DataSource = ListarDatos()
        datagridE.Columns(3).Visible = False
        datagridE.Columns(4).Visible = False

        cerrar_conexion()
    End Sub
    Sub AvisoAsignacionInstructorPersonalizado()
        reader.Close()
        Dim query, coordinador As String
        query = "select TOP 1  Tipo_documento,documento,concat(t1.nombre,' ',apellido) as aprendiz,t1.ficha as fichaAprendiz, t3.Nivel,t2.fechaInicio,t2.fechaFin, t3.Nombre_curso as programa,t1.telefono as telefonoAprendiz,t1.correo AS CorreoAprendiz, t5.cedula_NIT,t5.razonSocial,t6.Municipio as nombreMuni,t5.direccion as direccionE,t5.telefono as telefonoE, t7.nombreDepartamento as departamentoEmpresa,t2.instructor from aprendiz t1 inner join seguimiento t2 on t1.documento=t2.aprendiz inner join grupos t3 on t3.ficha=t1.ficha  inner join empresa t5 on t2.empresa=t5.cedula_NIT inner join municipios t6 on t5.municipio=t6.Id inner join departamento t7 on t6.departamento=t7.iddepartamento where t2.aprendiz='" & txtdocumento.Text & "' ORDER BY fechaAsignacion DESC"
        conectado()
        cmd = New SqlCommand(query, cnn)
        reader = cmd.ExecuteReader
        Dim null As String = "NULL"
        If reader.Read Then
            If reader("instructor") = "" Or reader("instructor") = null Then
                '
                coordinador = "ecpitrebo@misena.edu.co" '"khattherine@gmail.com"
                para = coordinador
                'style='color:#80BFFF'"
                cuerpo = "<HTML><BODY><h3 >Estimado Coordinador, el presente correo es para informarle que el siguiente aprendiz aún no se le ha asignado un instructor para el seguimiento de la etapa practica, por favor diligenciar la asignacion del instructor.</h3><br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "<b><p style='font-family:Engravers MT;font-size:30'>Informacion del Aprendiz:</p></b><br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Aprendiz: " & reader("aprendiz") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Documento: " & reader("documento") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Telefono: " & reader("telefonoAprendiz") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "<Ficha: " & reader("fichaAprendiz") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Nivel de programa de Formacion: " & reader("Nivel") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Programa de Formacion: " & reader("programa")
                cuerpo += vbCrLf & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "<b><p style='font-family:Engravers MT;font-size:30'>Informacion de la Empresa:</p></b>" & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "NIT: " & reader("cedula_NIT") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Empresa: " & reader("razonSocial") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Municipio: " & reader("nombreMuni") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Direccion: " & reader("direccionE") & "<br>"
                cuerpo += vbCrLf
                cuerpo += vbCrLf & "Telefono: " & reader("telefonoE") & "<br>"

                cuerpo += vbCrLf & "<br>"
                cuerpo += vbCrLf & "<br>"

                cuerpo += vbCrLf & "<p style='font-family:Engravers MT;font-size:40'>Cordialmente:</p>" & "<br>"

                cuerpo += vbCrLf & "<b> <p style='font-family:Monotype Corsiva;font-size:50'>  Sistema Automatico de Seguimiento de Aprendices</p></b></BODY></HTML>"



                asunto = "Informacion de Seguimiento de Aprendices en etapa practica"

                '  enviar_correosegpersonalizado()
                reader.Close()
            End If
            reader.Close()
        Else
            MsgBox("no se envio el correo")
            reader.Close()
        End If
        reader.Close()


    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'ImportarWord()
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            If datagridAS.CurrentRow IsNot Nothing Then
                sql = "select TOP 1  Tipo_documento,documento,concat(t1.nombre,' ',apellido) as aprendiz,t1.ficha as fichaAprendiz, t3.Nivel,t3.Nombre_curso as programa,t1.telefono as telefonoAprendiz,t1.correo AS CorreoAprendiz, t5.cedula_NIT,t5.razonSocial,t6.Municipio,t7.nombreDepartamento as departamentoEmpresa, t4.NUMERO_IDENTIFICACION_FUNCIONARIO,t4.NOMBRE_FUNCIONARIO,t4.Correo as CorreoInstructor from aprendiz t1 inner join seguimiento t2 on t1.documento=t2.aprendiz inner join grupos t3 on t3.ficha=t1.ficha inner join instructores t4 on t2.instructor=t4.NOMBRE_FUNCIONARIO inner join empresa t5 on t2.empresa=t5.cedula_NIT inner join municipios t6 on t5.municipio=t6.Id inner join departamento t7 on t6.departamento=t7.iddepartamento where t2.aprendiz='" & datagridAS.CurrentRow.Cells(1).Value.ToString & "' ORDER BY fechaAsignacion DESC"
                conectado()
                cmd = New SqlCommand(sql, cnn)
                reader = cmd.ExecuteReader

                If reader.Read Then
                    ficha_aprendiz = reader("fichaAprendiz")
                    print_carta_seguimiento()
                Else
                    MessageBox.Show("Este aprendiz no tiene registrados seguimientos", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    reader.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try




    End Sub
    Private Sub btnnuevogS_Click(sender As Object, e As EventArgs) Handles btnnuevogS.Click
        If btnnuevogS.Text = "Nuevo" Then
            txtInstructorS.Text = ""
            txtInstructorS.Enabled = True
            cmbInstructorS.Enabled = True
            txtempresaS.Text = ""
            txtempresaS.Enabled = True
            cmbempresaS.Enabled = True
            cmbestadoS.Enabled = True
            'cmbjefeISA.Enabled = True
            txtcargoAS.Enabled = True
            txtcargoAS.Text = ""
            ' txtjefeISA.Text = ""
            cmbmodalidadS.Enabled = True
            btnnuevogS.Text = "Guardar"
        End If

        Try

            If datagridAS.CurrentRow IsNot Nothing Then
                If txtempresaS.Text = "" Then
                    MessageBox.Show("Debe seleccionar una empresa", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    If datetimeInicioS.Value > datetimefinS.Value Then
                        MessageBox.Show("La fecha de inicio de etapa practica no puede ser mayor a la fecha final", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        sql = "select Estado from aprendiz   where documento ='" & datagridAS.CurrentRow.Cells(1).Value.ToString & "'"
                        conectado()
                        cmd = New SqlCommand(sql, cnn)
                        reader = cmd.ExecuteReader
                        If reader.Read Then
                            ' MsgBox(reader("Estado").ToString)
                            If reader("Estado").ToString = "" Or reader("Estado").ToString = DBNull.Value.ToString Or reader("Estado").ToString = "CANCELADO" Or reader("Estado").ToString = "TRASLADADO" Or reader("Estado").ToString = "CONDICIONADO" Or reader("Estado").ToString = "RETIRO VOLUNTARIO" Then
                                MessageBox.Show("Este aprendiz no tiene un estado aceptado por el sistema", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                reader.Close()
                                Exit Sub
                            Else
                                reader.Close()
                                sql = "INSERT INTO seguimiento (instructor, empresa, cargoAprendiz, fechaInicio, fechaFin, usuario, modalidad, fechaAsignacion ,estado, aprendiz, hora ) "
                                sql += "VALUES ('" & cmbInstructorS.Text & "', '" & cmbempresaS.SelectedValue.ToString & "','" & txtcargoAS.Text & "','" & datetimeInicioS.Value & "','" & datetimefinS.Value & "','" & lblestatus_user.Text & "', " & cmbmodalidadS.SelectedValue.ToString & ",'" & DateTimePicker1.Value.Date & "', '" & cmbestadoS.SelectedValue.ToString & "'," & datagridAS.CurrentRow.Cells(1).Value.ToString & ",'" & Now.ToString("HH:mm:ss") & "')"
                                'MsgBox(sql)
                                Agregar()
                                LlenarSeguimientoDetalleAprendiz()
                                AgregarcolorS()
                                LimpiarSeguimientoA()
                                cerrar_conexion()

                                Dim consulta, aprendiz, instructor As String
                                consulta = "select TOP 1  Tipo_documento,documento,concat(t1.nombre,' ',apellido) as aprendiz,t1.ficha as fichaAprendiz, t3.Nivel,t2.fechaInicio,t2.fechaFin, t3.Nombre_curso as programa,t1.telefono as telefonoAprendiz,t1.correo AS CorreoAprendiz, t5.cedula_NIT,t5.razonSocial,t6.Municipio as nombreMuni,t5.direccion as direccionE,t5.telefono as telefonoE, t7.nombreDepartamento as departamentoEmpresa, t4.NUMERO_IDENTIFICACION_FUNCIONARIO,t4.NOMBRE_FUNCIONARIO,t4.Correo as CorreoInstructor,t4.Telefono as telefonoI, [idseguimiento] from aprendiz t1 inner join seguimiento t2 on t1.documento=t2.aprendiz inner join grupos t3 on t3.ficha=t1.ficha inner join instructores t4 on t2.instructor=t4.NOMBRE_FUNCIONARIO inner join empresa t5 on t2.empresa=t5.cedula_NIT inner join municipios t6 on t5.municipio=t6.Id inner join departamento t7 on t6.departamento=t7.iddepartamento where t2.aprendiz='" & txtdocumento.Text & "' ORDER BY fechaAsignacion DESC"
                                conectado()
                                cmd = New SqlCommand(consulta, cnn)
                                reader = cmd.ExecuteReader

                                If reader.Read Then
                                    ' MsgBox("")
                                    aprendiz = reader("CorreoAprendiz")
                                    instructor = reader("CorreoInstructor")

                                    para = aprendiz & ";" & instructor
                                    cuerpo = "<HTML><BODY><h3 >Se ha registrado un seguimiento de etapa productiva.</h3>" & "<br>"
                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf & "<b><p style='font-family:Engravers MT;font-size:30'>Informacion del Aprendiz:</p></b><br>"
                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf & "Aprendiz: " & reader("aprendiz") & "<br>"
                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf & "Documento: " & reader("documento") & "<br>"
                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf & "Telefono: " & reader("telefonoAprendiz") & "<br>"
                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf & "Ficha: " & reader("fichaAprendiz") & "<br>"
                                    ficha_aprendiz = reader("fichaAprendiz")
                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf & "Nivel de programa de Formacion: " & reader("Nivel") & "<br>"
                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf & "Programa de Formacion: " & reader("programa") & "<br>"
                                    cuerpo += vbCrLf & "<br>"
                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf & "<b><p style='font-family:Engravers MT;font-size:30'>Informacion de la Empresa:</p></b>" & "<br>"
                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf & "NIT: " & reader("cedula_NIT") & "<br>"
                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf & "Empresa: " & reader("razonSocial") & "<br>"
                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf & "Municipio: " & reader("nombreMuni") & "<br>"
                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf & "Direccion: " & reader("direccionE") & "<br>"
                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf & "Telefono: " & reader("telefonoE") & "<br>"
                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf & "<br>"
                                    cuerpo += vbCrLf & "<b><p style='font-family:Engravers MT;font-size:30'>Informacion del instructor responsable del seguimiento:</p></b>" & "<br>"
                                    cuerpo += vbCrLf

                                    cuerpo += vbCrLf & "Instructor: " & reader("NOMBRE_FUNCIONARIO") & "<br>"
                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf & "Telefono: " & reader("telefonoI") & "<br>"
                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf & "Correo: " & instructor & "<br>"
                                    cuerpo += vbCrLf

                                    ' cuerpo += vbCrLf & "Se ha asignado seguimiento "

                                    cuerpo += vbCrLf
                                    cuerpo += vbCrLf & "<br>"

                                    cuerpo += vbCrLf & "<b><p style='font-family:Calibri;font-size:30'>Cordialmente:</p></b>"

                                    cuerpo += vbCrLf & " <b> <p style='font-family:Arial Rounded MT Bold'>JAVIER CARRILLO PINTO</p></b>"
                                    cuerpo += vbCrLf & "<b> <p style='font-family:Calibri '>Coordinador Academico</p></b></BODY></HTML>"


                                    asunto = "Seguimiento Aprendiz"

                                    enviar_correoseg()

                                    Try
                                        sql = "UPDATE seguimiento set "
                                        sql += "[confirmacionEA]='1',"
                                        sql += "[confirmacionEI]='1'"
                                        sql += "where [idseguimiento]=" & reader("idseguimiento") & ""
                                        Agregar()
                                    Catch ex As Exception
                                        MsgBox(ex.ToString)
                                    End Try


                                    print_carta_seguimiento()
                                    reader.Close()
                                    cerrar_conexion()
                                Else

                                    AvisoAsignacionInstructor(datagridAS.CurrentRow.Cells(1).Value.ToString)
                                    reader.Close()
                                    cerrar_conexion()
                                End If
                            End If
                        End If


                    End If

                End If

            Else
                MessageBox.Show("Debe Seleccionar un Aprendiz para hacer el seguimiento.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            cerrar_conexion()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
    Sub LimpiarSeguimientoA()



        txtInstructorS.Text = ""
        txtInstructorS.Enabled = False
        cmbInstructorS.Enabled = False
        txtempresaS.Text = ""
        txtempresaS.Enabled = False
        cmbempresaS.Enabled = False
        cmbestadoS.Enabled = False
        'cmbjefeISA.Enabled = False
        txtcargoAS.Text = ""
        txtcargoAS.Enabled = False
        ' txtjefeISA.Text = ""
        ' txtjefeISA.Enabled = False
        cmbmodalidadS.Enabled = False
        'cmbmodalidadS.DataSource = Nothing
        'cmbmodalidadS.Items.Clear()
        'btnnuevogS.Text = "Nuevo"

    End Sub
    Private Sub btneditarS_Click(sender As Object, e As EventArgs) Handles btneditarS.Click
        If btneditarS.Text = "Habilitar Campos" Then
            txtInstructorS.Enabled = False
            cmbInstructorS.Enabled = False
            txtempresaS.Enabled = False
            cmbempresaS.Enabled = False
            cmbestadoS.Enabled = False
            txtcargoAS.Enabled = False
            cmbmodalidadS.Enabled = False
            btneditarS.Text = "Editar"
        Else
            Try
                If lblidseguimiento.Text = "" Then
                    MsgBox("Debe seleccionar un seguimiento para editarlo")
                Else
                    If lblidseguimiento.Text = "" Or IsDBNull(datagridAS.CurrentRow.Index.ToString) Or txtInstructorS.Text = "" Or txtempresaS.Text = "" Then
                        MsgBox("Existen campos vacíos")
                    Else
                        sql = "UPDATE seguimiento set "
                        sql += "[instructor]='" & cmbInstructorS.SelectedValue.ToString & "',"
                        sql += "[empresa]='" & cmbempresaS.SelectedValue.ToString & "',"
                        sql += "[cargoAprendiz]='" & txtcargoAS.Text & "',"
                        sql += "[fechaInicio]='" & datetimeInicioS.Value & "',"
                        sql += "[fechaFin]='" & datetimefinS.Value & "',"
                        sql += "[modalidad]='" & cmbmodalidadS.SelectedValue.ToString & "',"
                        sql += "[estado]='" & cmbestadoS.SelectedValue.ToString & "'"


                        sql += "where [idseguimiento]='" & lblidseguimiento.Text & "'"
                        Agregar()
                        LlenarDetalleSeguimientosAs()
                        LimpiarSeguimientoA()
                    End If


                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If

        
    End Sub
    Private Sub detallesSAS_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles detallesSAS.CellClick
        Try
            Dim fila As Integer
            fila = detallesSAS.CurrentRow.Index.ToString


            ' txtaprendiz.Text = datagridseguimiento.Rows(fila).Cells(0).Value.ToString
            txtempresaS.Text = detallesSAS.Rows(fila).Cells(2).Value.ToString
            cmbempresaS.Text = detallesSAS.Rows(fila).Cells(2).Value.ToString
            If detallesSAS.Rows(fila).Cells(3).Value.Equals(DBNull.Value) Then
            Else
                datetimeInicioS.Value = detallesSAS.Rows(fila).Cells(3).Value.ToString
            End If
            If detallesSAS.Rows(fila).Cells(4).Value.Equals(DBNull.Value) Then

            Else
                datetimefinS.Value = detallesSAS.Rows(fila).Cells(4).Value.ToString
            End If
            txtInstructorS.Text = detallesSAS.Rows(fila).Cells(5).Value.ToString
            cmbInstructorS.Text = detallesSAS.Rows(fila).Cells(5).Value.ToString
            txtcargoAS.Text = detallesSAS.Rows(fila).Cells(6).Value.ToString
            lblidseguimiento.Text = detallesSAS.Rows(fila).Cells(7).Value.ToString
            cmbmodalidadS.Text = detallesSAS.Rows(fila).Cells(8).Value.ToString
            cmbestadoS.Text = detallesSAS.Rows(fila).Cells(9).Value.ToString
            'txtjefeISA.Text = datagridseguimiento.Rows(fila).Cells(8).Value.ToString
            'cmbjefeISA.Text = datagridseguimiento.Rows(fila).Cells(8).Value.ToString


        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub btnbuscarAN_Click(sender As Object, e As EventArgs) Handles btnbuscarAN.Click
        If txtnombreaprendiz.Text = "" Then
            sql = "select TOP 80 [documento],[nombre],[apellido],[telefono],correo,[ficha],[Estado],t1.municipio, [Tipo_documento],direccion from aprendiz t1"
            datagridaprendices.DataSource = ListarDatos()
            Agregarcolor()
            txtnombreaprendiz.Enabled = True
        Else
            ' sql = "select TOP 50 [documento],[nombre],[apellido],[telefono],correo,[ficha],[Estado],t2.Municipio,t1.municipio, [Tipo_documento],direccion,t3.nombreDepartamento from aprendiz t1 inner join municipios t2 on t1.municipio=t2.Id inner join departamento t3 on t2.departamento=t3.iddepartamento where [documento] like '%" & txtdocumento.Text & "%'"
            sql = "select TOP 80 [documento],[nombre],[apellido],[telefono],correo,[ficha],[Estado],t1.municipio, [Tipo_documento],direccion from aprendiz t1  where [nombre] like '%" & txtnombreaprendiz.Text & "%'"
            datagridaprendices.DataSource = ListarDatos()
            datagridaprendices.Columns(7).Visible = False
            Agregarcolor()

        End If
        cerrar_conexion()
    End Sub
    Private Sub rbtsinseguimiento_CheckedChanged(sender As Object, e As EventArgs)

        If rbtsinseguimiento.Checked Then
            sql = "SELECT  TOP 80 [documento],[nombre],[apellido],[telefono],correo,[ficha],t1.Estado,t1.municipio, [Tipo_documento],direccion FROM aprendiz t1 WHERE NOT EXISTS (SELECT NULL FROM seguimiento t2 WHERE t2.aprendiz = t1.documento)"
            datagridaprendices.DataSource = ListarDatos()
            datagridaprendices.Columns(7).Visible = False
        Else

        End If

    End Sub
    Private Sub rbtconseguimiento_CheckedChanged(sender As Object, e As EventArgs)
        If rbtconseguimiento.Checked Then
            sql = " select TOP 80 [documento],[nombre],[apellido],[telefono],correo,[ficha],t1.Estado,t1.municipio, [Tipo_documento],direccion from aprendiz t1 right join seguimiento t2 on t1.documento=t2.aprendiz "
            datagridaprendices.DataSource = ListarDatos()
            datagridaprendices.Columns(7).Visible = False
        Else

        End If
    End Sub
    Private Sub rbttodos_CheckedChanged_1(sender As Object, e As EventArgs) Handles rbttodos.CheckedChanged
        If rbttodos.Checked Then
            sql = "select TOP 80 [documento],[nombre],[apellido],[telefono],correo,[ficha],[Estado],t1.municipio, [Tipo_documento],direccion from aprendiz t1"
            datagridaprendices.DataSource = ListarDatos()
            datagridaprendices.Columns(7).Visible = False
            datagridaprendices.ClearSelection()
            datagridaprendices.CurrentCell = Nothing
            Agregarcolor()
        End If

    End Sub
    Private Sub rbtsinseguimiento_CheckedChanged_1(sender As Object, e As EventArgs) Handles rbtsinseguimiento.CheckedChanged
        If rbtsinseguimiento.Checked Then
            sql = "SELECT TOP 80 [documento],[nombre],[apellido],[telefono],correo,[ficha],t1.Estado,t1.municipio, [Tipo_documento],direccion  FROM aprendiz t1 WHERE NOT EXISTS (SELECT NULL FROM seguimiento t2 WHERE t2.aprendiz = t1.documento)"
            datagridaprendices.DataSource = ListarDatos()
            datagridaprendices.Columns(7).Visible = False
            datagridaprendices.ClearSelection()
            datagridaprendices.CurrentCell = Nothing
            Agregarcolor()

        End If

    End Sub
    Private Sub rbtconseguimiento_CheckedChanged_1(sender As Object, e As EventArgs) Handles rbtconseguimiento.CheckedChanged
        If rbtconseguimiento.Checked Then
            sql = "select TOP 80 [documento],[nombre],[apellido],[telefono],correo,[ficha],t1.Estado,t1.municipio, [Tipo_documento],direccion from aprendiz t1 right join seguimiento t2 on t2.aprendiz=t1.documento"
            datagridaprendices.DataSource = ListarDatos()
            datagridaprendices.Columns(7).Visible = False
            Agregarcolor()
            datagridaprendices.ClearSelection()
            datagridaprendices.CurrentCell = Nothing
        End If

    End Sub
    Sub SubExcelReport()
        Try
            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader

            XLApp = CreateObject("Excel.application")
            XLBook = XLApp.Workbooks.Open(My.Computer.FileSystem.CurrentDirectory & "\formatoSeguimiento.xls")
            XLSheet = XLBook.Worksheets(1)
            XLApp.Visible = True

            Dim i As Integer
            While reader.Read

                XLSheet.Range(i + 12 & ":" & i + 12).EntireRow.Copy()
                XLSheet.Range(i + 12 & ":" & i + 12).EntireRow.Insert()

                XLSheet.Range("B" & i + 12).Value = reader("Tipo_documento")
                XLSheet.Cells(i + 12, 3).value = reader("documento")
                XLSheet.Cells(i + 12, 4).value = reader("nombre")
                XLSheet.Cells(i + 12, 5).value = reader("telefonoA")
                XLSheet.Cells(i + 12, 6).value = reader("correo")
                XLSheet.Cells(i + 12, 7).value = reader("ficha")
                XLSheet.Cells(i + 12, 8).value = reader("estadoa")
                XLSheet.Cells(i + 12, 9).value = reader("DEPA")
                XLSheet.Cells(i + 12, 10).value = reader("muniA")
                XLSheet.Cells(i + 12, 11).value = reader("cargoAprendiz")
                XLSheet.Cells(i + 12, 12).value = reader("EstadoE")
                XLSheet.Cells(i + 12, 13).value = reader("fechaInicio")
                XLSheet.Cells(i + 12, 14).value = reader("fechaFin")
                XLSheet.Cells(i + 12, 15).value = reader("nombreModalidad")
                XLSheet.Cells(i + 12, 16).value = reader("instructor")
                XLSheet.Cells(i + 12, 17).value = reader("cedula_NIT")
                XLSheet.Cells(i + 12, 18).value = reader("razonSocial")
                XLSheet.Cells(i + 12, 19).value = reader("TelefonoEmpresa")
                XLSheet.Cells(i + 12, 20).value = reader("DireccionE")
                XLSheet.Cells(i + 12, 21).value = reader("DptE")
                XLSheet.Cells(i + 12, 22).value = reader("MuniE")
                'i = i + 1
            End While
            reader.Close()
        Catch ex As Exception
            reader.Close()
            MsgBox(ex.ToString)
        End Try
        
    End Sub
    Private Sub btnreporte_Click(sender As Object, e As EventArgs) Handles btnreporte.Click
        Try
            If cmbestadoEvaluacion.Text = "TODOS" Then
                sql = "select [documento],concat([nombre],' ',[apellido]) as nombre,t1.telefono AS telefonoA,correo,[ficha],t1.Estado as estadoa,t2.Municipio as muniA,t1.municipio, [Tipo_documento],t1.direccion,t3.nombreDepartamento as DEPA,t4.cargoAprendiz,t4.fechaInicio,t4.fechaFin,t6.nombreModalidad,t4.instructor, t5.cedula_NIT,t5.razonSocial,t5.telefono as TelefonoEmpresa,t5.direccion as DireccionE,t7.Municipio as MuniE, t8.nombreDepartamento as DptE,t9.nombreEstado as EstadoE from aprendiz t1 inner join municipios t2 on t1.municipio=t2.Id  inner join departamento t3 on t2.departamento=t3.iddepartamento inner join seguimiento t4 on t1.documento=t4.aprendiz inner join empresa t5 on t4.empresa=t5.cedula_NIT inner join modalidad t6 on t6.idmodalida=t4.modalidad inner join municipios t7 on t5.municipio=t7.Id inner join departamento t8 on t7.departamento=t8.iddepartamento inner join estado t9 on t9.idestado=t4.estado"
                SubExcelReport()
                cerrar_conexion()
            Else
                sql = "select [documento],concat([nombre],' ',[apellido]) as nombre,t1.telefono AS telefonoA,correo,[ficha],t1.Estado as estadoa,t2.Municipio as muniA,t1.municipio, [Tipo_documento],t1.direccion,t3.nombreDepartamento as DEPA,t4.cargoAprendiz,t4.fechaInicio,t4.fechaFin,t6.nombreModalidad,t4.instructor, t5.cedula_NIT,t5.razonSocial,t5.telefono as TelefonoEmpresa,t5.direccion as DireccionE,t7.Municipio as MuniE, t8.nombreDepartamento as DptE,t9.nombreEstado as EstadoE from aprendiz t1 inner join municipios t2 on t1.municipio=t2.Id  inner join departamento t3 on t2.departamento=t3.iddepartamento inner join seguimiento t4 on t1.documento=t4.aprendiz inner join empresa t5 on t4.empresa=t5.cedula_NIT inner join modalidad t6 on t6.idmodalida=t4.modalidad inner join municipios t7 on t5.municipio=t7.Id inner join departamento t8 on t7.departamento=t8.iddepartamento inner join estado t9 on t9.idestado=t4.estado where t9.nombreEstado='" & cmbestadoEvaluacion.Text & "'"
                SubExcelReport()
                cerrar_conexion()
            End If
            
            XLApp.Application.DisplayAlerts = False
            XLBook.SaveAs(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ReporteSeguimientos\seguimiento" & Now.Day & Now.Month & Now.Year & Now.Second & ".xls")
            cerrar_conexion()
        Catch ex As Exception
            reader.Close()
            MsgBox(ex.ToString)
        End Try
        reader.Close()

    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs)
        cmbinstructorSA.Enabled = True
        datetimeinicio.Enabled = True
        datetimefin.Enabled = True
        cmbmodalidadSA.Enabled = True
        lblestatus_user.Enabled = True
        cmbestadoSA.Enabled = True
    End Sub
    Private Sub rbtodos_CheckedChanged(sender As Object, e As EventArgs) Handles rbtodos.CheckedChanged
        If rbtodos.Checked Then
            sql = "select TOP 80  CONCAT(t1.nombre,' ',t1.apellido) as nombre,documento,t1.ficha,t3.Nombre_curso,t1.Estado,t1.municipio from  aprendiz t1 inner join grupos t3 on t1.ficha=t3.ficha "
            datagridAS.DataSource = ListarDatos()
            datagridAS.Columns(4).Visible = False
            datagridAS.Columns(5).Visible = False
            datagridAS.ClearSelection()
            datagridAS.CurrentCell = Nothing
            AgregarcolorS()
        Else
            Exit Sub
        End If

    End Sub
    Private Sub rbsinseguimiento_CheckedChanged(sender As Object, e As EventArgs) Handles rbsinseguimiento.CheckedChanged
        If rbsinseguimiento.Checked Then
            sql = "select TOP 80 CONCAT(t1.nombre,' ',t1.apellido) as nombre,documento,t1.ficha,t3.Nombre_curso,t1.Estado,t1.municipio from aprendiz t1 inner join grupos t3 on t1.ficha=t3.ficha WHERE NOT EXISTS (SELECT NULL FROM seguimiento t2 WHERE t2.aprendiz = t1.documento)"
            datagridAS.DataSource = ListarDatos()
            datagridAS.Columns(4).Visible = False
            datagridAS.Columns(5).Visible = False
            AgregarcolorS()
            datagridAS.ClearSelection()
            datagridAS.CurrentCell = Nothing
        Else
            Exit Sub
        End If
    End Sub
    Private Sub rbconseguimiento_CheckedChanged(sender As Object, e As EventArgs) Handles rbconseguimiento.CheckedChanged
        If rbconseguimiento.Checked Then
            sql = "select TOP 80  CONCAT(t1.nombre,' ',t1.apellido) as nombre,documento,t1.ficha,t3.Nombre_curso,t1.Estado,t1.municipio from aprendiz t1 right join seguimiento t2 on t2.aprendiz=t1.documento inner join grupos t3 on t1.ficha=t3.ficha"
            datagridAS.DataSource = ListarDatos()
            datagridAS.Columns(4).Visible = False
            datagridAS.Columns(5).Visible = False
            AgregarcolorS()
            datagridAS.ClearSelection()
            datagridAS.CurrentCell = Nothing
        Else
            Exit Sub
        End If
    End Sub
    Sub AgregarcolorS()
        Dim i As Integer
        For i = 0 To datagridAS.Rows.Count - 1

            If datagridAS.Rows(i).Cells(4).Value.ToString = "CANCELADO" Or datagridAS.Rows(i).Cells(4).Value.ToString = "CONDICIONADO" Or datagridAS.Rows(i).Cells(4).Value.ToString = "TRASLADADO" Or datagridAS.Rows(i).Cells(4).Value.ToString = "RETIRO VOLUNTARIO" Then
                datagridAS.Rows(i).DefaultCellStyle.BackColor = Color.Red
            ElseIf (datagridAS.Rows(i).Cells(4).Value.ToString = "POR CERTIFICAR") Then
                datagridAS.Rows(i).DefaultCellStyle.BackColor = Color.YellowGreen

            ElseIf (datagridAS.Rows(i).Cells(4).Value.ToString = "CERTIFICADO") Then
                datagridAS.Rows(i).DefaultCellStyle.BackColor = Color.Green
            ElseIf (datagridAS.Rows(i).Cells(4).Value.ToString = "") Then
                datagridAS.Rows(i).DefaultCellStyle.BackColor = Color.Yellow

            End If
        Next



    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs)
        txtdocumento.Text = ""
        txtdocumento.Enabled = True
        txtnombreaprendiz.Text = ""
        txtnombreaprendiz.Enabled = True
        txtapellidoaprendiz.Text = ""
        txtapellidoaprendiz.Enabled = True
        txtemailaprendiz.Text = ""
        txtemailaprendiz.Enabled = True
        txttelefonoaprendiz.Text = ""
        cmbestadoformacion.Enabled = True
        txtficha.Text = ""
        txtficha.Enabled = True
        txtdireccionaprendiz.Text = ""
        txtdireccionaprendiz.Enabled = True
        txttelefonoaprendiz.Enabled = True
        txtemailaprendiz.Enabled = True
        cmbmunicipioaprendiz.Text = ""
        cmbmunicipioaprendiz.DataSource = Nothing
        cmbmunicipioaprendiz.Items.Clear()
        cmbdepartamentoaprendiz.Enabled = True
        cmbtipodoc.Enabled = True
        txtprogramaA.Text = ""
        cmbmunicipioaprendiz.Enabled = False

    End Sub
    Private Sub Button13_Click(sender As Object, e As EventArgs)
        txtinstructorSA.Text = ""
        txtinstructorSA.Enabled = True
        cmbinstructorSA.Enabled = True

        txtempresaSA.Text = ""
        txtempresaSA.Enabled = True
        cmbempresaSA.Enabled = True
        cmbestadoSA.Enabled = True
        cmbmodalidadSA.Enabled = True
        'cmbjefeISA.Enabled = True
        txtcargoA.Enabled = True
        txtcargoA.Text = ""
        ' txtjefeISA.Text = ""
        cmbmodalidadS.Enabled = True
    End Sub
    Private Sub btndetalleS_Click(sender As Object, e As EventArgs)

    End Sub
    Private Sub Button14_Click(sender As Object, e As EventArgs)
        txtInstructorS.Text = ""
        txtInstructorS.Enabled = True
        cmbInstructorS.Enabled = True
        txtempresaS.Text = ""
        txtempresaS.Enabled = True
        cmbempresaS.Enabled = True
        cmbestadoS.Enabled = True
        'cmbjefeISA.Enabled = True
        txtcargoAS.Enabled = True
        txtcargoAS.Text = ""
        ' txtjefeISA.Text = ""
        cmbmodalidadS.Enabled = True
    End Sub
    Private Sub btnbuscarAA_Click(sender As Object, e As EventArgs) Handles btnbuscarAA.Click
        If txtapellidoaprendiz.Text = "" Then
            sql = "select TOP 80 [documento],[nombre],[apellido],[telefono],correo,[ficha],[Estado],t1.municipio, [Tipo_documento],direccion from aprendiz t1"
            datagridaprendices.DataSource = ListarDatos()
            Agregarcolor()
            txtapellidoaprendiz.Enabled = True
        Else
            ' sql = "select TOP 50 [documento],[nombre],[apellido],[telefono],correo,[ficha],[Estado],t2.Municipio,t1.municipio, [Tipo_documento],direccion,t3.nombreDepartamento from aprendiz t1 inner join municipios t2 on t1.municipio=t2.Id inner join departamento t3 on t2.departamento=t3.iddepartamento where [documento] like '%" & txtdocumento.Text & "%'"
            sql = "select TOP 80 [documento],[nombre],[apellido],[telefono],correo,[ficha],[Estado],t1.municipio, [Tipo_documento],direccion from aprendiz t1  where [apellido] like '%" & txtapellidoaprendiz.Text & "%'"
            datagridaprendices.DataSource = ListarDatos()
            datagridaprendices.Columns(7).Visible = False
            Agregarcolor()

        End If
        cerrar_conexion()
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If txtnit.Text = "" Then
            txtnit.Enabled = True
            sql = "SELECT TOP 80 t1.*,t2.Municipio,t3.nombreDepartamento FROM empresa t1 inner join municipios t2 on t1.municipio=t2.Id inner join departamento t3 on t2.departamento=t3.iddepartamento"
            datagridEmpresa.DataSource = ListarDatos()
            datagridEmpresa.Columns(5).Visible = False
            datagridEmpresa.Columns(6).Visible = False
            cerrar_conexion()
        Else
            sql = "SELECT  TOP 80 t1.*,t2.Municipio,t3.nombreDepartamento FROM empresa t1 inner join municipios t2 on t1.municipio=t2.Id inner join departamento t3 on t2.departamento=t3.iddepartamento where cedula_NIT like '%" & txtnit.Text & "%'"
            datagridEmpresa.DataSource = ListarDatos()
            datagridEmpresa.Columns(5).Visible = False
            datagridEmpresa.Columns(6).Visible = False
            cerrar_conexion()
        End If

    End Sub
    Private Sub Button15_Click(sender As Object, e As EventArgs)

    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If txtrazonS.Text = "" Then
            sql = "SELECT top 80 t1.*,t2.Municipio,t3.nombreDepartamento FROM empresa t1 inner join municipios t2 on t1.municipio=t2.Id inner join departamento t3 on t2.departamento=t3.iddepartamento"
            datagridEmpresa.DataSource = ListarDatos()
            datagridEmpresa.Columns(5).Visible = False
            datagridEmpresa.Columns(6).Visible = False
            cerrar_conexion()
        Else
            sql = "SELECT top 80 t1.*,t2.Municipio,t3.nombreDepartamento FROM empresa t1 inner join municipios t2 on t1.municipio=t2.Id inner join departamento t3 on t2.departamento=t3.iddepartamento where [razonSocial] like '%" & txtrazonS.Text & "%'"
            datagridEmpresa.DataSource = ListarDatos()
            datagridEmpresa.Columns(5).Visible = False
            datagridEmpresa.Columns(6).Visible = False
            cerrar_conexion()
        End If

    End Sub
    Private Sub Button16_Click(sender As Object, e As EventArgs)

    End Sub
    Private Sub Button6_Click_1(sender As Object, e As EventArgs)
        MsgBox(String.Format("{0:HH:mm:ss}", DateTime.Now))
    End Sub
    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        txtinstructorSA.Enabled = True
        If cmbinstructorSA.Text = "" Then
            ' sql = "select TOP 50 [documento],[nombre],[apellido],[telefono],correo,[ficha],[Estado],t2.Municipio,t1.municipio, [Tipo_documento],direccion,t3.nombreDepartamento from aprendiz t1 inner join municipios t2 on t1.municipio=t2.Id inner join departamento t3 on t2.departamento=t3.iddepartamento"
            sql = "select TOP 80 CONCAT(t2.nombre,' ',t2.apellido)as aprendiz, t3.razonSocial,fechaInicio,fechaFin,instructor,cargoAprendiz,t1.idseguimiento,t5.nombreModalidad,t2.documento,t6.nombreEstado from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT inner join modalidad t5 on t1.modalidad=t5.idmodalida inner join estado t6 on t1.Estado=t6.idestado"
            datagridseguimiento.DataSource = ListarDatos()
            datagridseguimiento.Columns(4).Visible = False
            datagridseguimiento.Columns(5).Visible = False
            datagridseguimiento.Columns(6).Visible = False
            datagridseguimiento.Columns(8).Visible = False

        Else

            sql = "select TOP 80 CONCAT(t2.nombre,' ',t2.apellido)as aprendiz, t3.razonSocial,fechaInicio,fechaFin,instructor,cargoAprendiz,t1.idseguimiento,t5.nombreModalidad,t2.documento,t6.nombreEstado from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT inner join modalidad t5 on t1.modalidad=t5.idmodalida inner join estado t6 on t1.Estado=t6.idestado where instructor='" & cmbinstructorSA.Text & "'"
            datagridseguimiento.DataSource = ListarDatos()
            datagridseguimiento.Columns(4).Visible = False
            datagridseguimiento.Columns(5).Visible = False
            datagridseguimiento.Columns(6).Visible = False
            datagridseguimiento.Columns(8).Visible = False

        End If
        cerrar_conexion()
    End Sub
    Private Sub rbtodoss_CheckedChanged(sender As Object, e As EventArgs) Handles rbtodoss.CheckedChanged
        sql = "select TOP 80 CONCAT(t2.nombre,' ',t2.apellido)as aprendiz, t3.razonSocial,fechaInicio,fechaFin,instructor,cargoAprendiz,t1.idseguimiento,t5.nombreModalidad,t2.documento,t6.nombreEstado from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT inner join modalidad t5 on t1.modalidad=t5.idmodalida inner join estado t6 on t1.Estado=t6.idestado "
        datagridseguimiento.DataSource = ListarDatos()
        datagridseguimiento.Columns(4).Visible = False
        datagridseguimiento.Columns(5).Visible = False
        datagridseguimiento.Columns(6).Visible = False
        datagridseguimiento.Columns(8).Visible = False
        cerrar_conexion()
    End Sub
    Private Sub rbporevaluar_CheckedChanged(sender As Object, e As EventArgs) Handles rbporevaluar.CheckedChanged
        sql = "select TOP 80 CONCAT(t2.nombre,' ',t2.apellido)as aprendiz, t3.razonSocial,fechaInicio,fechaFin,instructor,cargoAprendiz,t1.idseguimiento,t5.nombreModalidad,t2.documento,t6.nombreEstado from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT inner join modalidad t5 on t1.modalidad=t5.idmodalida inner join estado t6 on t1.Estado=t6.idestado  where t1.Estado='3'"
        datagridseguimiento.DataSource = ListarDatos()
        datagridseguimiento.Columns(4).Visible = False
        datagridseguimiento.Columns(5).Visible = False
        datagridseguimiento.Columns(6).Visible = False
        datagridseguimiento.Columns(8).Visible = False
        cerrar_conexion()
    End Sub
    Private Sub rbevaluado_CheckedChanged(sender As Object, e As EventArgs) Handles rbevaluado.CheckedChanged
        sql = "select TOP 80 CONCAT(t2.nombre,' ',t2.apellido)as aprendiz, t3.razonSocial,fechaInicio,fechaFin,instructor,cargoAprendiz,t1.idseguimiento,t5.nombreModalidad,t2.documento,t6.nombreEstado from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT inner join modalidad t5 on t1.modalidad=t5.idmodalida inner join estado t6 on t1.Estado=t6.idestado where t1.Estado='1'"
        datagridseguimiento.DataSource = ListarDatos()
        datagridseguimiento.Columns(4).Visible = False
        datagridseguimiento.Columns(5).Visible = False
        datagridseguimiento.Columns(6).Visible = False
        datagridseguimiento.Columns(8).Visible = False
        cerrar_conexion()
    End Sub
    Private Sub rbnoaprobado_CheckedChanged(sender As Object, e As EventArgs) Handles rbnoaprobado.CheckedChanged
        sql = "select TOP 80 CONCAT(t2.nombre,' ',t2.apellido)as aprendiz, t3.razonSocial,fechaInicio,fechaFin,instructor,cargoAprendiz,t1.idseguimiento,t5.nombreModalidad,t2.documento,t6.nombreEstado from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT inner join modalidad t5 on t1.modalidad=t5.idmodalida inner join estado t6 on t1.Estado=t6.idestado where t1.Estado='2'"
        datagridseguimiento.DataSource = ListarDatos()
        datagridseguimiento.Columns(4).Visible = False
        datagridseguimiento.Columns(5).Visible = False
        datagridseguimiento.Columns(6).Visible = False
        datagridseguimiento.Columns(8).Visible = False
        cerrar_conexion()
    End Sub
    Private Sub Button6_Click_2(sender As Object, e As EventArgs) Handles Button6.Click
        txtdocumento.Text = ""
        txtdocumento.Enabled = False
        txtnombreaprendiz.Text = ""
        txtnombreaprendiz.Enabled = False
        txtapellidoaprendiz.Text = ""
        txtapellidoaprendiz.Enabled = False
        txtemailaprendiz.Text = ""
        txtemailaprendiz.Enabled = False
        txttelefonoaprendiz.Text = ""
        cmbestadoformacion.Enabled = False
        txtficha.Text = ""
        txtficha.Enabled = False
        txtdireccionaprendiz.Text = ""
        txtdireccionaprendiz.Enabled = False
        txttelefonoaprendiz.Enabled = False
        txtemailaprendiz.Enabled = False
        cmbmunicipioaprendiz.Text = ""
        cmbmunicipioaprendiz.DataSource = Nothing
        cmbmunicipioaprendiz.Items.Clear()
        cmbdepartamentoaprendiz.Enabled = False
        cmbtipodoc.Enabled = False
        txtprogramaA.Text = ""
        cmbmunicipioaprendiz.Enabled = False
        btnagregarNA.Text = "Nuevo"
        btneditarA.Text = "Habilitar Campos"
    End Sub
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        txtinstructorSA.Text = ""
        txtinstructorSA.Enabled = False
        cmbinstructorSA.Enabled = False

        txtempresaSA.Text = ""
        txtempresaSA.Enabled = False
        cmbempresaSA.Enabled = False
        cmbestadoSA.Enabled = False
        cmbmodalidadSA.Enabled = False
        'cmbjefeISA.Enabled = True
        txtcargoA.Enabled = False
        txtcargoA.Text = ""
        ' txtjefeISA.Text = ""
        cmbmodalidadS.Enabled = False
        btnnuevoSeguimiennto.Text = "Nuevo"
        btneditarSeguimiento.Text = "Habilitar Campos"
    End Sub
    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        txtInstructorS.Text = ""
        txtInstructorS.Enabled = False
        cmbInstructorS.Enabled = False
        txtempresaS.Text = ""
        txtempresaS.Enabled = False
        cmbempresaS.Enabled = False
        cmbestadoS.Enabled = False
        'cmbjefeISA.Enabled = False
        txtcargoAS.Text = ""
        txtcargoAS.Enabled = False
        ' txtjefeISA.Text = ""
        ' txtjefeISA.Enabled = False
        cmbmodalidadS.Enabled = False
            btnnuevogS.Text = "Nuevo"
            btneditarS.Text = "Habilitar Campos"

    End Sub
    Private Sub Button13_Click_1(sender As Object, e As EventArgs) Handles Button13.Click
        txtnit.Text = ""
        txtrazonS.Text = ""
        txtdireccionE.Text = ""
        txttelefonoE.Text = ""
        cmbmunicipioE.Enabled = False
        txtdireccionE.Enabled = False
        txttelefonoE.Enabled = False
        txtnit.Enabled = False
        txtrazonS.Enabled = False
        cmbmunicipioE.DataSource = Nothing
        cmbmunicipioE.Items.Clear()
       
            btnnuevogEmpresa.Text = "Nuevo"

            btneditarEmpresa.Text = "Habilitar Campos"


    End Sub
    Private Sub Button14_Click_1(sender As Object, e As EventArgs) Handles Button14.Click
        txtnombreJI.Text = ""
        txtcargoJ.Text = ""
        txttelefonoJ.Text = ""
        txtcorreoJ.Text = ""
        txtempresaE.Text = ""
        txtnombreJI.Enabled = False
        txtcargoJ.Enabled = False
        txttelefonoJ.Enabled = False
        txtcorreoJ.Enabled = False
        txtempresaE.Enabled = False
        cmbempresaE.Enabled = False
        cmbempresaE.DataSource = Nothing
        cmbempresaE.Items.Clear()
        cmbempresaE.Enabled = False
      
            btneditarE.Text = "Habilitar Campos"
        
        
            btnnuevojE.Text = "Nuevo"

    End Sub

    Private Sub rbevaluados_CheckedChanged(sender As Object, e As EventArgs) Handles rbevaluados.CheckedChanged
        If rbevaluados.Checked Then
            sql = "select    ROW_NUMBER() OVER(ORDER BY idseguimiento ASC) AS 'No', t2.documento as 'Documento',CONCAT(t2.nombre,' ',t2.apellido)as 'Nombre Aprendiz',t2.telefono as 'Telefono Aprendiz',t2.correo AS 'Correo Aprendiz',t8.ficha as 'Ficha',t8.Nombre_curso as 'Nombre Curso', t3.razonSocial as 'Empresa',t5.nombreModalidad as 'Modalidad',fechaInicio as 'Fecha de Inicio',fechaFin as 'Fecha Fin',fechaAsignacion as 'Fecha Asignación',instructor as 'Instructor',t9.Correo as 'Correo Instructor',t9.telefono as 'Telefono Instructor',cargoAprendiz as 'Cargo del Aprendiz',t6.nombreEstado as 'Estado' from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT inner join modalidad t5 on t1.modalidad=t5.idmodalida inner join estado t6 on t1.Estado=t6.idestado inner join grupos t8 on t2.ficha=t8.ficha inner join instructores t9 on t1.instructor=t9.NOMBRE_FUNCIONARIO where nombreEstado='APROBADO' "
            dtinspeccionA.DataSource = ListarDatos()
            dtinspeccionA.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dtinspeccionA.Columns(0).Width = 30
            dtinspeccionA.Columns(1).Width = 76
            dtinspeccionA.Columns(2).Width = 170
            dtinspeccionA.Columns(3).Width = 79
            dtinspeccionA.Columns(4).Width = 150
            dtinspeccionA.Columns(5).Width = 58
            dtinspeccionA.Columns(6).Width = 105
            dtinspeccionA.Columns(7).Width = 150
            dtinspeccionA.Columns(8).Width = 75
            dtinspeccionA.Columns(9).Width = 75
            dtinspeccionA.Columns(10).Width = 75
            dtinspeccionA.Columns(11).Width = 72
            dtinspeccionA.Columns(12).Width = 220
            dtinspeccionA.Columns(13).Width = 150
            AgregarcolorInspeccion()
            dtinspeccionA.ClearSelection()
            dtinspeccionA.CurrentCell = Nothing
        Else
            Exit Sub
        End If

    End Sub

    Private Sub rbtporevaluar_CheckedChanged(sender As Object, e As EventArgs) Handles rbtporevaluar.CheckedChanged
        If rbtporevaluar.Checked Then
            sql = "select    ROW_NUMBER() OVER(ORDER BY idseguimiento ASC) AS 'No', t2.documento as 'Documento',CONCAT(t2.nombre,' ',t2.apellido)as 'Nombre Aprendiz',t2.telefono as 'Telefono Aprendiz',t2.correo AS 'Correo Aprendiz',t8.ficha as 'Ficha',t8.Nombre_curso as 'Nombre Curso', t3.razonSocial as 'Empresa',t5.nombreModalidad as 'Modalidad',fechaInicio as 'Fecha de Inicio',fechaFin as 'Fecha Fin',fechaAsignacion as 'Fecha Asignación',instructor as 'Instructor',t9.Correo as 'Correo Instructor',t9.telefono as 'Telefono Instructor',cargoAprendiz as 'Cargo del Aprendiz',t6.nombreEstado as 'Estado' from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT inner join modalidad t5 on t1.modalidad=t5.idmodalida inner join estado t6 on t1.Estado=t6.idestado inner join grupos t8 on t2.ficha=t8.ficha inner join instructores t9 on t1.instructor=t9.NOMBRE_FUNCIONARIO where nombreEstado='POR EVALUAR' "
            dtinspeccionA.DataSource = ListarDatos()
            dtinspeccionA.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dtinspeccionA.Columns(0).Width = 30
            dtinspeccionA.Columns(1).Width = 76
            dtinspeccionA.Columns(2).Width = 170
            dtinspeccionA.Columns(3).Width = 79
            dtinspeccionA.Columns(4).Width = 150
            dtinspeccionA.Columns(5).Width = 58
            dtinspeccionA.Columns(6).Width = 105
            dtinspeccionA.Columns(7).Width = 150
            dtinspeccionA.Columns(8).Width = 75
            dtinspeccionA.Columns(9).Width = 75
            dtinspeccionA.Columns(10).Width = 75
            dtinspeccionA.Columns(11).Width = 72
            dtinspeccionA.Columns(12).Width = 220
            dtinspeccionA.Columns(13).Width = 150
            AgregarcolorInspeccion()
            dtinspeccionA.ClearSelection()
            dtinspeccionA.CurrentCell = Nothing
        Else
            Exit Sub
        End If
    End Sub

    Private Sub rbtttodos_CheckedChanged(sender As Object, e As EventArgs) Handles rbtttodos.CheckedChanged
        If rbtttodos.Checked Then
            llenarAprendcesSeguimientoInspeccion()
        Else
            Exit Sub
        End If
    End Sub
    Sub AgregarcolorInspeccion()
        Dim i As Integer
        Dim l1, l2, l3, l4 As Integer

        Dim fechabdfin, fechaProx, fechacalculada, fechadia, fechaano, fechaano1 As String
        Dim fechafin, fechavalidacion, fechaproxima As Date

        Try
            For i = 0 To dtinspeccionA.Rows.Count - 1

                fechafin = dtinspeccionA.Rows(i).Cells(10).Value
                'Console.WriteLine(fechafin & " fecha fin")
                fechaProx = fechafin.Day

                fechacalculada = fechafin.Month - 1
                fechaano = fechafin.Year

                If fechacalculada = 0 Then
                    fechacalculada = 12
                    fechaano = fechafin.Year - 1
                End If

                If fechacalculada = -1 Then
                    fechacalculada = 11
                    fechaano = fechafin.Year - 1
                End If

                If fechacalculada = 2 And fechaProx = 31 Or fechacalculada = 2 And fechaProx = 30 Then
                    Dim daysInFeb As Integer = Date.DaysInMonth(Now.Year, "02")
                    fechaProx = daysInFeb
                End If
                If fechacalculada = 4 And fechaProx = 31 Or fechacalculada = 4 And fechaProx = 31 Then
                    Dim daysInFeb As Integer = Date.DaysInMonth(Now.Year, "04")
                    fechaProx = daysInFeb
                End If
                If fechacalculada = 6 And fechaProx = 31 Or fechacalculada = 6 And fechaProx = 31 Then
                    Dim daysInFeb As Integer = Date.DaysInMonth(Now.Year, "06")
                    fechaProx = daysInFeb
                End If
                If fechacalculada = 9 And fechaProx = 31 Or fechacalculada = 9 And fechaProx = 31 Then
                    Dim daysInFeb As Integer = Date.DaysInMonth(Now.Year, "09")
                    fechaProx = daysInFeb
                End If
                If fechacalculada = 11 And fechaProx = 31 Or fechacalculada = 11 And fechaProx = 31 Then
                    Dim daysInFeb As Integer = Date.DaysInMonth(Now.Year, "11")
                    fechaProx = daysInFeb
                End If

                fechavalidacion = fechaProx & "/" & fechacalculada & "/" & fechaano
                'Console.WriteLine(fechavalidacion & " fecha validacion")
                fechabdfin = fechafin.Month - 3
                fechaano1 = fechafin.Year
                If fechabdfin = 0 Then
                    fechabdfin = 12
                    fechaano1 = fechafin.Year - 1
                End If
                If fechabdfin = -1 Then
                    fechabdfin = 11
                    fechaano1 = fechafin.Year - 1
                End If
                If fechabdfin = -2 Then
                    fechabdfin = 10
                    fechaano1 = fechafin.Year - 1
                End If
                fechadia = fechafin.Day
                If fechabdfin = 2 And fechadia = 31 Or fechabdfin = 2 And fechadia = 30 Then
                    Dim daysInFeb As Integer = Date.DaysInMonth(Now.Year, "02")
                    fechadia = daysInFeb
                End If
                If fechabdfin = 4 And fechadia = 31 Or fechabdfin = 4 And fechadia = 31 Then
                    Dim daysInFeb As Integer = Date.DaysInMonth(Now.Year, "04")
                    fechadia = daysInFeb
                End If
                If fechabdfin = 6 And fechadia = 31 Or fechabdfin = 6 And fechadia = 31 Then
                    Dim daysInFeb As Integer = Date.DaysInMonth(Now.Year, "06")
                    fechadia = daysInFeb
                End If
                If fechabdfin = 9 And fechadia = 31 Or fechabdfin = 9 And fechadia = 31 Then
                    Dim daysInFeb As Integer = Date.DaysInMonth(Now.Year, "09")
                    fechadia = daysInFeb
                End If
                If fechabdfin = 11 And fechadia = 31 Or fechabdfin = 11 And fechadia = 31 Then
                    Dim daysInFeb As Integer = Date.DaysInMonth(Now.Year, "11")
                    fechadia = daysInFeb
                End If
                fechaproxima = fechadia & "/" & fechabdfin & "/" & fechaano1
                'Console.WriteLine(fechaproxima & " fecha proxima")
                If dtinspeccionA.Rows(i).Cells(16).Value.ToString = "POR EVALUAR" And fechafin <= Now.Date Then
                    dtinspeccionA.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    l1 = l1 + 1
                    lblsinevT.Text = l1
                ElseIf (dtinspeccionA.Rows(i).Cells(16).Value.ToString = "POR EVALUAR" And fechafin >= fechavalidacion And fechavalidacion >= Now.Date) Then
                    dtinspeccionA.Rows(i).DefaultCellStyle.BackColor = Color.MintCream
                    l2 = l2 + 1
                    L3eval.Text = l2
                ElseIf (dtinspeccionA.Rows(i).Cells(16).Value.ToString = "POR EVALUAR" And fechafin >= Now.Date And fechaproxima <= fechavalidacion) Then
                    dtinspeccionA.Rows(i).DefaultCellStyle.BackColor = Color.Yellow

                    l3 = l3 + 1
                    lblsinp.Text = l3
                ElseIf (dtinspeccionA.Rows(i).Cells(16).Value.ToString = "APROBADO") Then
                    dtinspeccionA.Rows(i).DefaultCellStyle.BackColor = Color.Green
                    l4 = l4 + 1
                    L4sinevalu.Text = l4
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

       
        'And Now.Date >= fechaproxima
    End Sub

    Private Sub TabPage5_Enter(sender As Object, e As EventArgs) Handles TabPage5.Enter
        llenarAprendcesSeguimientoInspeccion()

    End Sub

   
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            llenarAprendcesSeguimientoInspeccion()

        Else

            sql = "select    ROW_NUMBER() OVER(ORDER BY idseguimiento ASC) AS 'No', t2.documento as 'Documento',CONCAT(t2.nombre,' ',t2.apellido)as 'Nombre Aprendiz',t2.telefono as 'Telefono Aprendiz',t2.correo AS 'Correo Aprendiz',t8.ficha as 'Ficha',t8.Nombre_curso as 'Nombre Curso', t3.razonSocial as 'Empresa',t5.nombreModalidad as 'Modalidad',fechaInicio as 'Fecha de Inicio',fechaFin as 'Fecha Fin',fechaAsignacion as 'Fecha Asignación',instructor as 'Instructor',t9.Correo as 'Correo Instructor',t9.telefono as 'Telefono Instructor',cargoAprendiz as 'Cargo del Aprendiz',t6.nombreEstado as 'Estado' from seguimiento t1 inner join aprendiz t2 on t1.aprendiz=t2.documento inner join empresa t3 on t1.empresa=t3.cedula_NIT inner join modalidad t5 on t1.modalidad=t5.idmodalida inner join estado t6 on t1.Estado=t6.idestado inner join grupos t8 on t2.ficha=t8.ficha inner join instructores t9 on t1.instructor=t9.NOMBRE_FUNCIONARIO WHERE instructor LIKE '%" & TextBox1.Text & "%'"
            dtinspeccionA.DataSource = ListarDatos()
            dtinspeccionA.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dtinspeccionA.Columns(0).Width = 30
            dtinspeccionA.Columns(1).Width = 76
            dtinspeccionA.Columns(2).Width = 170
            dtinspeccionA.Columns(3).Width = 79
            dtinspeccionA.Columns(4).Width = 150
            dtinspeccionA.Columns(5).Width = 58
            dtinspeccionA.Columns(6).Width = 105
            dtinspeccionA.Columns(7).Width = 150
            dtinspeccionA.Columns(8).Width = 75
            dtinspeccionA.Columns(9).Width = 75
            dtinspeccionA.Columns(10).Width = 75
            dtinspeccionA.Columns(11).Width = 72
            dtinspeccionA.Columns(12).Width = 220
            dtinspeccionA.Columns(13).Width = 150
            AgregarcolorInspeccion()
            dtinspeccionA.ClearSelection()
            dtinspeccionA.CurrentCell = Nothing
        End If

       
    End Sub

    Private Sub dtinspeccionA_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dtinspeccionA.CellContentClick

    End Sub

    Private Sub Button15_Click_1(sender As Object, e As EventArgs) Handles Button15.Click
        Dim rpt As Integer = MessageBox.Show("¿Está seguro de que desea eliminar este seguimiento?" & vbNewLine & "Aprendiz: " & datagridseguimiento.CurrentRow.Cells(0).Value.ToString & vbNewLine & "Instructor: " & datagridseguimiento.CurrentRow.Cells(4).Value.ToString, "Advertencia", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
        If rpt = 6 Then
            Dim seguimiento As String = datagridseguimiento.CurrentRow.Cells(6).Value.ToString

            sql = "delete from  [seguimiento] where [idseguimiento]= '" & seguimiento & "'"
            Agregar()
            LlenarSeguimientosA()

        Else
            Exit Sub
        End If
    End Sub

   

    Private Sub Button16_Click_2(sender As Object, e As EventArgs) Handles Button16.Click
        XLApp = CreateObject("Excel.Application")
        XLBook = XLApp.Workbooks.Open(My.Computer.FileSystem.CurrentDirectory & "\Libro1.XLS")
        XLSheet = XLBook.Worksheets(1)
        XLSheet.Name = "REPORTE"
        XLApp.Visible = True
        Solocrealibro()
    End Sub
    Sub Solocrealibro()



        Dim contador As Integer = 12
        Dim dateaño As String = Now.Year - 3
        Dim fila, i As Integer
        Try
          
            For i = 0 To dtinspeccionA.Rows.Count - 1
                fila = contador



                XLSheet.Range(i + 12 & ":" & i + 12).EntireRow.Copy()
                XLSheet.Range(i + 12 & ":" & i + 12).EntireRow.Insert()

                'XLSheet.Range("B" & 12 & ":B" & 12).EntireRow.Copy()

                contador += 1

                XLSheet.Range("B" & fila).Value = dtinspeccionA.Rows(i).Cells(0).Value.ToString
                XLSheet.Range("C" & fila).Value = dtinspeccionA.Rows(i).Cells(1).Value.ToString
                XLSheet.Range("E" & fila).Value = dtinspeccionA.Rows(i).Cells(2).Value.ToString
                XLSheet.Range("G" & fila).Value = dtinspeccionA.Rows(i).Cells(3).Value.ToString
                XLSheet.Range("I" & fila).Value = dtinspeccionA.Rows(i).Cells(4).Value.ToString
                XLSheet.Range("K" & fila).Value = dtinspeccionA.Rows(i).Cells(5).Value.ToString
                XLSheet.Range("M" & fila).Value = dtinspeccionA.Rows(i).Cells(6).Value.ToString
                XLSheet.Range("N" & fila).Value = dtinspeccionA.Rows(i).Cells(7).Value.ToString
                XLSheet.Range("P" & fila).Value = dtinspeccionA.Rows(i).Cells(8).Value.ToString
                XLSheet.Range("R" & fila).Value = dtinspeccionA.Rows(i).Cells(9).Value
                XLSheet.Range("S" & fila).Value = dtinspeccionA.Rows(i).Cells(10).Value
                XLSheet.Range("T" & fila).Value = dtinspeccionA.Rows(i).Cells(11).Value
                XLSheet.Range("V" & fila).Value = dtinspeccionA.Rows(i).Cells(12).Value.ToString
                XLSheet.Range("Z" & fila).Value = dtinspeccionA.Rows(i).Cells(13).Value.ToString
                XLSheet.Range("AB" & fila).Value = dtinspeccionA.Rows(i).Cells(14).Value.ToString
                XLSheet.Range("AD" & fila).Value = dtinspeccionA.Rows(i).Cells(16).Value.ToString
            Next


            XLApp.Application.DisplayAlerts = False
            'XLBook.Visible = True
        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

  
   
    Private Sub dtinspeccionA_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dtinspeccionA.ColumnHeaderMouseClick
        AgregarcolorInspeccion()

        'Console.WriteLine("Registrado")

    End Sub

   
 
 
End Class