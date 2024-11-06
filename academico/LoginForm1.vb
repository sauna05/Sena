Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class LoginForm1

    ' TODO: inserte el código para realizar autenticación personalizada usando el nombre de usuario y la contraseña proporcionada 
    ' (Consulte http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' El objeto principal personalizado se puede adjuntar al objeto principal del subproceso actual como se indica a continuación: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' donde CustomPrincipal es la implementación de IPrincipal utilizada para realizar la autenticación. 
    ' Posteriormente, My.User devolverá la información de identidad encapsulada en el objeto CustomPrincipal
    ' como el nombre de usuario, nombre para mostrar, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        If txtpass.Text = "" Or txtuser.Text = "" Then
            MsgBox("Debe introducir el usuario y contraseña")
        Else

        End If
        Try
            sql = "select * from dbo.usuarios where id= '" & txtuser.Text & "'"
            conectado()
            cmd = New SqlCommand(sql, cnn)

            reader = cmd.ExecuteReader

            If reader.Read Then
                If reader("pass") = txtpass.Text Then

                    If reader("tipo de usuario") = "administrador" Then
                        MsgBox("Bienvenido " & reader("Nombre"))
                        Form1.lblusuario.Text = reader("Nombre")
                        Form1.lbltipousuario.Text = reader("cargo")
                        Form1.lbliduserLogin.Text = reader("Id")
                        Form1.lblcorreo.Text = reader("correo")
                        seguimiento_E_P.lbltipousuario.Text = reader("cargo")
                        reader.Close()
                        cerrar_conexion()
                        Form1.Show()
                    ElseIf reader("tipo de usuario") = "sofia" Then

                        MsgBox("Bienvenido " & reader("Nombre"))
                        Form3.lblusuariosofia.Text = reader("Nombre")
                        Form3.lblzona.Text = reader("Zona")
                        Form3.lbluser.Text = reader("Id")
                        reader.Close()
                        cerrar_conexion()
                        Form3.Show()
                    ElseIf reader("tipo de usuario") = "seguimiento" Then
                        seguimiento_E_P.Show()
                        seguimiento_E_P.lblestatus_user.Text = reader("id")
                        seguimiento_E_P.lblnombreusuario.Text = reader("Nombre")
                        seguimiento_E_P.lbltipousuario.Text = reader("cargo")
                    ElseIf (reader("tipo de usuario") = "inspector") Then

                        seguimiento_E_P.lbltipousuario.Text = reader("cargo")
                        seguimiento_E_P.Show()
                    End If
                    Me.Hide()
                Else
                    MsgBox("Contraseña incorrecta")
                End If
            Else
                MsgBox("El usuario no existe")
            End If
            
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

 
End Class
