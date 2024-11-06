Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class LoginForm2

    ' TODO: inserte el código para realizar autenticación personalizada usando el nombre de usuario y la contraseña proporcionada 
    ' (Consulte http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' El objeto principal personalizado se puede adjuntar al objeto principal del subproceso actual como se indica a continuación: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' donde CustomPrincipal es la implementación de IPrincipal utilizada para realizar la autenticación. 
    ' Posteriormente, My.User devolverá la información de identidad encapsulada en el objeto CustomPrincipal
    ' como el nombre de usuario, nombre para mostrar, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        If txtpass.Text = "" Or txtuser.Text = "" Or TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Debe introducir el usuario y contraseña")
        ElseIf TextBox1.Text <> TextBox2.Text Then
            MsgBox("La nueva contraseñas debe coincidir")
        Else

            Try
                sql = "select * from dbo.usuarios where id= '" & txtuser.Text & "'"
                conectado()
                cmd = New SqlCommand(sql, cnn)

                reader = cmd.ExecuteReader

                If reader.Read Then
                    If reader("pass") = txtpass.Text Then
                        cnn.Close()
                        reader.Close()

                        Try
                            sql = "update dbo.usuarios set "
                            sql += " pass= '" & TextBox1.Text & "'"

                            sql += " Where id = '" & txtuser.Text & "'"
                            conectado()

                           

                            cmd = New SqlCommand(sql, cnn)
                            cmd.ExecuteNonQuery()

                            MsgBox("Contraseña cambiada exitosamente")

                        Catch ex As Exception
                            MsgBox(ex.ToString)
                        End Try
                        Me.Close()
                    Else
                        MsgBox("Contraseña incorrecta")
                    End If
                Else
                    MsgBox("El usuario no existe")
                End If

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If

    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub LoginForm2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
