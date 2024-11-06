Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class fmrinputbox

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        Form3.Enabled = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If txterror.Text = "" Then
            MsgBox("Debe introducir el motivo del error")
        Else
            Try
                sql = "update dbo.programacion set estado_de_registro= 'Error', motivo_de_error= '" & txterror.Text & "' Where id = " & Form3.lblidcompetencia.Text & ""
                conectado()
                cmd = New SqlCommand(sql, cnn)
                cmd.ExecuteNonQuery()
                cerrar_conexion()

                MsgBox("Motivo de error, Registrado con Exito")

            Catch ex As Exception
                MsgBox(ex.ToString)

            End Try

            Me.Close()
            Form3.Enabled = True

        End If
    End Sub
End Class