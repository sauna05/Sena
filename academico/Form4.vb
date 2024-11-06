Imports System.Data.SqlClient

Public Class FmrZona

    Private Sub FmrZona_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Cbzona.SelectedIndex = 0
        Form1.Enabled = False

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim rta As Integer
        rta = MsgBox("Seguro desea guardar la informacion de zona del instructor?", 4)
        If rta = 7 Then
        Else
            Try

                sql = "update dbo.instructores set Zona= '" & Cbzona.Text & "' Where NOMBRE_FUNCIONARIO = '" & Form1.ComboBox4.Text & "'"
                conectado() '***********************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
                cmd = New SqlCommand(sql, cnn)
                cmd.ExecuteNonQuery()
                cerrar_conexion()
                MsgBox(" Instructor Actualizado con exito, Actualice nuevamente la Programacion del instructor ")
                Form1.Enabled = True
                Me.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If
    End Sub
End Class