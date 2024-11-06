
Imports System.Net.Mail.SmtpClient
Imports System.Net.Mail
Public Class Form2

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TabControl1.SelectTab(TabPage2)
    End Sub
    Public Property FileName() As String


    Private Sub btnadjuntar_Click(sender As Object, e As EventArgs) Handles btnadjuntar.Click
        abrir.ShowDialog()

        Dim nombrefile As String = abrir.SafeFileName
        Dim rutafile As String = abrir.FileName
        If lblnomb_archivo.Text = "" Then
            lblnomb_archivo.Text = nombrefile
        Else
            lblnomb_archivo.Text += " , " & nombrefile
        End If

        If abrir.FileName = "" Then
        Else
            message.Attachments.Add(New Attachment(rutafile))

            lblnomb_archivo.ForeColor = Color.Green
            lblnomb_archivo.Visible = True
        End If


    End Sub

    Private Sub cmbdescripcion_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbdescripcion.SelectedIndexChanged
        Select Case cmbdescripcion.Text

            Case "Outlook"
                txtservidor.Text = "smtp.live.com"

                txtpuerto.Text = "465"
                lbldefault.Text = "465"
                lbldefault.Visible = True
            Case "Gmail"
                txtservidor.Text = "smtp.gmail.com"
                txtpuerto.Text = "587"
                lbldefault.Text = "587 , 465"
                lbldefault.Visible = True
        End Select

        servidor_email = txtservidor.Text
        puerto = txtpuerto.Text
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles btnenviar.Click
        para = txtpara.Text
        desde = txtdesde.Text
        contasena = txtcontraseña.Text
        cuerpo = txtcuerpo_msg.Text
        asunto = txtasunto.Text
        enviar()

    End Sub
   

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

End Class