Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail.SmtpClient
Imports System.Net.Mail

Imports System.Runtime.InteropServices
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Core
Imports System.Reflection
Imports Microsoft.Win32


Module conexion
    Public cnn As SqlConnection
    Public sql As String
    Public nombre As String
    Public cmd As SqlCommand
    Public reader As SqlDataReader
    Public da As SqlClient.SqlDataAdapter
    Public cb As SqlClient.SqlCommandBuilder
    Public ds As DataSet
    Public ret As Integer
    Public servidor As Integer
    Public strLine As String
    Public message As New MailMessage
    Public maximo As Double
    Public para, desde, mensaje, usuario, contasena, cuerpo, asunto, servidor_email, puerto, adjunto, ficha_aprendiz As String





    Public Sub conectado()
        Form1.Timer1.Enabled = False
        Try


            cnn = New SqlConnection
            'cnn.ConnectionString = ("data source= 127.0.0.1; initial catalog=Academic2015; integrated security=true")
            'cnn.ConnectionString = ("data source= LAPTOP-AN4HNJ4L; initial catalog=Academicsoft2019; user id = katherine; password = 123456789")
            'cnn.ConnectionString = ("data source= FONDFPCAA7F007; initial catalog=Academicsoft2019; user id = Administrador; password = 123456789")
            'cnn.ConnectionString = ("data source= 10.100.179.142; initial catalog=Academicsoft2015; user id = javier_carrillo; password = 17959076")
            cnn.ConnectionString = ("data source= 10.194.128.254; initial catalog=Academicsoft; user id =jaider; password = 17959076")

            cnn.Open()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try




    End Sub
    Public Sub conectado1()
        Form1.Timer1.Enabled = False
        Try


            cnn = New SqlConnection
            'cnn.ConnectionString = ("data source= 127.0.0.1; initial catalog=Academic2015; integrated security=true")
            'cnn.ConnectionString = ("data source= LAPTOP-AN4HNJ4L; initial catalog=Academicsoft2019; user id = katherine; password = 123456789")
            'cnn.ConnectionString = ("data source= FONDFPCAA7F007; initial catalog=Academicsoft2019; user id = Administrador; password = 123456789")
            'cnn.ConnectionString = ("data source= 10.100.179.142; initial catalog=Academicsoft2015; user id = javier_carrillo; password = 17959076")
            cnn.ConnectionString = ("data source= 10.194.128.254; initial catalog=acceso; user id =jaider; password = 17959076")

            cnn.Open()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try




    End Sub
    Public Sub maximo_id_competencia()

        sql = "Select id from dbo.competencia order by id desc"
        conectado()
        cmd = New SqlCommand(sql, cnn)
        reader = cmd.ExecuteReader

        If reader.Read Then
            maximo = reader("Id")
        End If
        cerrar_conexion()
    End Sub

    Public Sub maximo_id_programacion()

        sql = "Select id from dbo.programacion order by id desc"
        conectado()
        cmd = New SqlCommand(sql, cnn)
        reader = cmd.ExecuteReader

        If reader.Read Then
            maximo = reader("Id")
        End If
        cerrar_conexion()
    End Sub


    Public Sub cerrar_conexion()
        cnn.Close()
        ' Form1.Timer1.Enabled = True
    End Sub


    Public Sub enviar()
        Dim smtp As New System.Net.Mail.SmtpClient(servidor_email, puerto)



        message.From = New MailAddress(desde)
        message.To.Add(para)
        message.Body = cuerpo
        message.Subject = asunto
        message.Priority = MailPriority.Normal
        smtp.EnableSsl = True

        smtp.UseDefaultCredentials = True
        smtp.Credentials = New Net.NetworkCredential(desde, contasena)
        Try
            'smtp.Timeout = 100
            smtp.Send(message)
            MsgBox("Mensaje enviado con Exito!")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub




    Dim oApp As Outlook._Application

    Dim oMsg As Outlook._MailItem




    Public Sub enviar_correo()

        oApp = New Outlook.Application()
        oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)


        oApp = New Outlook.Application()


        oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)

        oMsg.Subject = asunto
        oMsg.Body = cuerpo

        oMsg.To = para

        Dim sSource As String = adjunto
        Dim sDisplayName As String = "adjunto"
        Dim sBodyLen As String = oMsg.Body.Length
        Dim oAttachs As Outlook.Attachments = oMsg.Attachments
        Dim oAttach As Outlook.Attachment

        oAttach = oAttachs.Add(sSource, , sBodyLen + 1, sDisplayName)


        oMsg.Send()
        MsgBox("Mensaje enviado con exito")
        'oApp = Nothing
        'oMsg = Nothing
        'oAttach = Nothing
        'oAttachs = Nothing


    End Sub
    Public Sub enviar_correosegn()
        'Creamos un Objeto que hará referencia a nuestra aplicación Outlook 
        Dim m_OutLook As Outlook.Application
        Try
            'Creamos un Objeto tipo Mail 
            Dim objMail As Outlook.MailItem
            'Inicializamos nuestra apliación OutLook 
            m_OutLook = New Outlook.Application
            'Creamos una instancia de un objeto tipo MailItem 
            objMail = m_OutLook.CreateItem(Outlook.OlItemType.olMailItem)
            'Asignamos las propiedades a nuestra Instancial del objeto 
            'MailItem 
            objMail.To = "khattherine@gmail.com" 'para
            objMail.Subject = asunto
            objMail.Body = cuerpo

            'Si queremos enviar un archivo adjunto usamos este codigo… 
            ' Dim sSource As String = ""
            ' Dim sDisplayName As String = "adjunto"
            '  Dim sBodyLen As String = objMail.Body.Length
            '  Dim oAttachs As Outlook.Attachments = objMail.Attachments
            '  Dim oAttach As Outlook.Attachment
            ' oAttach = oAttachs.Add(sSource, , sBodyLen + 1, sDisplayName)

            'Enviamos nuestro Mail y listo! 
            objMail.Send()
            'Desplegamos un mensaje indicando que todo fue exitoso 
            MessageBox.Show("Envío exitoso.", "Enviar Mail", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Catch ex As Exception
            'Si se produce algun Error 
            MessageBox.Show("Error al enviar mail " & ex.ToString)
        Finally
            m_OutLook = Nothing ' Destruimos el objeto (recoger la basura…) 
        End Try

    End Sub
    Public Sub enviar_correoseg()
        'Creamos un Objeto que hará referencia a nuestra aplicación Outlook 

        Try
            oApp = New Outlook.Application()
            oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)


            oApp = New Outlook.Application()


            oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)
            oMsg.Subject = asunto
            oMsg.HTMLBody = cuerpo
            oMsg.To = para '"khattherine@gmail.com" 


            oMsg.Send()

            'Desplegamos un mensaje indicando que todo fue exitoso 
            MessageBox.Show("Envío exitoso de Seguimiento.", "Enviar Mail", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Catch ex As Exception
            'Si se produce algun Error 
            MessageBox.Show("Error al enviar mail " & ex.ToString)
        Finally
            oApp = Nothing ' Destruimos el objeto (recoger la basura…) 
        End Try

    End Sub
    Public Sub enviar_correosegAdjunto()
        'Creamos un Objeto que hará referencia a nuestra aplicación Outlook 

        Try
            oApp = New Outlook.Application()
            oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)


            oApp = New Outlook.Application()


            oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)
            oMsg.Subject = asunto
            oMsg.HTMLBody = cuerpo
            oMsg.To = para '"khattherine@gmail.com" 

            'Si queremos enviar un archivo adjunto usamos este codigo… 
            Dim sSource As String = adjunto
            Dim sDisplayName As String = "adjunto"
            Dim sBodyLen As String = oMsg.Body.Length
            Dim oAttachs As Outlook.Attachments = oMsg.Attachments
            Dim oAttach As Outlook.Attachment
            oAttach = oAttachs.Add(sSource, , sBodyLen + 1, sDisplayName)

            'Enviamos nuestro Mail y listo! 

            oMsg.Send()

            'Desplegamos un mensaje indicando que todo fue exitoso 
            MessageBox.Show("Envío exitoso de Seguimiento.", "Enviar Mail", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Catch ex As Exception
            'Si se produce algun Error 
            MessageBox.Show("Error al enviar mail " & ex.ToString)
        Finally
            oApp = Nothing ' Destruimos el objeto (recoger la basura…) 
        End Try

    End Sub

    Public Sub enviar_correo1()
        oApp = New Outlook.Application()
        oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)


        oApp = New Outlook.Application()


        oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)
        oMsg.Subject = asunto

        oMsg.Body = cuerpo

        oMsg.To = para

        ' Dim sSource As String = adjunto
        ' Dim sDisplayName As String = "adjunto"
        ' Dim sBodyLen As String = oMsg.Body.Length
        ' Dim oAttachs As Outlook.Attachments = oMsg.Attachments
        ' Dim oAttach As Outlook.Attachment

        ' oAttach = oAttachs.Add(sSource, , sBodyLen + 1, sDisplayName)


        oMsg.Send()
        MsgBox("Mensaje enviado con exito")
        'oApp = Nothing
        'oMsg = Nothing
        'oAttach = Nothing
        'oAttachs = Nothing


    End Sub
    Public Sub LlenarDataGrids(DataGridView As DataGridView)

        conectado()
        Try
            da = New SqlClient.SqlDataAdapter(sql, cnn)
            cb = New SqlClient.SqlCommandBuilder(da)
            ds = New DataSet
            da.Fill(ds, "tabla")
            DataGridView.DataSource = ds
            DataGridView.DataMember = "tabla"
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        cerrar_conexion()
    End Sub
    Public Sub enviar_correosinconfirmar()
        oApp = New Outlook.Application()
        oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)


        oApp = New Outlook.Application()


        oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)
        oMsg.Subject = asunto
        oMsg.Body = cuerpo

        oMsg.To = para
        If adjunto <> "" Then
            Dim sSource As String = adjunto
            Dim sDisplayName As String = "adjunto"
            Dim sBodyLen As String = oMsg.Body.Length
            Dim oAttachs As Outlook.Attachments = oMsg.Attachments
            Dim oAttach As Outlook.Attachment

            oAttach = oAttachs.Add(sSource, , sBodyLen + 1, sDisplayName)
        End If

        oMsg.Send()

        'oApp = Nothing
        'oMsg = Nothing
        'oAttach = Nothing
        'oAttachs = Nothing

        ' MsgBox("Mensaje enviado con Exito!")
    End Sub

    '/* Llena los combos de departamentos*/'
    Public Sub llenarcombos(combo As ComboBox, display As String, value As String)


        conectado()
        Try

            da = New SqlDataAdapter(sql, cnn)
            cb = New SqlClient.SqlCommandBuilder(da)
            ds = New DataSet
            da.Fill(ds, "tabla")
            combo.DataSource = ds.Tables("tabla")
            combo.DisplayMember = display
            combo.ValueMember = value
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        cerrar_conexion()
    End Sub

    Public Sub Agregar()
        conectado()
        cmd = New SqlClient.SqlCommand(sql, cnn)
        ' MsgBox(sql)
        cmd.ExecuteNonQuery()

        MessageBox.Show("Transaccion Exitosa.", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information)

        cerrar_conexion()
    End Sub

    Public Function ListarDatos() As DataTable
        Dim MiAdapter As SqlDataAdapter = New SqlDataAdapter(sql, cnn)

        Dim MiDataSet As New DataSet
        MiAdapter.Fill(MiDataSet)
        Return MiDataSet.Tables(0)

    End Function
    Public Function ListarDatos1() As DataTable
        Dim MiAdapter As SqlDataAdapter = New SqlDataAdapter(sql, cnn)

        Dim MiDataSet As New DataSet
        MiAdapter.Fill(MiDataSet)
        Return MiDataSet.Tables(0)

    End Function




    Sub ImportarWord()
        'Dim oWord As Word.Application
        ' Dim oDoc As Word.Document
        ' Dim oTable As Word.Table
        'Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph
        'Dim oPara3 As Word.Paragraph, oPara4 As Word.Paragraph
        'Dim oRng As Word.Range
        ' Dim oShape As Word.InlineShape
        ' Dim oChart As Object
        ' Dim Pos As Double

        'Start Word and open the document template.
        'oWord = CreateObject("Word.Application")
        ' oWord.Visible = True
        ' oDoc = oWord.Documents.Add



        'FileCopy(My.Computer.FileSystem.CurrentDirectory & "\Reporte.doc", "c:\ReporteSeguimientos\" & datagridEvidencias.Rows(1).Cells(1).Value & ".doc")
        'oDoc = oWord.Documents.Open("c:\ReporteSeguimientos\" & datagridEvidencias.Rows(1).Cells(1).Value & ".doc")
        'oPara1 = oDoc.Content.Paragraphs.Add
        'oPara1.Range.Text = "Segimiento del  Aprendiz: " & datagridEvidencias.Rows(0).Cells("aprendiz").Value
        'oPara1.Range.Font.Bold = True
        'oPara1.Format.SpaceAfter = 24    '24 pt spacing after paragraph.
        'oPara1.Range.InsertParagraphAfter()

        'oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        ' oPara2.Range.Text = "Ficha:  " & datagridEvidencias.Rows(1).Cells(11).Value
        ' oPara2.Format.SpaceAfter = 24
        ' oPara2.Range.InsertParagraphAfter()

        ' oPara3 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        ' oPara3.Range.Text = "Programa de Formación: " & datagridEvidencias.Rows(1).Cells(10).Value
        'oPara3.Format.SpaceAfter = 6
        'oPara3.Range.InsertParagraphAfter()

        ' oPara3 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        ' oPara3.Range.Text = "Empresa: " & datagridEvidencias.Rows(1).Cells(2).Value
        ' oPara3.Format.SpaceAfter = 6
        ' oPara3.Range.InsertParagraphAfter()


        'For i = 0 To datagridEvidencias.RowCount - 1
        'If datagridEvidencias.Rows(i).Cells(0).Value <> "" Then
        'Insert a paragraph at the beginning of the document.


        ' End If

        ' Next
        'Insert a paragraph at the beginning of the document.
        'oPara1 = oDoc.Content.Paragraphs.Add
        ' oPara1.Range.Text = "Heading 1"
        'oPara1.Range.Font.Bold = True
        'oPara1.Format.SpaceAfter = 24    '24 pt spacing after paragraph.
        'oPara1.Range.InsertParagraphAfter()

        'Insert a paragraph at the end of the document.
        '** \endofdoc is a predefined bookmark.
        'oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        'oPara2.Range.Text = "Heading 2"
        'oPara2.Format.SpaceAfter = 6
        'oPara2.Range.InsertParagraphAfter()

        'Insert another paragraph.
        'oPara3 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        'oPara3.Range.Text = "This is a sentence of normal text. Now here is a table:"
        'oPara3.Range.Font.Bold = False
        'oPara3.Format.SpaceAfter = 24
        'oPara3.Range.InsertParagraphAfter()

        'Insert a 3 x 5 table, fill it with data, and make the first row
        'bold and italic.
        ' Dim r As Integer, c As Integer
        'oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 3, 5)
        'oTable.Range.ParagraphFormat.SpaceAfter = 6
        'For r = 1 To 3
        'For c = 1 To 5
        'oTable.Cell(r, c).Range.Text = "r" & r & "c" & c
        'Next
        ' Next
        ' oTable.Rows.Item(1).Range.Font.Bold = True
        ' oTable.Rows.Item(1).Range.Font.Italic = True

        'Add some text after the table.
        'oTable.Range.InsertParagraphAfter()
        'oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        'oPara4.Range.InsertParagraphBefore()
        'oPara4.Range.Text = "And here's another table:"
        ' oPara4.Format.SpaceAfter = 24
        'oPara4.Range.InsertParagraphAfter()

        ''Insert a 5 x 2 table, fill it with data, and change the column widths.
        'oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 2)
        'oTable.Range.ParagraphFormat.SpaceAfter = 6
        ' For r = 1 To 5
        'For c = 1 To 2
        'oTable.Cell(r, c).Range.Text = "r" & r & "c" & c
        'Next
        ' Next
        ' oTable.Columns.Item(1).Width = oWord.InchesToPoints(2)   'Change width of columns 1 & 2
        'oTable.Columns.Item(2).Width = oWord.InchesToPoints(3)

        'Keep inserting text. When you get to 7 inches from top of the
        'document, insert a hard page break.
        ' Pos = oWord.InchesToPoints(7)
        ' oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()
        'Do
        'oRng = oDoc.Bookmarks.Item("\endofdoc").Range
        ' oRng.ParagraphFormat.SpaceAfter = 6
        ' oRng.InsertAfter("A line of text")
        ' oRng.InsertParagraphAfter()
        ' Loop While Pos >= oRng.Information(Word.WdInformation.wdVerticalPositionRelativeToPage)
        ' oRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        ' oRng.InsertBreak(Word.WdBreakType.wdPageBreak)
        ' oRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        ' oRng.InsertAfter("We're now on page 2. Here's my chart:")
        'oRng.InsertParagraphAfter()

        'Insert a chart and change the chart.
        'oShape = oDoc.Bookmarks.Item("\endofdoc").Range.InlineShapes.AddOLEObject( _
        '    ClassType:="MSGraph.Chart.8", FileName _
        '    :="", LinkToFile:=False, DisplayAsIcon:=False)
        'oChart = oShape.OLEFormat.Object
        'oChart.charttype = 4 'xlLine = 4
        'oChart.Application.Update()
        'oChart.Application.Quit()
        'If desired, you can proceed from here using the Microsoft Graph 
        'Object model on the oChart object to make additional changes to the
        'chart.
        ' oShape.Width = oWord.InchesToPoints(6.25)
        ' oShape.Height = oWord.InchesToPoints(3.57)

        'Add text after the chart.
        'oRng = oDoc.Bookmarks.Item("\endofdoc").Range
        ' oRng.InsertParagraphAfter()
        ' oRng.InsertAfter("THE END.")

        'All done. Close this form.
        ' Me.Close()




    End Sub
End Module
