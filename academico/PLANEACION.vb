Imports Microsoft.Office.Interop
Imports System.Data.SqlClient


Module PLANEACION
    Public XLApp As Excel.Application  'Aplicación Excel en varaible XLApp
    Public XLBook As Excel.Workbook    'Libro de Excel en variable XLBook
    Public XLSheet As Excel.Worksheet  'Hoja de cálculo en variable XLSheet
    Public XLSheet2 As Excel.Worksheet  'Hoja de cálculo en variable XLSheet
    Dim fecha_calculo As Date


    Sub crealibro_planeacion()
        Dim sin_proyecto As Integer = 0
        Dim cant_cursos As Integer = 0
        Dim competecnias As Integer = 0
        Dim cant_horas As Double = 0

        Dim i As Integer
        Dim fecha_inicio As Date
        Dim fecha_final As Date
        Dim color As Boolean

        For i = 0 To Form1.dtplaneacion.RowCount - 1

            Dim contador As Integer = 7
            XLSheet.Range("B7:B22").EntireRow.Copy()
            contador += 16
            XLSheet.Range("B23").EntireRow.Insert()

            XLSheet.Range("B24").Value = Form1.dtplaneacion.Rows(i).Cells("ficha").Value
            cant_cursos += 1
            sql = "Select * from dbo.grupos where ficha=" & XLSheet.Range("B24").Value & ""
            Clipboard.SetText(sql)
            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader

            If reader.Read Then

                fecha_inicio = reader("Fecha_inicio")
                XLSheet.Range("K24").Value = fecha_inicio.Day
                XLSheet.Range("M24").Value = fecha_inicio.Month
                XLSheet.Range("O24").Value = fecha_inicio.Year

                fecha_final = reader("Fecha_terminacion")
                XLSheet.Range("V24").Value = fecha_final.Day
                XLSheet.Range("X24").Value = fecha_final.Month
                XLSheet.Range("Z24").Value = fecha_final.Year

                XLSheet.Range("AG24").Value = reader("Aprendices_matriculados")
                XLSheet.Range("L26").Value = reader("Nombre_curso")
                XLSheet.Range("H28").Value = reader("Municipio")
                XLSheet.Range("Z28").Value = reader("Lugar")
                XLSheet.Range("N30").Value = reader("Instructor_responsable")

            Else
                GoTo siguiente
            End If

            cerrar_conexion()
            Dim texto_instructores_fin, texto_instructores_secionada As String
            texto_instructores_fin = ""
            texto_instructores_secionada = ""
            sql = "Select * from programacion where ficha ='" & XLSheet.Range("B24").Value & "' and  instructor is not null"
            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader
            While reader.Read
                
                    texto_instructores_fin += reader("instructor") & " - "
             
            End While
            XLSheet.Range("AY26").Value = texto_instructores_fin
            XLSheet.Range("AZ26").Value = texto_instructores_secionada
            Dim fill_compe As Integer = 26

            cerrar_conexion()
            sql = "Select instructor, count(instructor) as numero from programacion where ficha= '" & XLSheet.Range("B24").Value & "' and instructor is not null  group by instructor"
            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader
            While reader.Read
                XLSheet.Range("BA" & fill_compe).Value = reader("instructor")
                XLSheet.Range("BB" & fill_compe).Value = reader("numero")
                fill_compe += 1
            End While


            sql = "Select * from dbo.proyecto where ficha=" & XLSheet.Range("B24").Value & ""
            conectado()
            cmd = New SqlCommand(sql, cnn)
            reader = cmd.ExecuteReader

            If reader.Read Then

                XLSheet.Range("G32").Value = reader("codigo") & " - " & reader("nombre")
            Else
                XLSheet.Range("G32").Value = "NO SE HA REGISTRADO EL PROYECTO EN EL PORTAFOLIO DEL COORDINADOR ACADEMICO"
                sin_proyecto += 1

            End If
            cerrar_conexion()


            Dim fila_competencia As Integer = 36
            If color Then
                color = False
            Else
                color = True
            End If

            While Form1.dtplaneacion.Rows(i).Cells("ficha").Value = XLSheet.Range("B24").Value

                XLSheet.Range("B" & fila_competencia).EntireRow.Copy()
                fila_competencia += 1
                XLSheet.Range("B" & fila_competencia).EntireRow.Insert()
                XLSheet.Range("B" & fila_competencia - 1).Value = Form1.dtplaneacion.Rows(i).Cells("id").Value
                XLSheet.Range("C" & fila_competencia - 1).Value = Form1.dtplaneacion.Rows(i).Cells("competencia").Value
                competecnias += 1
                XLSheet.Range("W" & fila_competencia - 1).Value = Form1.dtplaneacion.Rows(i).Cells("Duracion").Value


                XLSheet2.Range("A" & fila_competencia - 27).EntireRow.Copy()
                XLSheet2.Range("A" & fila_competencia - 26).EntireRow.Insert()
                XLSheet2.Range("B" & fila_competencia - 26).Value = XLSheet.Range("B24").Value
                XLSheet2.Range("C" & fila_competencia - 26).Value = XLSheet.Range("L26").Value
                XLSheet2.Range("D" & fila_competencia - 26).Value = XLSheet.Range("H28").Value
                XLSheet2.Range("E" & fila_competencia - 26).Value = Form1.dtplaneacion.Rows(i).Cells("competencia").Value
                XLSheet2.Range("X" & fila_competencia - 26).Value = Form1.dtplaneacion.Rows(i).Cells("Duracion").Value
                XLSheet2.Range("A" & fila_competencia - 26).Value = Form1.dtplaneacion.Rows(i).Cells("id").Value
                If color Then
                    Dim IND As Integer = 33
                    XLSheet2.Range("B" & fila_competencia - 26 & ":AY" & fila_competencia - 26).Interior.Color = RGB(229, 255, 255)
                   
                End If


                i += 1

            End While
            XLSheet.Range("B" & fila_competencia).EntireRow.Delete()
            Dim k As Integer
            Dim suma As Integer = 0
            If Form1.cheksimulacion.Checked Then
                For k = 36 To fila_competencia
                    suma += XLSheet.Range("W" & k).Value
                Next

                XLSheet.Range("W" & fila_competencia).Value = suma
                If fecha_final.Year > Form1.cbano.Text Then
                    fecha_calculo = "31/12/" & Form1.cbano.Text
                    calcula_entre_fechas(fecha_calculo)
                Else
                    fecha_calculo = fecha_final.Date
                    calcula_entre_fechas(fecha_calculo)
                End If

                cant_horas += suma

            End If
           



siguiente:

        Next
        XLSheet.Range("G6").Value = cant_cursos
        XLSheet.Range("U6").Value = competecnias
        XLSheet.Range("AA6").Value = cant_horas
        XLSheet.Range("AT6").Value = sin_proyecto
        XLSheet.Range("B7:B22").EntireRow.Delete()
        XLApp.Application.DisplayAlerts = False
        XLBook.SaveAs("C:\academico\" & "planeacion" & Now.Date.Day & Now.Date.Month & Now.Date.Year & Now.Date.Hour & Now.Date.Minute & Now.Date.Second & ".xls")


    End Sub

    Dim fecha_imprimir_LV As Date
    Dim fecha_imprimir_LS As Date


    Sub calcula_entre_fechas(fecha As Date)

        Dim horas_LV As Integer = 0
        Dim horas_LS As Integer = 0

        Dim sabado As Integer
        sql = "Select * from dbo.calendario"
        conectado()
        Dim MiTabla As New DataTable()
        Dim Comando As New SqlDataAdapter(sql, cnn)
        Comando.Fill(MiTabla)
        Dim m As Integer
        Dim rta As Boolean
        For sabado = 0 To 1
            Dim dias As Integer = 0
            Dim horas_cuenta As Integer = XLSheet.Range("AO24").Value
            Dim fecha_inicio_calculo As Date = "01/01/" & Form1.cbano.Text
            While fecha_inicio_calculo <= fecha.AddDays(-1)
                rta = 0
                For m = 0 To MiTabla.Rows.Count - 1
                    If fecha_inicio_calculo = MiTabla.Rows(m).Item("fecha_completa") Then
                        rta = 1
                    End If
                Next



                If rta Then
                    fecha_inicio_calculo = fecha_inicio_calculo.AddDays(1)
                ElseIf sabado = 1 Then

                    If fecha_inicio_calculo.DayOfWeek <> DayOfWeek.Sunday Then
                        horas_LS += 8
                        dias += 1
                        XLSheet.Range("AY25").Value = dias
                        XLSheet.Range("AV24").Value = horas_LS
                        XLSheet.Range("AY25").Value = dias

                        If horas_cuenta > 0 Then
                            fecha_imprimir_LS = fecha_inicio_calculo
                            XLSheet.Range("AW25").Value = fecha_imprimir_LS
                            horas_cuenta -= 8
                        End If


                    End If
                    fecha_inicio_calculo = fecha_inicio_calculo.AddDays(1)
                ElseIf sabado = 0 Then
                    If fecha_inicio_calculo.DayOfWeek <> DayOfWeek.Sunday And fecha_inicio_calculo.DayOfWeek <> DayOfWeek.Saturday Then
                        horas_LV += 8
                        dias += 1
                        XLSheet.Range("AY24").Value = dias
                        XLSheet.Range("AT24").Value = horas_LV
                        If horas_cuenta > 0 Then
                            fecha_imprimir_LV = fecha_inicio_calculo
                            XLSheet.Range("AW24").Value = fecha_imprimir_LV
                            horas_cuenta -= 8
                        End If

                    End If
                    fecha_inicio_calculo = fecha_inicio_calculo.AddDays(1)

                End If



            End While

        Next

      








    End Sub




End Module
