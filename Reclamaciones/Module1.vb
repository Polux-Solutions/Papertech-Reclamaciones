Option Explicit On
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop.Outlook
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Diagnostics
Imports System
Imports System.IO
Imports Microsoft.VisualBasic
Imports System.Text

Module Module1
    Public sqlConn As System.Data.SqlClient.SqlConnection
    Public Servidor As String
    Public BD As String
    Public Empresa As String
    Public Usuario As String
    Public Password As String
    Public FicheroLog As String
    Public ds As DataSet
    Public dsCalidad As DataSet

    Sub Main()
        Dim xId As String
        Dim xTipo As String

        If Environment.GetCommandLineArgs.Length < 3 Then
            MsgBox("No se ha pasado Tipo ni identificador")
            Exit Sub
        End If


        xId = Environment.GetCommandLineArgs(1)
        xTipo = Environment.GetCommandLineArgs(2)

        Dim MiProceso As String = System.Diagnostics.Process.GetCurrentProcess.ProcessName.ToString
        Dim i As Integer

        i = InStr(MiProceso, ".", CompareMethod.Text)
        If i > 0 Then MiProceso = MiProceso.Substring(0, i - 1)
        If UBound(Diagnostics.Process.GetProcessesByName(MiProceso)) > 0 Then
            Exit Sub
        End If

        System.Threading.Thread.Sleep(2000)

        If Leer_Parametros() Then
            Informes(xTipo, xId)
        End If
    End Sub

    Private Sub Informes(xTipo As String, xId As String)
        Dim wdapp As Microsoft.Office.Interop.Word.Application
        Dim wdDoc As Microsoft.Office.Interop.Word.Document
        Dim wdRange As Microsoft.Office.Interop.Word.Range
        Dim WdFields As Microsoft.Office.Interop.Word.Fields
        Dim oApp As Microsoft.Office.Interop.Outlook.Application
        Dim oMsg As Microsoft.Office.Interop.Outlook.MailItem
        Dim oAttach As Microsoft.Office.Interop.Outlook.Attachments
        Dim FileIn As String
        Dim FileOut As String
        Dim FileTMP As String
        Dim n As Integer
        Dim Firma As String
        Dim str As New StringBuilder

        If Not Cargar_Dataset(xId) Then Exit Sub

        n = 0

        For Each dt As DataRow In ds.Tables(0).Rows
            n += 1
            If n = 1 Then
                oApp = New Microsoft.Office.Interop.Outlook.Application
                oMsg = oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
                oAttach = oMsg.Attachments
                oMsg.To = dt.Item("Correo Incidencias").ToString
                'oMsg.To = "jcarlos.ramos@polux-solutions.es"
                oMsg.CC = "mariola.llera@sonoco.com"
                If xTipo = "H" Then oMsg.Subject = "INCIDENCIA HUMEDAD " + dt.Item("Buy-from Vendor Name").ToString
                If xTipo = "C" Then oMsg.Subject = "INCIDENCIA CALIDAD " + dt.Item("Buy-from Vendor Name").ToString
                If xTipo = "HC" Then oMsg.Subject = "INCIDENCIA HUMEDAD - CALIDAD " + dt.Item("Buy-from Vendor Name").ToString

                oMsg.BodyFormat = OlBodyFormat.olFormatHTML
                oMsg.Display()
                Firma = oMsg.HTMLBody
                str.Append("<html> <head>  </head>" +
                            " <body>" + _
                            " <br/><br/>" + _
                            " Estimado proveedor: <br/><br/>" + _
                            " Adjunto envío documento sobre la incidencia detectada en nuestras instalaciones <br/>" + _
                            " Para cualquier consulta no dude en contactar conmigo. <br/><br/>" + _
                            dt.Item("Informe No_").ToString + _
                            "</body></html>")

                str.Append(Firma)
                oMsg.HTMLBody = str.ToString
            End If

            oMsg.Subject += " " + dt.Item("Albaran Pesaje").ToString

            FileIn = ""
            Select Case xTipo
                Case "H" : FileIn = dt.Item("Documento Reclamación Humedad").ToString
                Case "C" : FileIn = dt.Item("Documento Reclamación Calidad").ToString
                Case "HC" : FileIn = dt.Item("Documento Reclamación HU-CA").ToString
            End Select

            FileOut = dt.Item("Informe No_")
            If FileOut = "" Then FileOut = "007"

            FileTMP = My.Computer.FileSystem.GetTempFileName
            System.IO.File.Copy(FileIn, FileTMP, True)
            FileOut = Path.GetTempPath + FileOut + ".PDF"

            wdapp = New Microsoft.Office.Interop.Word.Application
            wdDoc = wdapp.Documents.Open(FileTMP, True, False)
            wdapp.Visible = False
            WdFields = wdDoc.Fields

            wdRange = WdFields.Item(1).Result
            wdRange.Text = dt.Item("Buy-from Vendor Name")
            wdRange = WdFields.Item(2).Result
            wdRange.Text = Format(Now, "dd-MM-yyyy")
            wdRange = WdFields.Item(3).Result
            wdRange.Text = Format(dt.Item("FechaPesada"), "dd-MM-yyyy")
            wdRange = WdFields.Item(4).Result
            wdRange.Text = dt.Item("Matricula")
            wdRange = WdFields.Item(5).Result
            wdRange.Text = dt.Item("Albaran Pesaje")
            wdRange = WdFields.Item(6).Result
            wdRange.Text = dt.Item("No_")
            If dt.Item("No_").ToString.Length > 2 Then
                If dt.Item("No_").ToString.Substring(0, 2) = "WP" Then wdRange.Text = dt.Item("No_").ToString.Substring(2, dt.Item("No_").ToString.Length - 2)
            End If

            If xTipo = "H" Then
                wdRange = WdFields.Item(7).Result
                wdRange.Text = Format(dt.Item("Cantidad Báscula") * 1000, "#,##0")
                wdRange = WdFields.Item(8).Result
                wdRange.Text = Format(dt.Item("Porcentaje Dto_ Báscula") / 100, "#,##0.000")
                wdRange = WdFields.Item(9).Result
                wdRange.Text = Format(dt.Item("Quantity") * 1000, "#,##0")
            End If


            If xTipo = "C" Then
                Cargar_Dataset_Calidad(dt.Item("Document No_"), dt.Item("Pesaje Id"))

                If dsCalidad.Tables(0).Rows.Count = 2 Then
                    Dim dt1 As DataRow = dsCalidad.Tables(0).Rows(0)
                    Dim dt2 As DataRow = dsCalidad.Tables(0).Rows(1)

                    wdRange = WdFields.Item(7).Result
                    wdRange.Text = Format((dt1.Item("Cantidad Báscula") + dt2.Item("Cantidad Báscula")) * 1000, "#,##0")

                    Dim porcentaje As Single

                    porcentaje = Math.Round(dt1.Item("Cantidad Báscula") / (dt1.Item("Cantidad Báscula") + dt2.Item("Cantidad Báscula")) * 100, 2)
                    wdRange = WdFields.Item(8).Result
                    wdRange.Text = Format(porcentaje, "#,##0")
                    wdRange = WdFields.Item(9).Result
                    If dt1.Item("No_").ToString.Length > 2 Then
                        If dt1.Item("No_").ToString.Substring(0, 2) = "WP" Then wdRange.Text = dt1.Item("No_").ToString.Substring(2, dt1.Item("No_").ToString.Length - 2)
                    End If
                    wdRange = WdFields.Item(10).Result
                    wdRange.Text = Format(dt1.Item("Cantidad Báscula") * 1000, "#,##0")

                    porcentaje = Math.Round(dt2.Item("Cantidad Báscula") / (dt1.Item("Cantidad Báscula") + dt2.Item("Cantidad Báscula")) * 100, 2)
                    wdRange = WdFields.Item(11).Result
                    wdRange.Text = Format(porcentaje, "#,##0")
                    wdRange = WdFields.Item(12).Result
                    If dt2.Item("No_").ToString.Length > 2 Then
                        If dt2.Item("No_").ToString.Substring(0, 2) = "WP" Then wdRange.Text = dt2.Item("No_").ToString.Substring(2, dt2.Item("No_").ToString.Length - 2)
                    End If
                    wdRange = WdFields.Item(13).Result
                    wdRange.Text = Format(dt2.Item("Cantidad Báscula") * 1000, "#,##0")
                End If

            End If


            If xTipo = "HC" Then
                Cargar_Dataset_Calidad(dt.Item("Document No_"), dt.Item("Pesaje Id"))

                If dsCalidad.Tables(0).Rows.Count = 2 Then
                    Dim dt1 As DataRow = dsCalidad.Tables(0).Rows(0)
                    Dim dt2 As DataRow = dsCalidad.Tables(0).Rows(1)

                    wdRange = WdFields.Item(7).Result
                    wdRange.Text = Format((dt1.Item("Cantidad Báscula") + dt2.Item("Cantidad Báscula")) * 1000, "#,##0")
                    wdRange = WdFields.Item(8).Result
                    wdRange.Text = Format(dt.Item("Porcentaje Dto_ Báscula") / 100, "#,##0.000")
                    wdRange = WdFields.Item(9).Result
                    wdRange.Text = Format((dt.Item("Quantity") + dt2.Item("Quantity")) * 1000, "#,##0")
                    Dim porcentaje As Single

                    porcentaje = Math.Round(dt1.Item("Cantidad Báscula") / (dt1.Item("Cantidad Báscula") + dt2.Item("Cantidad Báscula")) * 100, 2)
                    wdRange = WdFields.Item(10).Result
                    wdRange.Text = Format(porcentaje, "#,##0")
                    wdRange = WdFields.Item(11).Result
                    If dt1.Item("No_").ToString.Length > 2 Then
                        If dt1.Item("No_").ToString.Substring(0, 2) = "WP" Then wdRange.Text = dt1.Item("No_").ToString.Substring(2, dt1.Item("No_").ToString.Length - 2)
                    End If
                    wdRange = WdFields.Item(12).Result
                    wdRange.Text = Format(dt1.Item("Quantity") * 1000, "#,##0")

                    porcentaje = Math.Round(dt2.Item("Cantidad Báscula") / (dt1.Item("Cantidad Báscula") + dt2.Item("Cantidad Báscula")) * 100, 2)
                    wdRange = WdFields.Item(13).Result
                    wdRange.Text = Format(porcentaje, "#,##0")
                    wdRange = WdFields.Item(14).Result
                    If dt2.Item("No_").ToString.Length > 2 Then
                        If dt2.Item("No_").ToString.Substring(0, 2) = "WP" Then wdRange.Text = dt2.Item("No_").ToString.Substring(2, dt2.Item("No_").ToString.Length - 2)
                    End If
                    wdRange = WdFields.Item(15).Result
                    wdRange.Text = Format(dt2.Item("Quantity") * 1000, "#,##0")
                End If

            End If

            wdapp.ActiveDocument.ExportAsFixedFormat(FileOut, _
                        Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, False, _
                        Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForPrint, _
                        Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument)
            wdDoc.Close()
            wdapp.Quit()

            oAttach.Add(FileOut)
        Next

        If n = 0 Then oMsg.Delete()
    End Sub


    Public Function Leer_Parametros() As Boolean
        Dim Config As New Configuration.AppSettingsReader

        Leer_Parametros = True
        Try
            Servidor = Config.GetValue("SERVER", GetType(System.String)).ToString
            BD = Config.GetValue("BD", GetType(System.String)).ToString
            Empresa = Config.GetValue("EMPRESA", GetType(System.String)).ToString
            Usuario = Config.GetValue("USUARIO", GetType(System.String)).ToString
            Password = Config.GetValue("CONTRASEÑA", GetType(System.String)).ToString
            FicheroLog = Config.GetValue("LOG", GetType(System.String)).ToString
        Catch ex As System.Exception
            MsgBox("Error al leer parámetros " & ex.Message, MsgBoxStyle.Critical)
            Leer_Parametros = False
        End Try
    End Function


    Public Sub Log(ByVal texto As String)
        Dim sr As System.IO.StreamWriter

        Try
            sr = New System.IO.StreamWriter(FicheroLog, True)
            sr.WriteLine("V7.0 " & Format(Now, "dd.MM.yy hh:mm:ss") & "   " & texto)
            sr.Close()
        Catch ex As System.Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try

    End Sub

End Module
