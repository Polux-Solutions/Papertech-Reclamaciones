Option Explicit On
Imports System.Data.SqlClient
Module Datos

    Public Function Abrir_BD() As Boolean
        Abrir_BD = True
        Try
            sqlConn = New SqlConnection("server=" & Servidor & ";uid=" & Usuario & ";pwd=" & Password & ";database=" & BD)
            sqlConn.Open()
        Catch ex As SqlClient.SqlException
            Log("ERROR APERTURA CONEXION CON BASE DE DATOS " & Servidor + " " + ex.Message)
            Abrir_BD = False
        End Try
    End Function

    Public Function Cerrar_BD() As Boolean
        Cerrar_BD = True
        Try
            sqlConn.Close()
        Catch ex As SqlClient.SqlException
            Log("ERROR CERRAR CONEXION CON BASE DE DATOS" & Servidor + " " + ex.Message)
            Cerrar_BD = False
        End Try
    End Function

    Public Function Cargar_Dataset(xId As String) As Boolean
        Cargar_Dataset = Abrir_BD()

        If Cargar_Dataset Then
            Try
                Dim tx As String

                tx = " SELECT ST.[Documento Reclamación Humedad], ST.[Documento Reclamación Calidad],ST.[Documento Reclamación HU-CA]," +
                              "PH.[Buy-from Vendor No_], PH.[Buy-from Vendor Name], " +
                              "PL.[Informe No_], PL.[No_], PL.[Document No_], PL.[Line No_], PL.[Pesaje Id], PL.[Cantidad Báscula], PL.[Porcentaje Dto_ Báscula], PL.[Quantity]," +
                              "PS.[FechaPesada], PS.[Matricula], PS.[Albaran], PL.[Albaran Pesaje]," +
                              "VE.[Correo Incidencias] " +
                     " FROM [Papertech$Purchase Line] PL " +
                     " inner join [Papertech$Purchase Header] PH ON PH.[Document Type] = PL.[Document Type] AND PH.[No_] = PL.[Document No_]" +
                     " Inner join [Papertech$Vendor] VE ON VE.[No_] = PH.[Buy-from Vendor No_]" +
                     " Inner Join [Papertech$Pesaje] PS ON PS.[Id] = PL.[Pesaje Id] " +
                     " Inner Join [Papertech$Purchases & Payables Setup] ST ON ST.[Documento Reclamación Humedad] <> ''" +
                     " WHERE PL.[Document Type] = 1 AND PL.[Marca] = '" + xId + "'"
                Dim da = New SqlClient.SqlDataAdapter(tx, sqlConn)
                ds = New DataSet
                da.Fill(ds, "Informe")
                da.Dispose()
            Catch ex As Exception
                Log("Error Cargar Dataset " + ex.Message)
                MsgBox("Error Cargar Dataset " + ex.Message, MsgBoxStyle.Critical)
                Cargar_Dataset = False
            End Try
        End If

        If Cargar_Dataset Then Cargar_Dataset = Cerrar_BD()
    End Function

    Public Function Cargar_Dataset_Calidad(xDoc As String, xPesaje As Long) As Boolean
        Cargar_Dataset_Calidad = Abrir_BD()

        If Cargar_Dataset_Calidad Then
            Try
                Dim tx As String

                tx = " SELECT PL.[No_], PL.[Pesaje Id], PL.[Cantidad Báscula], PL.[Porcentaje Dto_ Báscula], PL.[Quantity]" + _
                     " FROM [Papertech$Purchase Line] PL " + _
                     " WHERE PL.[Document Type] = 1 AND PL.[Document No_] = '" + xDoc + "' AND [Pesaje Id] = " + xPesaje.ToString
                Dim da = New SqlClient.SqlDataAdapter(tx, sqlConn)
                If dsCalidad Is Nothing Then dsCalidad = New DataSet
                dsCalidad.Clear()
                da.Fill(dsCAlidad, "Informe")
                da.Dispose()
            Catch ex As Exception
                Log("Error Cargar Dataset " + ex.Message)
                MsgBox("Error Cargar Dataset " + ex.Message, MsgBoxStyle.Critical)
                Cargar_Dataset_Calidad = False
            End Try
        End If

        If Cargar_Dataset_Calidad Then Cargar_Dataset_Calidad = Cerrar_BD()
    End Function
    Public Function Cargar_Calidad(xOrder As String, xLine As Integer, xPesaje As Integer, ByRef xCantidad As Single, ByRef xRef As String) As Boolean
        Cargar_Calidad = Abrir_BD()

        xCantidad = 0
        xRef = ""

        If Cargar_Calidad Then
            Dim oComm As New SqlClient.SqlCommand
            Dim oRead As SqlClient.SqlDataReader

            Try
                oComm.Connection = sqlConn
                oComm.CommandText = "SELECT [Cantidad Báscula], [No_] FROM [Papertech$Purchase Line] PL " + _
                                    "  WHERE [Document No_] = '" + xOrder + "'" + _
                                    "    AND [Line No_] >" + xLine.ToString + _
                                    "    AND [Pesaje Id] = " + xPesaje.ToString
                oRead = oComm.ExecuteReader
                If oRead.Read Then
                    xCantidad = oRead.Item("Cantidad Báscula")
                    xRef = oRead.Item("No_")
                End If
                oRead.Close()
                oComm.Dispose()
            Catch ex As Exception
                Log("Error Cargar Calidad " + ex.Message)
                MsgBox("Error Cargar Calidad " + ex.Message, MsgBoxStyle.Critical)
                Cargar_Calidad = False
            End Try
        End If

        If Cargar_Calidad Then Cargar_Calidad = Cerrar_BD()
    End Function

End Module