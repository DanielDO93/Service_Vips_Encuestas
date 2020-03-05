
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Threading
Imports System.Web

Public Class VipsTrigger

    Private oTimer As System.Threading.Timer
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Agregue el código aquí para iniciar el servicio. Este método debería poner
        ' en movimiento los elementos para que el servicio pueda funcionar.
        If TimeOfDay >= #8:00:00 PM# Then

        Else
            SendEncuesta()
        End If

        Dim oCallback As New TimerCallback(AddressOf OnTimedEvent)
        ' oTimer = New System.Threading.Timer(oCallback, Nothing, 90000, 90000)
        oTimer = New System.Threading.Timer(oCallback, Nothing, 1800000, 1800000)

    End Sub

    Private Sub OnTimedEvent(ByVal state As Object)

        If TimeOfDay >= #8:00:00 PM# Then

        Else

            Try


                Dim conexion As New SqlConnection("Data Source=10.0.0.40; Initial Catalog=CCS;Persist Security Info=False;Application Name=Calidad; Pwd=PPl4t3nt0?;User ID=sa")
                Dim da As New System.Data.SqlClient.SqlDataAdapter
                Dim ds As New System.Data.DataSet
                '' Dim cmd As SqlCommand = New SqlCommand("SELECT * FROM CCS.dbo.Telefonos", conexion)

                'Este es el string Productivo

                Dim cmd As SqlCommand = New SqlCommand("SELECT DISTINCT tel_1 FROM CRM_VIPS.dbo.SYS_Interacciones where tip_02 = 'Nuevo Pedido' AND fecha_fin BETWEEN DATEADD(MINUTE, -180, GETDATE()) AND DATEADD(MINUTE, -150, GETDATE())", conexion)

                conexion.Open()
                cmd.CommandType = CommandType.Text
                da.SelectCommand = cmd
                ds.Tables.Add()
                da.Fill(ds.Tables(0))
                conexion.Close()


                For Each dr As DataRow In ds.Tables(0).Rows

                    Dim telefono As String = dr.Item(0).ToString()

                    Dim laUri As New Uri("http://api2.tkm.mx/api/aud?key=4bv35036w8uj1weu94bv030w95053&cli=577&psd=ccs4p1.2019&phone=" & telefono & "&id_enc=378")

                    Dim request As WebRequest = WebRequest.CreateDefault(laUri)
                    request.Method = "GET"
                    request.Headers("Authorization") = "Basic NTc3OmNjczRwMS4yMDE5"
                    Dim response As WebResponse

                    Try
                        response = request.GetResponse()
                    Catch exc As WebException
                        response = exc.Response
                    End Try

                    If response Is Nothing Then Throw New HttpException(CInt(HttpStatusCode.NotFound), "The requested url could not be found.")

                    Using reader As StreamReader = New StreamReader(response.GetResponseStream())
                        Dim requestedText As String = reader.ReadToEnd()
                    End Using

                Next
            Catch ex As Exception

            End Try

        End If





    End Sub
    Sub SendEncuesta()

        Dim laUri As New Uri("http://api2.tkm.mx/api/aud?key=4bv35036w8uj1weu94bv030w95053&cli=577&psd=ccs4p1.2019&phone=5511907300&id_enc=378")

        Dim request As WebRequest = WebRequest.CreateDefault(laUri)
        request.Method = "GET"
        request.Headers("Authorization") = "Basic NTc3OmNjczRwMS4yMDE5"
        Dim response As WebResponse

        Try
            response = request.GetResponse()
        Catch exc As WebException
            response = exc.Response
        End Try

        If response Is Nothing Then Throw New HttpException(CInt(HttpStatusCode.NotFound), "The requested url could not be found.")

        Using reader As StreamReader = New StreamReader(response.GetResponseStream())
            Dim requestedText As String = reader.ReadToEnd()
        End Using




        'Dim laUri2 As New Uri("http://api2.tkm.mx/api/aud?key=4bv35036w8uj1weu94bv030w95053&cli=577&psd=ccs4p1.2019&phone=5548506973&id_enc=378")

        'Dim request2 As WebRequest = WebRequest.CreateDefault(laUri2)
        'request2.Method = "GET"
        'request2.Headers("Authorization") = "Basic NTc3OmNjczRwMS4yMDE5"
        'Dim response2 As WebResponse

        'Try
        '    response2 = request2.GetResponse()
        'Catch exc As WebException
        '    response2 = exc.Response
        'End Try

        'If response2 Is Nothing Then Throw New HttpException(CInt(HttpStatusCode.NotFound), "The requested url could not be found.")

        'Using reader2 As StreamReader = New StreamReader(response2.GetResponseStream())
        '    Dim requestedText As String = reader2.ReadToEnd()
        'End Using


    End Sub
    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.
    End Sub

End Class
