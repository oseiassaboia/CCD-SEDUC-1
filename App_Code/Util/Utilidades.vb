Imports System.IO
Imports System.Xml.Serialization
Imports System.Data
Imports System.Net
Imports CepFacil.API

Public Class Utilidades

    Public Shared Function CPF_Valido(ByVal CPF As String) As Boolean
        Dim dadosArray() As String = {"111.111.111-11", "222.222.222-22", "333.333.333-33", "444.444.444-44", "555.555.555-55", "666.666.666-66", "777.777.777-77", "888.888.888-88", "999.999.999-99"}
        Dim i, x, n1, n2 As Integer

        CPF = CPF.Trim

        For i = 0 To dadosArray.Length - 1
            If CPF.Length <> 14 Or dadosArray(i).Equals(CPF) Then
                Return False
            End If
        Next

        'remove a maskara
        CPF = CPF.Substring(0, 3) + CPF.Substring(4, 3) + CPF.Substring(8, 3) + CPF.Substring(12)

        For x = 0 To 1
            n1 = 0

            For i = 0 To 8 + x
                n1 = n1 + Val(CPF.Substring(i, 1)) * (10 + x - i)
            Next

            n2 = 11 - (n1 - (Int(n1 / 11) * 11))

            If n2 = 10 Or n2 = 11 Then n2 = 0

            If n2 <> Val(CPF.Substring(9 + x, 1)) Then
                Return False
            End If
        Next

        Return True

    End Function

    Public Shared Function DataCompleta(ByVal Data As Date) As String
        Dim strMes As String

        Select Case Month(Data)
            Case 1
                strMes = "Janeiro"
            Case 2
                strMes = "Fevereiro"
            Case 3
                strMes = "Março"
            Case 4
                strMes = "Abril"
            Case 5
                strMes = "Maio"
            Case 6
                strMes = "Junho"
            Case 7
                strMes = "Julho"
            Case 8
                strMes = "Agosto"
            Case 9
                strMes = "Setembro"
            Case 10
                strMes = "Outubro"
            Case 11
                strMes = "Novembro"
            Case 12
                strMes = "Dezembro"
            Case Else

                strMes = ""

        End Select

        Return Day(Data) & " de " & strMes & " de " & Year(Data)

    End Function

    Public Shared Function Meses(Optional ByVal PrimeiroItem As String = "") As Data.DataTable
        Dim dtMeses As New Data.DataTable
        Dim drMeses As Data.DataRow

        dtMeses.Columns.Add("CODIGO")
        dtMeses.Columns.Add("DESCRICAO")

        drMeses = dtMeses.NewRow
        drMeses("CODIGO") = ""
        drMeses("DESCRICAO") = PrimeiroItem
        dtMeses.Rows.Add(drMeses)

        For x = 1 To 12
            drMeses = dtMeses.NewRow
            drMeses("CODIGO") = x
            Select Case x
                Case 1
                    drMeses("DESCRICAO") = "JANEIRO"
                Case 2
                    drMeses("DESCRICAO") = "FEVEREIRO"
                Case 3
                    drMeses("DESCRICAO") = "MARÇO"
                Case 4
                    drMeses("DESCRICAO") = "ABRIL"
                Case 5
                    drMeses("DESCRICAO") = "MAIO"
                Case 6
                    drMeses("DESCRICAO") = "JUNHO"
                Case 7
                    drMeses("DESCRICAO") = "JULHO"
                Case 8
                    drMeses("DESCRICAO") = "AGOSTO"
                Case 9
                    drMeses("DESCRICAO") = "SETEMBRO"
                Case 10
                    drMeses("DESCRICAO") = "OUTUBRO"
                Case 11
                    drMeses("DESCRICAO") = "NOVEMBRO"
                Case 12
                    drMeses("DESCRICAO") = "DEZEMBRO"
            End Select
            dtMeses.Rows.Add(drMeses)
        Next

        Return dtMeses
    End Function

    Public Shared Sub DownloadFile(ByVal CaminhoOrigem As String, ByVal NomeDestino As String)

        Dim executingPage As Object = HttpContext.Current.Handler
        Dim fullpath As String = System.IO.Path.GetFullPath(CaminhoOrigem)
        Dim name As String = System.IO.Path.GetFileName(fullpath)
        Dim ext As String = System.IO.Path.GetExtension(fullpath).ToLower
        Dim Type As String = ""

        Select Case ext
            Case ".htm", ".html"
                Type = "text/HTML"
            Case ".txt"
                Type = "text/plain"
            Case ".doc", ".rtf"
                Type = "Application/msword"
            Case ".csv", ".xls"
                Type = "Application/x-msexcel"
            Case ".pdf"
                Type = "application/pdf"
            Case Else
                Type = "text/plain"
        End Select

        Try
            executingPage.Response.AppendHeader("content-disposition", "attachment; filename=" + NomeDestino)
            executingPage.Response.ContentType = Type
            executingPage.Response.WriteFile(fullpath)
            executingPage.Response.End()

        Catch
            MsgBox("Arquivo não Encontrado!")
        End Try

    End Sub

    Public Shared Function SenhaRandomica(Optional ByVal Tamanho As Integer = 6, Optional ByVal UsarLetras As Boolean = True, Optional ByVal UsarNumeros As Boolean = True) As String
        'a - z = 97 - 122
        '0 - 9 = 48 - 57
        Dim x As String

        SenhaRandomica = ""

        If UsarLetras = False And UsarNumeros = False Then
            Return SenhaRandomica
        End If

        While Len(SenhaRandomica) < Tamanho
            Randomize()
            'Int((Rnd * (Maximo * 1000) - 1) / 1000) + 1
            x = CInt(Rnd() * 1000)
            If (x >= 48 And x <= 57 And UsarNumeros = True) Or (x >= 97 And x <= 122 And UsarLetras = True) Then
                SenhaRandomica = SenhaRandomica & Chr(x)
            End If
        End While
    End Function

    Public Shared Function ObterAtributoStringConexao(ByVal StringConexao As String, ByVal Atributo As String) As String
        Return Mid(Mid(StringConexao, InStr(StringConexao, Atributo, CompareMethod.Text), InStr(InStr(StringConexao, Atributo, CompareMethod.Text), StringConexao, ";", CompareMethod.Text) - InStr(StringConexao, Atributo, CompareMethod.Text)), Len(Atributo) + 2)
    End Function

    Public Shared Function Encript(ByVal VarTxt$) As String
        Dim TxtCript$, i%

        TxtCript = ""

        For i = 1 To Len(VarTxt)
            TxtCript = TxtCript & Chr(Asc(Mid$(VarTxt, i, 1)) + 130 - i)
        Next

        Encript = TxtCript
    End Function

    Public Shared Function Decript(ByVal VarTxt$) As String
        Dim TxtCript$, i%

        On Error Resume Next

        TxtCript = ""

        For i = 0 To Len(VarTxt)
            TxtCript = TxtCript & Chr(Asc(Mid$(VarTxt, Len(VarTxt) - i, 1)) - 130 + Len(VarTxt) - i)
        Next

        Decript = StrReverse(TxtCript)

    End Function

    Public Shared Function EnviarEmail(ByVal Assunto As String, ByVal Mensagem As String, ByVal Destinatario As String, Optional ByVal Anexo As String = "") As Boolean
        Dim Destinatarios(0) As Net.Mail.MailAddress
        Dim Anexos(0) As Net.Mail.Attachment

        Destinatarios(0) = New Net.Mail.MailAddress(Destinatario)
        If Anexo <> "" Then
            Anexos(0) = New Net.Mail.Attachment(Anexo)
        End If

        Return EnviarEmail(Assunto, Mensagem, Destinatarios, Anexos)
    End Function

    Public Shared Function EnviarEmail(ByVal Assunto As String, ByVal Mensagem As String, ByVal Destinatarios() As Net.Mail.MailAddress, Optional ByVal Anexos() As Net.Mail.Attachment = Nothing) As Boolean
        Dim objEmail As New System.Net.Mail.MailMessage

        Try
            objEmail.From = New System.Net.Mail.MailAddress("naoresponder@seati.ma.gov.br", "Seletivo SEDUC")

            For Each Destinatario As System.Net.Mail.MailAddress In Destinatarios
                If Not Destinatario Is Nothing Then
                    objEmail.To.Add(Destinatario)
                End If
            Next

            If Not Anexos Is Nothing Then
                For Each Anexo As Net.Mail.Attachment In Anexos
                    If Not Anexo Is Nothing Then
                        objEmail.Attachments.Add(Anexo)
                    End If
                Next
            End If

            objEmail.Priority = System.Net.Mail.MailPriority.Normal
            objEmail.IsBodyHtml = True

            objEmail.Subject = Assunto
            objEmail.Body = Mensagem

            objEmail.SubjectEncoding = System.Text.Encoding.GetEncoding("ISO-8859-1")
            objEmail.BodyEncoding = System.Text.Encoding.GetEncoding("ISO-8859-1")

            Dim objSmtp As New System.Net.Mail.SmtpClient

            Dim basicAuthenticationInfo As New System.Net.NetworkCredential("naoresponder@seati.ma.gov.br", "seletivo2016")

            objSmtp.Host = "correio.ma.gov.br"
            objSmtp.Port = "25" '995
            objSmtp.UseDefaultCredentials = False
            objSmtp.Credentials = basicAuthenticationInfo
            objSmtp.EnableSsl = True

            objSmtp.Send(objEmail)

            objEmail.Dispose()

            objEmail = Nothing

            Return True

        Catch ex As Exception
            Return False

        End Try

    End Function

    Public Shared Function Abreviado(ByVal Palavra As String) As Boolean
        Dim Retorno As Boolean = False
        Dim i As Integer

        For i = 0 To Palavra.Length - 1
            If Palavra(i) = "." Then
                Retorno = True
            End If
            If i <> 0 And i <> (Palavra.Length - 1) Then
                If Palavra(i - 1) = " " And Palavra(i + 1) = " " Then
                    Retorno = True
                End If
            ElseIf i = (Palavra.Length - 1) Then
                If Palavra(i - 1) = " " Then
                    Retorno = True
                End If
            ElseIf i = 0 Then
                If Palavra(i + 1) = " " Then
                    Retorno = True
                End If
            End If

        Next

        Return Retorno
    End Function

    Public Shared Function NewID() As String
        Dim Retorno As String = ""
        Dim cnn As New Conexao

        Try
            Retorno = cnn.AbrirDataTable("select NEWID() ").Rows(0)(0).ToString
        Catch ex As Exception

        End Try

        cnn = Nothing
        Return Retorno
    End Function

    Public Shared Function Serializar(ByVal algumObjeto As Object) As String
        Dim writer As StringWriter = New StringWriter()
        Dim serializer As XmlSerializer = New XmlSerializer(algumObjeto.[GetType]())

        serializer.Serialize(writer, algumObjeto)

        Return writer.ToString()
    End Function

    Public Function BuscarCep(ByVal Cep As String) As Data.DataTable
        Dim ds As New DataSet()
        Dim dt As New DataTable

        'Fonte de varios webservices de CEP http://www.pinceladasdaweb.com.br/blog/2013/02/15/apis-para-consulta-de-cep/

        'Dim localidade As Localidade = CepFacilBusca.BuscarLocalidade("13214080", "SEU CÓDIGO DE FILIAÇÃO")
        'Dim Endereco As String = RequestDadosWeb("http://api.postmon.com.br/v1/cep/" & Replace(Replace(Cep, ".", ""), "-", "") & "?format=xml")
        Dim Endereco As String = RequestDadosWeb("http://cep.republicavirtual.com.br/web_cep.php?cep=" & Replace(Replace(Cep, ".", ""), "-", "") & "&formato=xml")
        'Dim Endereco As String = RequestDadosWeb("http://correios.w2info.com.br/Service.svc/Buscar?cep=" & Replace(Replace(Cep, ".", ""), "-", ""))
        'Dim Endereco As String = RequestDadosWeb("http://cep.osgestor.info/" & Replace(Replace(Cep, ".", ""), "-", "") & ".xml")

        'Endereco = Mid$(Endereco, 11)
        'Endereco = Left(Endereco, Endereco.Length - 11)

        'Endereco = "<enderecos>" + Endereco + "</enderecos>"

        ds.ReadXml(New StringReader(Endereco))

        dt = ds.Tables(0)

        Return dt
    End Function

    Public Shared Function ImageResize(ByVal image As System.Drawing.Image, ByVal width As Int32, ByVal height As Int32) As System.Drawing.Image
        Dim bitmap As System.Drawing.Bitmap = New System.Drawing.Bitmap(width, height, image.PixelFormat)
        If bitmap.PixelFormat = Drawing.Imaging.PixelFormat.Format1bppIndexed Or
           bitmap.PixelFormat = Drawing.Imaging.PixelFormat.Format4bppIndexed Or
           bitmap.PixelFormat = Drawing.Imaging.PixelFormat.Format8bppIndexed Or
           bitmap.PixelFormat = Drawing.Imaging.PixelFormat.Undefined Or
           bitmap.PixelFormat = Drawing.Imaging.PixelFormat.DontCare Or
           bitmap.PixelFormat = Drawing.Imaging.PixelFormat.Format16bppArgb1555 Or
           bitmap.PixelFormat = Drawing.Imaging.PixelFormat.Format16bppGrayScale Then
            'More Info http://msdn.microsoft.com/library/default.asp?_   
            'url=/library/en-us/cpref/html/frlrfSystemDrawingGraphicsClassFromImageTopic.asp   
            Throw New NotSupportedException("Pixel format of the image is not supported.")
        End If
        Dim graphicsImage As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(bitmap)
        graphicsImage.SmoothingMode = Drawing.Drawing2D.SmoothingMode.HighQuality
        graphicsImage.InterpolationMode = Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
        graphicsImage.DrawImage(image, 0, 0, bitmap.Width, bitmap.Height)
        graphicsImage.Dispose()
        Return bitmap
    End Function
End Class