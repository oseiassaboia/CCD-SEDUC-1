Imports System.Drawing
Imports System.Drawing.Text
Imports System.Drawing.Imaging

Partial Class Captcha
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        GerarCaptcha()
    End Sub

    Private Sub GerarCaptcha()
        Dim objBMP As Bitmap = New System.Drawing.Bitmap(150, 40)

        Dim objGraphics As Graphics = System.Drawing.Graphics.FromImage(objBMP)

        objGraphics.Clear(Color.SkyBlue)

        objGraphics.TextRenderingHint = TextRenderingHint.AntiAlias

        ' Fonte configurada para ser usada no texto do captcha

        Dim objFont As New Font("Times New Roman", 24, FontStyle.Strikeout)

        Dim captchaValue As String = ""

        Dim valuesArray As Integer() = New Integer(7) {}

        Dim x As Integer

        'Cria o valor randomicamente e adiciona ao array

        Dim autoRand As New Random()

        For x = 0 To 7

            valuesArray(x) = System.Convert.ToInt32(autoRand.[Next](0, 9))

            captchaValue += (valuesArray(x).ToString())
        Next

        'Adiciona o valor gerado para o captcha na sessão
        'para ser validado posteriormente

        Session.Add("CaptchaValue", captchaValue)

        'Desenha a imagem com o nosso texto

        objGraphics.DrawString(captchaValue, objFont, Brushes.White, 3, 3)

        'Determina o tipo de conteúdo da imagem do captcha

        Response.ContentType = "image/GIF"

        'Salva em stream

        objBMP.Save(Response.OutputStream, ImageFormat.Gif)

        'Libera os objeto da memória pois os mesmos não são mais necessários

        objFont.Dispose()

        objGraphics.Dispose()

        objBMP.Dispose()

    End Sub

End Class
