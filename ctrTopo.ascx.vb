
Partial Class ctrTopo
    Inherits System.Web.UI.UserControl
    Dim blnTestar As Boolean = CBool(System.Configuration.ConfigurationManager.AppSettings("Teste"))
    Dim Manutencao As Boolean = CBool(System.Configuration.ConfigurationManager.AppSettings("Manutencao"))
    Dim AmbienteTeste As String = System.Configuration.ConfigurationManager.AppSettings("AmbienteTeste").ToString
    Dim UsuarioTeste As String = System.Configuration.ConfigurationManager.AppSettings("UsuarioTeste").ToString

    Protected Sub ctrTopo_Init(sender As Object, e As EventArgs) Handles Me.Init

        If Not Page.IsPostBack Then

            If Session("CodigoUsuario") Is Nothing Then
                'Session.Abandon()


            End If

            If Not Session("CodigoUsuario") Is Nothing Then
                Dim Nome As String = ""
                Dim objUsuario As New Usuario(Session("CodigoUsuario"))

                If Manutencao And Not objUsuario.Programador Then
                    Response.Redirect("frmManutencao.aspx")
                End If

                If Nome <> "" Then
                    Dim splitNome() As String = Nome.Split(" ")

                    If splitNome.Length = 1 Then
                        'lblUsuario.Text = splitNome(0)
                        'lblNomeUsuario.Text = splitNome(0)
                    Else
                        'lblUsuario.Text = splitNome(0) + " " + splitNome(1)
                        'lblNomeUsuario.Text = splitNome(0) + " " + splitNome(1)
                    End If

                    'imgFoto.ImageUrl = "img/perfil_sombra.jpg"
                Else
                    'lblUsuario.Text = "Não Identificado"
                End If

            End If



        End If

    End Sub
    Private Sub ctrTopo_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then

        End If
    End Sub

    'Private Sub lnkSair_Click(sender As Object, e As EventArgs) Handles lnkSair.Click
    '    Session.Abandon()
    'End Sub
End Class
