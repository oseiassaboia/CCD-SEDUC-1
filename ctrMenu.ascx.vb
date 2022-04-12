
Partial Class ctrMenu
    Inherits System.Web.UI.UserControl

    'Private Sub btnEnviar_Click(sender As Object, e As EventArgs) Handles btnEnviar.Click
    '    'Utilidades.EnviarEmail("Teste de Email", "Testando o webmail da SEDUC", "igor.prohmann@gmail.com")
    'End Sub

    Private Sub ctrMenu_Init(sender As Object, e As EventArgs) Handles Me.Init
        If Not Page.IsPostBack Then
            If Not Session("CodigoUsuario") Is Nothing Then
                Dim Nome As String = (New ComponenteAcesso.Permissao).ObterNome(Session("CodigoUsuario"))

                If Nome <> "" Then
                    Dim splitNome() As String = Nome.Split(" ")

                    lblUsuario.Text = splitNome(0)
                    imgFoto.ImageUrl = "img/perfil_sombra.jpg"
                Else
                    lblUsuario.Text = "Não Identificado"
                End If
            Else
                lblUsuario.Text = "Visitante"
            End If

            imgFoto.ImageUrl = "frmPrincipalFotos.aspx?idPessoa=" & Session("CodigoPessoa").ToString
        End If
    End Sub

    Private Sub ctrMenu_Load(sender As Object, e As EventArgs) Handles Me.Load

        If Session("MenuSub") <> "" Then

        End If

        If Not Page.IsPostBack Then
            'imgFoto.ImageUrl = "frmPrincipalFotos.aspx?idPessoa=" & Session("CodigoPessoa").ToString
        End If

    End Sub

    Private Sub LimparSessao()


        Session("CodigoLotacaoMapeamento") = Nothing
        Session("CodigoPessoaCadastro") = Nothing
        Session("CodigoPessoaCargaHoraria") = Nothing
        Session("CodigoApresentacao") = Nothing
        Session("CodigoDisciplina") = Nothing
        Session("CodigoEtapa") = Nothing
        Session("CodigoMatricula") = Nothing
        Session("CodigoProfessor") = Nothing
        Session("CodigoTurma") = Nothing
        Session("CodigoMomento") = Nothing
        Session("CodigoAvaliacao") = Nothing
        Session("CodigoFrequencia") = Nothing
        Session("CodigoLotacaoCadastro") = Nothing
        Session("CodigoLicenca") = Nothing
        Session("CodigoServidorRecadastramento") = Nothing
        Session("CodigoServidorRecadastramentoAnual") = Nothing

        Session("CodigoPeriodoFerias") = Nothing
        Session("AnoRecadastramentoAnual") = Nothing

        Session("CodigoMapeamentoLotacao") = Nothing
        Session("CodigoMapeamentoServidor") = Nothing
        Session("CodigoMapeamentoServidorLotacao") = Nothing
        Session("CodigoMapeamentoPessoa") = Nothing

        Session("CodigoPeriodo") = Nothing

    End Sub

    'Area Ensino/CargaHoraria
    Private Sub lnkServidorCargaHoraria_Click(sender As Object, e As EventArgs) Handles lnkServidorCargaHoraria.Click
        LimparSessao()
        Session("MenuSub") = "liSubMenuApresentacao"
        Response.Redirect("frmResumoServidor.aspx")
    End Sub

End Class
