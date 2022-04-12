
Partial Class MasterPageCargaHoraria
    Inherits System.Web.UI.MasterPage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not Page.IsPostBack Then

            If Session("CodigoUsuario") = 0 Or Session("CodigoUsuario") Is Nothing Then
                Session.Abandon()
                Response.Redirect("http://sistemas.educacao.ma.gov.br/acesso")
            End If

            If Session("CodigoPessoaCargaHoraria") Is Nothing Then
                Cadastro.Visible = False
                Listagem.Visible = True
            Else
                Cadastro.Visible = True
                Listagem.Visible = False
            End If

            ViewState("CodigoLotacao") = 0

            'Ajeita a aba
            Select Case Me.Page.ToString.ToLower

                Case "asp.frmresumoservidor_aspx"
                    lbtServidor.CssClass = "active"

                Case "asp.frmhabilidade_aspx"
                    lbthabilidade.CssClass = "active"

            End Select

        End If

        validacao.Outros(txtCpf,False,"CPF",,Validacao.eFormato.CPF)
    End Sub



    Protected Sub lbtServidor_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbtServidor.Click
        Response.Redirect("frmResumoServidor.aspx")
    End Sub

    Protected Sub lbthabilidade_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbthabilidade.Click
        If Session("CodigoPessoaCargaHoraria") Is Nothing Then
            MsgBox("Selecione primeiro um pessoa!")
        Else
            Response.Redirect("frmHabilidade.aspx")
        End If

    End Sub


#Region "Funções de Listagem"

    Private Sub CarregarGrid()
        Dim objPessoa As New Pessoa

        grdPessoa.DataSource = objPessoa.Pesquisar(ViewState("OrderBy"),, Replace(Replace(txtCpf.Text, ".", ""), "-", ""), txtLocalizar.Text,,,,,,,,,,,,,,,)
        grdPessoa.DataBind()

        objPessoa = Nothing

        lblRegistros.Text = DirectCast(grdPessoa.DataSource, Data.DataTable).Rows.Count & " registro(s)"
    End Sub

#End Region

#Region "Eventos de Listagem"

    Protected Sub btnCadastrar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCadastrar.Click
        Cadastro.Visible = True
        Listagem.Visible = False
    End Sub

    Protected Sub btnLocalizar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLocalizar.Click
        CarregarGrid()
    End Sub


    Protected Sub grdPessoa_RowCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles grdPessoa.RowCommand
        If e.CommandName = "" Then
            Session("CodigoPessoaCargaHoraria") = grdPessoa.DataKeys(e.CommandArgument).Item(0)
            Response.Redirect(Request.Url.ToString)

        End If

    End Sub

    Private Sub grdPessoa_PageIndexChanging(ByVal source As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles grdPessoa.PageIndexChanging
        grdPessoa.PageIndex = e.NewPageIndex
        CarregarGrid()
    End Sub

#End Region
End Class

