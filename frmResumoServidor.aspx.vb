Imports AjaxControlToolkit
Imports System.Data

Partial Class frmResumoServidor
    Inherits System.Web.UI.Page
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not Page.IsPostBack Then




        End If


    End Sub

#Region "Funções de Listagem"

    Private Sub CarregarImagem(ByVal Pessoa As Integer)
        Dim objDocumentoPessoa As New DocumentoPessoa

        dtlProfessor.DataSource = objDocumentoPessoa.Pesquisar(,, Pessoa)
        dtlProfessor.DataBind()

        objDocumentoPessoa = Nothing

    End Sub

    Private Sub CarregarNome(ByVal Pessoa As Integer)
        Dim objPessoa As New Pessoa

        Dim dt = objPessoa.PesquisarPerfil(Pessoa)

        If dt.Rows.Count > 0 Then
            Dim splitNome() As String = dt.Rows(0)("RH01_NM_PESSOA").Split(" ")
            lblNome.Text = splitNome(0) + " " + splitNome(1)
            lblCPF.Text = dt.Rows(0)("CPF")
            lblMatricula.Text = IIf(dt.Rows(0)("RH02_CD_MATRICULA") Is DBNull.Value, "", dt.Rows(0)("RH02_CD_MATRICULA"))
            lblCargo.Text = dt.Rows(0)("RH16_NM_CARGO")
            lblCargaHoraria.Text = IIf(dt.Rows(0)("RH79_NM_TIPO_CARGA_HORARIA") Is DBNull.Value, "", dt.Rows(0)("RH79_NM_TIPO_CARGA_HORARIA"))
            lblSituação.Text = IIf(dt.Rows(0)("RH07_NM_SITUACAO_SERVIDOR") Is DBNull.Value, "", dt.Rows(0)("RH07_NM_SITUACAO_SERVIDOR"))
            lblVinculo.Text = IIf(dt.Rows(0)("RH05_NM_TIPO_VINCULO") Is DBNull.Value, "", dt.Rows(0)("RH05_NM_TIPO_VINCULO"))
        End If

        objPessoa = Nothing
    End Sub

    Private Sub CarregarPerfilUsario(ByVal Usuario As Integer)
        Dim objUsuario As New Usuario

        Dim dt = objUsuario.ObterNomePerfil(Usuario)

        If dt.Rows.Count > 0 Then
            lblPerfil.Text = Left(dt.Rows(0)("CA03_DES_PERFIL").ToString, 15)
        End If

        objUsuario = Nothing
    End Sub

    Private Sub CarregarHabilidadePerfil(ByVal Pessoa As Integer)
        Dim objPessoa As New Pessoa

        grdPerfilHabilidades.DataSource = objPessoa.PesquisarHabilidade(Pessoa)
        grdPerfilHabilidades.DataBind()

        objPessoa = Nothing

    End Sub

    Private Sub CarregarLotacoesPerfil(ByVal Pessoa As Integer)
        Dim objPessoa As New Pessoa

        grdPerfilLotacoes.DataSource = objPessoa.PesquisarLotacoesMatricula(Pessoa)
        grdPerfilLotacoes.DataBind()

        objPessoa = Nothing

    End Sub

    Private Sub CarregarGrid(ByVal Pessoa As Integer)
        Dim objPessoa As New Pessoa

        Dim dsCargaHoraria As DataSet = New DataSet("Cargahoraria")
        dsCargaHoraria.Tables.Add(objPessoa.PesquisarAlocacaoCargaHoraria(Pessoa))

        Accordion1.DataSource = dsCargaHoraria.Tables(0).DefaultView
        Accordion1.DataBind()

        Accordion1.SelectedIndex = -1
        dsCargaHoraria = Nothing
        objPessoa = Nothing
    End Sub

    Protected Sub btnVoltar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVoltar.Click
        Session("CodigoPessoaCargaHoraria") = Nothing
        Response.Redirect(Request.Url.ToString)
    End Sub

#End Region

#Region "Eventos de Listagem"

    Private Sub Accordion1_ItemDataBound(sender As Object, e As AccordionItemEventArgs) Handles Accordion1.ItemDataBound
        If e.ItemType = AjaxControlToolkit.AccordionItemType.Content Then
            Dim CodigoAlocacao As Integer
            Dim objPessoa As New Pessoa


            CodigoAlocacao = DirectCast(e.AccordionItem.FindControl("hdnCodigoAlocacao"), HiddenField).Value


            Dim grd As New GridView()

            grd = DirectCast(e.AccordionItem.FindControl("grdhabilitacao"), GridView)

            grd.DataSource = objPessoa.PesquisarHabilitacoes(CodigoAlocacao)
            grd.DataBind()

            objPessoa = Nothing
        End If


    End Sub

#End Region

End Class
