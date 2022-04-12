
Imports System.Data

Partial Class Habilidade
    Inherits System.Web.UI.Page

    Private Sub Habilidade_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then

            CarregarGrid()

        End If


    End Sub

    Private Sub CarregarGrid()
        Dim objHabilidade As New HabilidadeServidor

        grdHabilidade.DataSource = objHabilidade.Pesquisar(ViewState("OrderBy"),, ,,, ,,,,, Session("CodigoPessoaCargaHoraria"), False)
        grdHabilidade.DataBind()

        objHabilidade = Nothing

        lblRegistros.Text = DirectCast(grdHabilidade.DataSource, Data.DataTable).Rows.Count & " Registro(s)"

    End Sub



    Private Sub limparcampos()
        ViewState("CodigoHabilidadeServidor") = Nothing
    End Sub


    Private Sub grdHabilidade_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles grdHabilidade.RowDataBound
        Dim lnkCancelar As LinkButton
        Dim iCancelar As New Object
        Dim lblCancelar As New Label

        If e.Row.RowType = DataControlRowType.DataRow Then

            lnkCancelar = DirectCast(e.Row.Cells(3).FindControl("lnkCancelar"), LinkButton)
            iCancelar = DirectCast(e.Row.Cells(3).FindControl("iCancelar"), Object)
            lblCancelar = DirectCast(e.Row.Cells(3).FindControl("lblCancelar"), Label)
            lnkCancelar.CommandArgument = e.Row.RowIndex

            'Dim rowView As DataRowView = DirectCast(e.Row.DataItem, DataRowView)

            If grdHabilidade.DataKeys(e.Row.RowIndex).Item(2) = 1 Then

                iCancelar.Attributes.Add("class", "fa fa-exchange")
                lnkCancelar.Attributes.Add("class", "badge bg-green btn-block")
                lnkCancelar.CommandName = "Desvio"
                lblCancelar.Text = " Desvio"

            Else
                e.Row.ForeColor = Drawing.Color.DarkRed
                lblCancelar.Text = " "
            End If

        End If

        lblCancelar = Nothing
        lnkCancelar = Nothing
        iCancelar = Nothing
    End Sub

    Private Sub grdHabilidade_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles grdHabilidade.RowCommand
        If e.CommandName = "Desvio" Then
            MsgModal(2, "Deseja alterar o tipo de habilidade?")
            btnEntendi.Text = "Não"
            btnHabilidade.Text = "Sim"
            btnHabilidade.Visible = True
            ViewState("TipoHabilidade") = grdHabilidade.DataKeys(e.CommandArgument).Item(0)
        End If


        If e.CommandName = "" Then

        End If
    End Sub

    Private Sub MsgModal(Optional TipoMsg As Integer = 1, Optional Mensagem As String = "", Optional Mensagem2 As String = "")
        'Parametros de tipo (1 = sucesso | 2 = aviso | 3 = erro)
        If TipoMsg = 1 Then
            ScriptManager.RegisterStartupScript(Me.Page, Me.Page.GetType, "openModal", "openModal('sucesso');", True)
        ElseIf TipoMsg = 2 Then
            ScriptManager.RegisterStartupScript(Me.Page, Me.Page.GetType, "openModal", "openModal('aviso');", True)
        ElseIf TipoMsg = 3 Then
            ScriptManager.RegisterStartupScript(Me.Page, Me.Page.GetType, "openModal", "openModal('erro');", True)
        End If
        lblMensagem.Text = Mensagem
        lblMensagem2.Text = Mensagem2
    End Sub

    Private Sub btnHabilidade_Click(sender As Object, e As EventArgs) Handles btnHabilidade.Click
        TrocarTipoHabilidade(ViewState("TipoHabilidade"))
        limparcampos()
        CarregarGrid()
    End Sub

    Private Sub grdHabilidade_PageIndexChanging(ByVal source As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles grdHabilidade.PageIndexChanging
        grdHabilidade.PageIndex = e.NewPageIndex
        CarregarGrid()
    End Sub

    Private Sub grdHabilidade_Sorting(ByVal source As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles grdHabilidade.Sorting
        ViewState("OrderByDirection") = IIf(ViewState("OrderByDirection") = "asc", "desc", "asc")
        ViewState("OrderBy") = e.SortExpression & " " & ViewState("OrderByDirection")
        CarregarGrid()
    End Sub

    Private Sub TrocarTipoHabilidade(ByVal CodigoHabilidade As Integer)
        Dim objHabilidade As New HabilidadeServidor(CodigoHabilidade)

        With objHabilidade

            .DataHoraDesativacao = Nothing
            .TpHabilidade = 2

        End With

        Try
            'objHabilidade.Salvar()
        Catch ex As Exception
            Dim erro As String = ex.ToString
        End Try

        objHabilidade = Nothing

    End Sub

End Class
