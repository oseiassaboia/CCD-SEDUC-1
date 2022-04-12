
Imports System.Data

Partial Class Habilidade
    Inherits System.Web.UI.Page

    Private Sub Habilidade_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then

            ViewState("Permissao") = (New ComponenteAcesso.Permissao).Funcionalidade(Session("CodigoUsuario"), 1602)
            If ViewState("Permissao") = 1 Then
                pnlCadastro.Enabled = False
                grdHabilidade.Enabled = False
                btnSalvar.Enabled=False

            ElseIf ViewState("Permissao") = 2 Then
                pnlCadastro.Enabled = True
               
                btnSalvar.Enabled=true
            Else
                Response.Redirect("frmPrincipal.aspx")
            End If


            If Session("CodigoUsuario") Is Nothing Then
                Response.Redirect("frmPrincipal.aspx")
            End If


            'CarregarComboTabela(drpModalidade, New Modalidade, "...")
            CarregarComboServidor()
            CarregarComboTabela(drpDisciplina, New Disciplina)

            CarregarGrid()

        End If

        Validacao.Combo(drpServidor, True, "Servidor")
         'Validacao.Combo(drpEtapa, True, "Etapa")
        Validacao.Combo(drpDisciplina, True, "Disciplica")
        Validacao.Combo(drpTipodeHabilidade, True, "Tipo de Habilidade")
        Validacao.Finalizar(btnSalvar,, True)
    End Sub

    Private Sub CarregarGrid()
        Dim objHabilidade As New HabilidadeServidor

        grdHabilidade.DataSource = objHabilidade.Pesquisar(ViewState("OrderBy"),, drpServidor.SelectedValue,, drpDisciplina.SelectedValue, drpTipodeHabilidade.SelectedValue,,,,, Session("CodigoPessoaCargaHoraria"), False)
        grdHabilidade.DataBind()

        objHabilidade = Nothing

        lblRegistros.Text = DirectCast(grdHabilidade.DataSource, Data.DataTable).Rows.Count & " Registro(s)"

    End Sub

    Private Sub CarregarComboServidor()
        Dim objServidor As New Servidor

        With drpServidor
            .Items.Clear()

            .DataValueField = "RH02_ID_SERVIDOR"
            .DataTextField = "DESCRICAO"

            .DataSource = objServidor.Pesquisar(,, (Session("CodigoPessoaCargaHoraria")),,,,,, Servidor.Situacao.ATIVO)
            .DataBind()

            .Items.Insert(0, New ListItem("", 0))
        End With

        objServidor = Nothing
    End Sub

    'Private Sub drpModalidade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles drpModalidade.SelectedIndexChanged
    '    If drpModalidade.SelectedValue > 0 Then
    '        CarregarComboTabelaRelacionada(drpNivel, New Nivel, drpModalidade.SelectedValue, "...")
    '        drpNivel.Enabled = True
    '    Else
    '        drpNivel.Enabled = False
    '        drpEtapa.Enabled = False
    '        drpNivel.ClearSelection()
    '    End If

    'End Sub
    'Private Sub drpNivel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles drpNivel.SelectedIndexChanged

    '    If drpNivel.SelectedValue > 0 Then
    '        CarregarComboTabelaRelacionada(drpEtapa, New Etapa, drpNivel.SelectedValue, "...")
    '        drpEtapa.Enabled = True
    '    Else
    '        drpEtapa.Enabled = False
    '        drpEtapa.ClearSelection()
    '    End If

    'End Sub

    Private Sub btnNovo_Click(sender As Object, e As EventArgs) Handles btnNovo.Click
        limparcampos()
    End Sub

    Private Sub limparcampos()
        drpServidor.ClearSelection()
        drpServidor.Enabled = True
        'drpModalidade.ClearSelection()
        'drpNivel.ClearSelection()
        'drpEtapa.ClearSelection()
        drpDisciplina.ClearSelection()
        drpDisciplina.Enabled = True
        drpTipodeHabilidade.ClearSelection()
        ViewState("CodigoHabilidadeServidor") = Nothing


    End Sub

    Private Sub btnSalvar_Click(sender As Object, e As EventArgs) Handles btnSalvar.Click
        try
            Salvar()
            limparcampos()
            CarregarGrid()
            MsgBox(eTipoMensagem.SALVAR_SUCESSO)
        Catch ex As Exception
            MsgBox(eTipoMensagem.SALVAR_erro)
        End Try

    End Sub

    Private Sub Salvar()

        If (ViewState("CodigoHabilidadeServidor") Is Nothing) Then
            With (New HabilidadeServidor).Pesquisar(,, drpServidor.SelectedValue, , drpDisciplina.SelectedValue, drpTipodeHabilidade.SelectedValue,,,,,,)
                If .Rows.Count > 0 Then
                    MsgBox("Já Existe um registro ativo para os itens selecionados. Por favor, edite o registro existente")
                    Exit Sub
                End If
            End With
            With (New HabilidadeServidor).Pesquisar(,, drpServidor.SelectedValue, , drpDisciplina.SelectedValue, 2,,,,,, False)
                If .Rows.Count > 0 Then
                    MsgBox("Já Existe um registro para os itens selecionados. Por favor, edite o registro existente")
                    Exit Sub
                End If
            End With
        End If


        With (New HabilidadeServidor).Pesquisar(,, drpServidor.SelectedValue,,,,,,,,,)
            'CASO SEJA A PRIMEIRA HABILIDADE, ELA NÃO PODE SER DESVIO (É OBRIGATÓRIO QUE AO MENOS UMA DISCIPLINA SEJA COMUM)
            If .Rows.Count = 0 And drpTipodeHabilidade.SelectedValue = 2 Then
                MsgBox("Pelo menos uma habilidade precisa ser do tipo Comum")
                Exit Sub
            End If
        End With

        ' CASO EDITANDO UMA HABILIDADE
        If Not (ViewState("CodigoHabilidadeServidor") Is Nothing) Then

            'VERIFICO AS HABILIDADES DO SERVIDOR
            With (New HabilidadeServidor).Pesquisar(,, drpServidor.SelectedValue,,,,,,,,,)

                'CASO O MESMO JÁ POSSUA HABILIDADES, VERIFICO SE O MESMO ESTÁ TENTANDO EDITAR PARA UMA TIPO "DESVIO"
                If .Rows.Count > 0 And drpTipodeHabilidade.SelectedValue = 2 Then

                    'ADICIONO UM CONTADOR PARA A QUANTIDADE DE HABILIDADES TIPO "DESVIO", INCLUINDO A NOVA QUE O MESMO ESTÁ TENTANDO ADICIONAR
                    Dim desvioCount As Integer = 1

                    'ITERO PELOS REGISTROS VERIFICANDO A QUANTIDADE DE HABILIDADES "DESVIO"
                    For Each row As DataRow In .Rows
                        If row("RH72_TP_HABILIDADE") = 2 Then

                            'SE FOR DO TIPO DESVIO, ADICIONO NO CONTADOR
                            desvioCount = desvioCount + 1
                        End If
                    Next

                    'VERIFICO SE A QUANTIDADE DE REGISTROS, INCLUINDO O NOVO, DO TIPO DESVIO
                    'SE TODOS OS VALORES FICARÃO SENDO DO TIPO "DESVIO", RETORNO UMA MENSAGEM POIS É NECESSÁRIO QUE PELO MENOS UMA HABILIDADE SEJA DO TIPO COMUM
                    If desvioCount = .Rows.Count Then
                        MsgBox("Pelo menos uma habilidade precisa ser do tipo COMUM")
                        Exit Sub
                    End If
                End If
            End With
        End If




        If drpTipodeHabilidade.SelectedValue = 1 Then
            With (New HabilidadeServidor).Pesquisar(,, drpServidor.SelectedValue,,, drpTipodeHabilidade.SelectedValue)
                If .Rows.Count > 0 Then
                    MsgBox("Habilidade COMUM já cadastrada!")
                    Exit Sub
                End If
            End With
        End If

        Dim objHabilidade As New HabilidadeServidor(ViewState("CodigoHabilidadeServidor"))

        Using objHabilidade

            objHabilidade.IdServidor = drpServidor.SelectedValue
            'objHabilidade.IdEtapa = drpEtapa.SelectedValue
            objHabilidade.IdDisciplina = drpDisciplina.SelectedValue
            objHabilidade.TpHabilidade = drpTipodeHabilidade.SelectedValue

            objHabilidade.IdUsuario = Session("CodigoUsuario")
            objHabilidade.DataHoraCadastro = Date.Now

            If Not ViewState("CodigoHabilidadeServidor") = Nothing Then
                With New HabilidadeServidor().Pesquisar(, ViewState("CodigoHabilidadeServidor"))
                    If .Rows.Count > 0 Then
                        objHabilidade.IdDisciplina = .Rows(0)("DE09_ID_DISCIPLINA")
                        objHabilidade.IdServidor = .Rows(0)("RH02_ID_SERVIDOR")
                    End If

                End With

            End If

            objHabilidade.Salvar()

            MsgBox(eTipoMensagem.SALVAR_SUCESSO)

        End Using
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
            TrocarTipoHabilidade(grdHabilidade.DataKeys(e.CommandArgument).Item(0))
            limparcampos()
            CarregarGrid()
        End If


        If e.CommandName = "" Then

        End If
    End Sub

    Private Sub CarregarHabilidadeServidor(Codigo As Integer)
        Dim objHabilitacao As New HabilidadeServidor(Codigo)
        drpServidor.SelectedValue = objHabilitacao.IdServidor
        drpServidor.Enabled = False
        drpDisciplina.SelectedValue = objHabilitacao.IdDisciplina
        drpDisciplina.Enabled = False
        drpTipodeHabilidade.SelectedValue = objHabilitacao.TpHabilidade
        ViewState("CodigoHabilidadeServidor") = Codigo

    End Sub

    Private Sub HabilitarHabilidade(Codigo As Integer)
        Dim objHabilidade As New HabilidadeServidor(Codigo)

        With (New HabilidadeServidor).Pesquisar(,, objHabilidade.IdServidor,,, 1,,,,,,)

            'VERIFICO SE A MATRICULA POSSUI DISCIPLINA COMUM E VERIFICANDO SE A HABILIDADE QUE ESTÁ ATIVANDO É COMUM
            If .Rows.Count > 0 And objHabilidade.TpHabilidade = 1 Then

                MsgBox("Somente uma habilidade pode ser do tipo comum")
                Exit Sub

            End If
        End With

        With objHabilidade

            .IdUsuarioAlteracao = Session("CodigoUsuario")
            .DataHoraDesativacao = ""

            .Salvar()

        End With

        objHabilidade = Nothing
    End Sub


    Private Sub DesabilitarHabilitacao(Codigo As integer)
       Dim objhabilitacao as New habilitacao

        objhabilitacao.DesabilitarPorHabilidade(Codigo,Session("CodigoUsuario"))

        objhabilitacao = Nothing
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
            objHabilidade.Salvar()
        Catch ex As Exception
            Dim erro As String = ex.ToString
        End Try

        objHabilidade = Nothing

    End Sub

    Private Sub btnLocalizar_Click(sender As Object, e As EventArgs) Handles btnLocalizar.Click
        CarregarGrid()
    End Sub
End Class
