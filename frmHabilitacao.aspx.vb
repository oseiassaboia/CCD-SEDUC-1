
Imports System.Activities.Statements
Imports System.Data

Partial Class frmHabilitacao
    Inherits System.Web.UI.Page

    Private Sub frmHabilitacao_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then

            ViewState("Permissao") = (New ComponenteAcesso.Permissao).Funcionalidade(Session("CodigoUsuario"), 1602)
            If ViewState("Permissao") = 1 Then
                pnlCadastro.Enabled = False
                grdhabilitacao.Enabled = False
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


            CarregarComboPeriodo()
            CarregarComboServidor()
            drpPeriodo.SelectedIndex = 1
            drpPeriodo_SelectedIndexChanged(drpPeriodo, New EventArgs())
            CarregarGrid()

        End If


        Validacao.Combo(drpAlocacao, True, "Alocação Carga Horária",)
        Validacao.Combo(drphabilidade, True, "Habilidade")
        Validacao.Livre(txtHoraAlocada, True, "Quantidade de Hora(s)")
        Validacao.Combo(drpPeriodo, True, "Período")

        Validacao.Finalizar(btnSalvar,, True)

    End Sub
    Private Sub CarregarComboPeriodo()
        Dim Cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select RH88_ID_PERIODO as CODIGO, RH88_NM_PERIODO as DESCRICAO ")
        strSQL.Append(" from RH88_PERIODO ")
        strSQL.Append(" order by RH88_NM_PERIODO desc  ")

        With drpPeriodo
            .Items.Clear()

            .DataValueField = "CODIGO"
            .DataTextField = "DESCRICAO"

            .DataSource = Cnn.AbrirDataTable(strSQL.ToString)
            .DataBind()

            .Items.Insert(0, New ListItem("", ""))
        End With


        strSQL.Length = 0
        Cnn = Nothing
    End Sub
    Private Sub CarregarGrid()
        Dim objHabilitacao As New Habilitacao

        grdhabilitacao.DataSource = objHabilitacao.Pesquisar(ViewState("OrderBy"),, Val(drphabilidade.SelectedValue), Val(drpAlocacao.SelectedValue),,,,,, drpServidor.SelectedValue, Session("CodigoPessoaCargaHoraria"), False, Val(drpCargaHoraria.SelectedValue), Val(drpLotacao.SelectedValue), drpPeriodo.SelectedItem.ToString)
        grdhabilitacao.DataBind()

        objHabilitacao = Nothing

        If Val(drpCargaHoraria.SelectedValue) > 0 Then

            With (New Habilitacao).Pesquisar(ViewState("OrderBy"),,, val(drpAlocacao.SelectedValue),,,,,,,,true,Val(drpCargaHoraria.SelectedValue),val(drplotacao.selectedvalue))
                'If .Rows.Count > 0 Then
                    'With DirectCast(grdhabilitacao.DataSource, DataTable)
                        Dim objAlocacaoCargaHoraria As New AlocacaoCargaHoraria(Val(drpAlocacao.SelectedValue))
                        dim objServidorCargaHoraria as New ServidorCargaHoraria(Val(drpCargaHoraria.SelectedValue))
                        Dim objCargaHoraria as New CargaHoraria(val(objServidorCargaHoraria.IdCargaHoraria))
                        Dim Horas As Integer

                        If .Rows.Count > 0 Then
                            For Each dr As DataRow In .Rows
                                Horas += Val(dr("RH74_QT_HORA_ALOCADA"))
                            Next

                            If Horas = objCargaHoraria.QtdMaxHrAula or objAlocacaoCargaHoraria.QtdHotaAlocada - Horas = 0 Then
                                MsgBox("Quantidade máxima de Hora/Aula mapeada!")
                                'btnSalvar.Visible = False
                            End If

                            else 
                                Horas =0

                        end if

                        If drpAlocacao.SelectedValue > 0 Then
                                lblHoraAulaRestante.text  =  objAlocacaoCargaHoraria.QtdHotaAlocada - Horas   & " Hora(s) restante(s)"
                            Else 
                                lblHoraAulaRestante.text  =  objCargaHoraria.QtdMaxHrAula - Horas   & " Hora(s) restante(s)"
                        End If
                        
                        ViewState("Hora") =   objAlocacaoCargaHoraria.QtdHotaAlocada - Horas



                        lblHoraAula.Text = Horas & " Hora (s) Alocadas(s)"

                        Horas = Nothing
                        objAlocacaoCargaHoraria = Nothing

                    'End With
                'End If
            End With

            else
            lblHoraAula.Text = ""
            lblHoraAulaRestante.Text = ""

        End If

        lblRegistros.Text = DirectCast(grdhabilitacao.DataSource, Data.DataTable).Rows.Count & " Registro(s)"
    End Sub
    Private Sub CarregarComboServidor()
        Dim objServidor As New Servidor

        With drpServidor
            .Items.Clear()

            .DataValueField = "RH02_ID_SERVIDOR"
            .DataTextField = "DESCRICAO"

            .DataSource = objServidor.PesquisarServidor(,, (Session("CodigoPessoaCargaHoraria")),,,,,, Servidor.Situacao.ATIVO)
            .DataBind()

            .Items.Insert(0, New ListItem("...", 0))
        End With

        objServidor = Nothing
    End Sub

    Private Sub drpServidor_SelectedIndexChanged(sender As Object, e As EventArgs) Handles drpServidor.SelectedIndexChanged


        If drpServidor.SelectedValue > 0 Then
                CarregarLotacaoServidor()
                CarregarCombohabilidade()
                CarregarComboServidorCargaHoraria()


                drpCargaHoraria.Enabled = True
                drphabilidade.Enabled = True

            Else
                drpAlocacao.Enabled = False
                drphabilidade.Enabled = False
                drpCargaHoraria.Enabled = False

                drpAlocacao.ClearSelection()
                drphabilidade.ClearSelection()
                drpCargaHoraria.ClearSelection()
            End If

        CarregarGrid()
    End Sub

    Private Sub CarregarLotacaoServidor()
       Dim objServidorLotacao as New LotacaoServidor

        with drplotacao

            .Items.Clear()

            .DataValueField = "RH14_ID_LOTACAO_SERVIDOR"
            .DataTextField = "DESCRICAO"

            .DataSource = objServidorLotacao.Pesquisar(,,, drpServidor.SelectedValue,,,,,,,,, True,,, drpPeriodo.SelectedItem.ToString)
            .DataBind()

            .Items.Insert(0, New ListItem("...", 0))


        End With

        objServidorLotacao = nothing
    End Sub

    Private Sub CarregarComboServidorCargaHoraria()

        Dim objServidorCargaHoraria As New ServidorCargaHoraria

        With drpCargaHoraria
            .Items.Clear()

            .DataValueField = "RH78_ID_SERVIDOR_CARGA_HORARIA"
            .DataTextField = "DESCRICAO"

            .DataSource = objServidorCargaHoraria.Pesquisar(,, drpServidor.SelectedValue,,,,,,,,,,, drpPeriodo.SelectedItem.ToString)
            .DataBind()

            .Items.Insert(0, New ListItem("...", 0))
        End With

        objServidorCargaHoraria = Nothing

    End Sub

    Private Sub CarregarCombohabilidade()

        With (New HabilidadeServidor).Pesquisar(,, drpServidor.SelectedValue,,,,,,,,,)
            If .Rows.Count > 0 Then
                Dim objhabilidade As New HabilidadeServidor
                With drphabilidade

                    .Items.Clear()

                    .DataValueField = "RH72_ID_HABILIDADE"
                    .DataTextField = "DESCRICAO"

                    .DataSource = objhabilidade.Pesquisar(,, drpServidor.SelectedValue)
                    .DataBind()

                    .Items.Insert(0, New ListItem("...", 0))


                End With

                objhabilidade = Nothing

            End If
        End With

    End Sub

    Private Sub drpCargaHoraria_SelectedIndexChanged(sender As Object, e As EventArgs) Handles drpCargaHoraria.SelectedIndexChanged
        If drpCargaHoraria.SelectedValue > 0 Then
            CarregarComboAlocacao()
            drpAlocacao.Enabled = True
        Else

            drpAlocacao.Enabled = False
            drpAlocacao.ClearSelection()
            lblHoraAula.text     = ""
            lblHoraAulaRestante.text     = ""
        End If

        CarregarGrid()
    End Sub

    Private Sub CarregarComboAlocacao()
        Dim objAlocacaoCargaHoraria As New AlocacaoCargaHoraria

        With drpAlocacao

            .Items.Clear()

            .DataValueField = "RH80_ID_ALOCACAO_CARGA_HORARIA"
            .DataTextField = "DESCRICAO"

            .DataSource = objAlocacaoCargaHoraria.Pesquisar(,, drpLotacao.SelectedValue, Val(drpCargaHoraria.SelectedValue),,,,,,,,,,,,, drpPeriodo.SelectedItem.ToString)
            .DataBind()

            .Items.Insert(0, New ListItem("...", 0))
        End With

        objAlocacaoCargaHoraria = Nothing

    End Sub

    Private Sub btnNovo_Click(sender As Object, e As EventArgs) Handles btnNovo.Click
        LimparCampos()
        carregargrid()
    End Sub

    Private Sub LimparCampos()
        drpServidor.ClearSelection()
        drpAlocacao.ClearSelection()
        drpCargaHoraria.ClearSelection()
        drphabilidade.ClearSelection()
        drplotacao.ClearSelection


        ViewState("Codigo")  = nothing

        txtHoraAlocada.Text = ""
    End Sub

    Private Sub btnSalvar_Click(sender As Object, e As EventArgs) Handles btnSalvar.Click

        
        try
            Salvar()
            LimparCampos()
            CarregarGrid()

            MsgBox(eTipoMensagem.SALVAR_SUCESSO)
        Catch ex As Exception
            MsgBox(eTipoMensagem.SALVAR_ERRO)
        End Try

    End Sub

    Private Sub Salvar()
        Dim objHabilitacacao As New Habilitacao(ViewState("Codigo"))
        Dim objHabilidade As New HabilidadeServidor(objHabilitacacao.IdHabilidade)

        'Verifico se Existe regristro ativo na DE13_HORARIO_TURMA e caso exista, verifico a quantidade de registros, pois 
        'para cada a quatidade de horas não deve divergir


        If Not ViewState("Codigo") Is Nothing Then
            With (New HorarioTurma).Pesquisar(,,,, objHabilitacacao.IdAlocacaoCargaHoraria,,,,,, objHabilidade.IdDisciplina)
                If .Rows.Count > 0 Then

                    Dim auxHorarioTurma As Integer = 0

                    While auxHorarioTurma < .Rows.Count

                        auxHorarioTurma += 1

                    End While

                    If txtHoraAlocada.Text < auxHorarioTurma Then

                        MsgBox("Impossível alterar a quantidade de horas, pois o professor possui " & auxHorarioTurma & " hora(s) em turma.")

                        .Dispose()

                        auxHorarioTurma = Nothing

                        Exit Sub

                    End If

                End If

                .Dispose()
            End With
        End If

        Using objHabilitacacao

            objHabilitacacao.IdHabilidade = drphabilidade.SelectedValue
            objHabilitacacao.IdAlocacaoCargaHoraria = drpAlocacao.SelectedValue
            objHabilitacacao.IdUsuario = Session("CodigoUsuario")
            objHabilitacacao.QtdHoraAlocada = txtHoraAlocada.Text
            objHabilitacacao.DataHoraCadastro = Date.Now

            'AO CADASTRAR UM NOVO REGISTRO
            If ViewState("Codigo") Is Nothing Then
                'CASO ESTEJA CADASTRANDO UMA NOVA DISTRIBUICAO DE CARGA HORARIA, VERIFICO SE HÁ UMA DISTRIBUIÇÃO PARA AQUELA DISCIPLINA E ALOCAÇÃO, POIS NÃO DEVE SER PERMITIDO
                'CADASTRAR DUAS DISTRIBUIÇÕES NO MESMO TURNO, COM A MESMA DISCIPLINA NO MESMO LOCAL (ALOCAÇÃO ENGLOBA LOCAL E TURNO)

                With (New Habilitacao).Pesquisar(,, drphabilidade.SelectedValue, drpAlocacao.SelectedValue)
                    If .Rows.Count > 0 Then
                        MsgBox("Você já possui um registro com essas informações! Para atualizar a quantidade de horas, clique no registro existente.")
                        .Dispose()
                        Exit Sub
                    End If
                End With

            End If

            'VERIFICO SE A QUANTIDADE DE HORAS DISTRIBUIDA É COMPATIVEL COM A CH ALOCADA

            'SE JÁ EXISTE DISTRIBUICAO C.H PARA AQUELA ALOCAÇÃO, VERIFICO A QUANTIDADE JÁ ALOCADA
            With (New Habilitacao().Pesquisar(,,, objHabilitacacao.IdAlocacaoCargaHoraria))

                If .Rows.Count > 0 Then

                    Dim objAlocacaoCargaHoraria As New AlocacaoCargaHoraria(drpAlocacao.SelectedValue)
                    Dim auxTotalCargaHorariaDistribuida As Integer = 0

                    For Each dr As DataRow In .Rows

                        'VERIFICO SE ESTOU EDITANDO UM REGISTRO DE DISTRIBUIÇÃO C.H PARA CONTABILIZAR A QUANTIDADE DE HORAS DISTRIBUÍDAS
                        If Not Val(dr("RH74_ID_HABILITACAO")) = ViewState("Codigo") Then

                            auxTotalCargaHorariaDistribuida += Val(dr("RH74_QT_HORA_ALOCADA"))

                        End If

                    Next

                    'VERIFICO SE A QUANTIDADE DE HORAS NA DISTRIBUIÇÃO É COMPATÍVEL COM A QUANTIDADE DE HORAS DISPONÍVEIS PARA DISTRIBUIÇÃO
                    If Val(txtHoraAlocada.Text) > (Val(objAlocacaoCargaHoraria.QtdHotaAlocada) - auxTotalCargaHorariaDistribuida) Then
                        MsgBox("Quantidade de horas alocadas é superior a quantidade de horas restantes")
                        Exit Sub
                    End If
                End If
            End With

            objHabilitacacao.Salvar()

        End Using
    End Sub

    Private Sub grdhabilitacao_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles grdhabilitacao.RowDataBound


        Dim lnkCancelar As LinkButton
        Dim iCancelar As New Object
        Dim lblCancelar As New Label

        If e.Row.RowType = DataControlRowType.DataRow Then

            lnkCancelar = DirectCast(e.Row.Cells(8).FindControl("lnkCancelar"), LinkButton)
            iCancelar = DirectCast(e.Row.Cells(8).FindControl("iCancelar"), Object)
            lblCancelar = DirectCast(e.Row.Cells(8).FindControl("lblCancelar"), Label)
            lnkCancelar.CommandArgument = e.Row.RowIndex
            
            'Dim rowView As DataRowView = DirectCast(e.Row.DataItem, DataRowView)

            If grdhabilitacao.DataKeys(e.Row.RowIndex).Item(1) = 0 Then
               
                iCancelar.Attributes.Add("class", "fa fa-trash-o")
                lnkCancelar.Attributes.Add("class", "badge bg-red btn-block")
                lnkCancelar.CommandName = "Desativar"
                ' lnkCancelar.ToolTip = "Disciplina Em andamento ou sem professor"
                lblCancelar.Text = "Desativar"

            Else
                iCancelar.Attributes.Add("class", "fa fa-check")
                lnkCancelar.Attributes.Add("class", "badge bg-blue btn-block")
                lnkCancelar.CommandName = "Ativar"
                'lnkCancelar.ToolTip = "Disciplina Cancelada"
                lblCancelar.Text = "Ativar"
            End If

        End If

        lblCancelar = Nothing
        lnkCancelar = Nothing
        iCancelar = Nothing


    End Sub

    Private Sub grdhabilitacao_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles grdhabilitacao.RowCommand
        If e.CommandName = "Desativar" Then

            DesativarHabilitacao(grdhabilitacao.DataKeys(e.CommandArgument).Item(0))
            LimparCampos()
            CarregarGrid()
        End If


        If e.CommandName = "Ativar" Then
            HabilitarHabilitacao(grdhabilitacao.DataKeys(e.CommandArgument).Item(0))
            CarregarGrid()
        End If

        If e.CommandName = "" Then
            Carregar(grdhabilitacao.DataKeys(e.CommandArgument).Item(0))
            CarregarGrid()
        End If


    End Sub

    Private Sub Carregar(Codigo as integer)
        Dim objHabilitacao As New Habilitacao(Codigo)
        Dim objhabilidade As New HabilidadeServidor(objHabilitacao.IdHabilidade)
        Dim objAlocacaoCargaHoraria as New AlocacaoCargaHoraria(objHabilitacao.IdAlocacaoCargaHoraria)
        Dim objServidorCargaHoraria as New ServidorCargaHoraria(objAlocacaoCargaHoraria.IdServidorCargaHoraria)
       'Dim ObjServidorLotacao As New LotacaoServidor(objAlocacaoCargaHoraria.IdLotacaoServidor)

        ViewState("Codigo") = Codigo

        CarregarComboServidor()
        SelecionarCombo(drpServidor, objServidorCargaHoraria.IdServidor )

        CarregarComboServidorCargaHoraria()
        SelecionarCombo(drpCargaHoraria,objAlocacaoCargaHoraria.IdServidorCargaHoraria)

        CarregarLotacaoServidor()
        SelecionarCombo(drpLotacao, objAlocacaoCargaHoraria.IdLotacaoServidor)

        CarregarComboAlocacao()
        SelecionarCombo(drpAlocacao,objAlocacaoCargaHoraria.IdAlocacaoCargaHoraria)

        
        CarregarCombohabilidade()
        SelecionarCombo(drphabilidade, objHabilitacao.IdHabilidade)

       

        txtHoraAlocada.Text = objHabilitacao.QtdHoraAlocada


        objHabilitacao = Nothing
        objhabilidade = Nothing
        objAlocacaoCargaHoraria=nothing
        objServidorCargaHoraria = nothing
    End Sub

    Private Sub HabilitarHabilitacao(Codigo As integer)
        Dim objHabilitacao as New Habilitacao(Codigo)


        with (New AlocacaoCargaHoraria).Pesquisar(,objHabilitacao.IdAlocacaoCargaHoraria,,,,,,,,,,,,,)
            If .Rows.Count = 0 Then
                MsgBox("Ative Primeiro o Registro de Alocação!")
                exit sub
            End If
        End With


        With (New HabilidadeServidor).Pesquisar(,objHabilitacao.IdHabilidade,,,,,,,,,,)
            If .Rows.Count = 0 Then
                MsgBox("Ative Primeiro o Registro de habilidade!")
                exit sub
            End If
        End With


        With objHabilitacao

            .IdUsuarioAlteracao = Session("CodigoUsuario")
            .DataHoraDesativacao = ""

            .salvar

        End With

        objHabilitacao = Nothing
    End Sub

    Private Sub grdhabilitacao_PageIndexChanging(ByVal source As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles grdhabilitacao.PageIndexChanging
        grdhabilitacao.PageIndex = e.NewPageIndex
        CarregarGrid()
    End Sub

    Private Sub grdhabilitacao_Sorting(ByVal source As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles grdhabilitacao.Sorting
        ViewState("OrderByDirection") = IIf(ViewState("OrderByDirection") = "asc", "desc", "asc")
        ViewState("OrderBy") = e.SortExpression & " " & ViewState("OrderByDirection")
        CarregarGrid()
    End Sub

    Private Sub DesativarHabilitacao(Codigo As Integer)
        Dim objHabilitacao As New Habilitacao(Codigo)
        dim objHabilidadeServidor As New HabilidadeServidor(objHabilitacao.IdHabilidade)

        With (New HorarioTurma).Pesquisar(,,,, objHabilitacao.IdAlocacaoCargaHoraria,,,,,, objHabilidadeServidor.IdDisciplina, drpPeriodo.SelectedItem.ToString)
            If .Rows.Count > 1 Then
                MsgBox("Servidor enturmado, impossível realizar desativação")
                Exit Sub
            End If
        End With

        With objHabilitacao

            .DataHoraDesativacao = Date.Now
            .IdUsuarioAlteracao = Session("CodigoUsuario")

            .Salvar()
        End With

        objHabilitacao = Nothing
    End Sub

    Private Sub drpAlocacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles drpAlocacao.SelectedIndexChanged
        CarregarGrid()
    End Sub

    Private Sub drpLotacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles drpLotacao.SelectedIndexChanged
        If Val(drpLotacao.SelectedValue) > 0 Then
            CarregarGrid()
            drpCargaHoraria.ClearSelection()
            drpAlocacao.ClearSelection()
        End If
    End Sub

    Private Sub drpPeriodo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles drpPeriodo.SelectedIndexChanged
        If Val(drpPeriodo.SelectedValue) Then

            drpServidor.Enabled = True
            drpLotacao.Enabled = True
        Else
            drpServidor.ClearSelection()
            drpLotacao.ClearSelection()

            drpServidor.Enabled = False
            drpLotacao.Enabled = False
            drpCargaHoraria.ClearSelection()
            drpServidor_SelectedIndexChanged(Me, Nothing)
            drpLotacao_SelectedIndexChanged(Me, Nothing)
        End If

        CarregarGrid()
    End Sub
End Class
