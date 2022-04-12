Public Module JavaScript

    Public Sub MsgBox(ByVal Mensagem As String)
        Dim executingPage As Object = HttpContext.Current.Handler
        ScriptManager.RegisterStartupScript(executingPage, executingPage.GetType, "mensagem", String.Format("alert('{0}');", Mensagem), True)
    End Sub

    Public Sub MsgBox(ByVal Mensagem As String, ByVal Pagina As String)
        Dim executingPage As Object = HttpContext.Current.Handler
        ScriptManager.RegisterStartupScript(executingPage, executingPage.GetType, "mensagem", String.Format("alert('{0}');", Mensagem) & String.Format("window.location='{0}';", Pagina), True)
    End Sub

    Public Sub FecharShowProgress()
        Dim executingPage As Object = HttpContext.Current.Handler
        ScriptManager.RegisterStartupScript(executingPage, executingPage.GetType, "showprogress", "$( document ).ajaxStop(function() {  $( '.loading' ).hide(); });", True)
    End Sub

    Public Enum eTipoMensagem As Short
        EXCLUIR_ERRO = 0
        EXCLUIR_SUCESSO = 1
        INCLUIR_ERRO = 2
        INCLUIR_SUCESSO = 3
        ALTERAR_ERRO = 4
        ALTERAR_SUCESSO = 5
        SALVAR_ERRO = 6
        SALVAR_SUCESSO = 7
        RESULTADO_VAZIO = 8
    End Enum

    Public Sub MsgBox(ByVal TipoMensagem As eTipoMensagem)
        Select Case TipoMensagem

            Case eTipoMensagem.EXCLUIR_ERRO
                MsgBox("Erro! Verifique Informações Relacionadas antes de Excluir.")

            Case eTipoMensagem.EXCLUIR_SUCESSO
                MsgBox("Sucesso ao Excluir as Informações!")

            Case eTipoMensagem.INCLUIR_ERRO
                MsgBox("Erro ao Incluir as Informações!")

            Case eTipoMensagem.INCLUIR_SUCESSO
                MsgBox("Sucesso ao Incluir as Informações!")

            Case eTipoMensagem.ALTERAR_ERRO
                MsgBox("Erro ao Alterar as Informações!")

            Case eTipoMensagem.ALTERAR_SUCESSO
                MsgBox("Sucesso ao Alterar as Informações!")

            Case eTipoMensagem.SALVAR_ERRO
                MsgBox("Erro ao Salvar Informações!")

            Case eTipoMensagem.SALVAR_SUCESSO
                MsgBox("Sucesso ao Salvar as Informações!")

            Case eTipoMensagem.RESULTADO_VAZIO
                MsgBox("Sua busca não retornou resultado!")

        End Select
    End Sub

    Public Enum eTipoConfirmacao As Short
        NOVO = 0
        SALVAR = 1
        EXCLUIR = 2
        IMPRIMIR = 3
        VOLTAR = 4
    End Enum

    Public Sub ExibirConfirmacao(ByVal Botao As Button, ByVal TipoConfirmacao As eTipoConfirmacao)
        Select Case TipoConfirmacao
            Case eTipoConfirmacao.NOVO
                ExibirConfirmacao(Botao, "Deseja limpar as informações da tela para inserir um novo registro?")

            Case eTipoConfirmacao.SALVAR
                ExibirConfirmacao(Botao, "Deseja salvar as informações da tela?")

            Case eTipoConfirmacao.EXCLUIR
                ExibirConfirmacao(Botao, "Deseja excluir as informações da tela?")

            Case eTipoConfirmacao.IMPRIMIR
                ExibirConfirmacao(Botao, "Deseja imprimir as informações da tela?")

            Case eTipoConfirmacao.VOLTAR
                ExibirConfirmacao(Botao, "Deseja voltar para a listagem das informações?")

        End Select
    End Sub

    Public Sub ExibirConfirmacao(ByVal Botao As Button, ByVal Mensagem As String)
        Botao.Attributes.Add("OnClick", "if (!confirm('" & Mensagem & "')) { return(false); } ")
    End Sub

    Public Sub ExibirConfirmacao(ByVal Botao As Button, ByVal Mensagem As String, ByVal YesChoosed As String, ByVal NoChoosed As String)
        Botao.Attributes.Add("OnClick", "if (confirm('" & Mensagem & "')) { " & YesChoosed & " } else { " & NoChoosed & "} ")
    End Sub

    Public Sub ExibirConfirmacao(ByVal Botao As LinkButton, ByVal TipoConfirmacao As eTipoConfirmacao)
        Select Case TipoConfirmacao
            Case eTipoConfirmacao.NOVO
                ExibirConfirmacao(Botao, "Deseja limpar as informações da tela para inserir um novo registro?")

            Case eTipoConfirmacao.SALVAR
                ExibirConfirmacao(Botao, "Deseja salvar as informações da tela?")

            Case eTipoConfirmacao.EXCLUIR
                ExibirConfirmacao(Botao, "Deseja excluir as informações da tela?")

            Case eTipoConfirmacao.IMPRIMIR
                ExibirConfirmacao(Botao, "Deseja imprimir as informações da tela?")

            Case eTipoConfirmacao.VOLTAR
                ExibirConfirmacao(Botao, "Deseja voltar para a listagem das informações?")

        End Select
    End Sub

    Public Sub ExibirConfirmacao(ByVal Botao As LinkButton, ByVal Mensagem As String)
        Botao.Attributes.Add("OnClick", "if (!confirm('" & Mensagem & "')) { return(false); } ")
    End Sub

    Public Sub ExibirConfirmacao(ByVal Botao As LinkButton, ByVal Mensagem As String, ByVal YesChoosed As String, ByVal NoChoosed As String)
        Botao.Attributes.Add("OnClick", "if (confirm('" & Mensagem & "')) { " & YesChoosed & " } else { " & NoChoosed & "} ")
    End Sub

    Public Sub DoPostBack()
        Dim executingPage As Object = HttpContext.Current.Handler
        ScriptManager.RegisterStartupScript(executingPage, executingPage.GetType, "PostBack", "__doPostBack('', '');", True)
    End Sub


End Module
