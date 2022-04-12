Imports System.Data
Imports Microsoft.VisualBasic

Public Class TipoLicenca

    Implements IDisposable

    Private RH29_ID_TIPO_LICENCA As Integer
    Private RH75_ID_REGRA_VALIDACAO_LICENCA As Integer
    Private RH29_NM_TIPO_LICENCA As String
    Private RH29_CABECALHO_RELATORIO_LICENCA As String
    Private RH29_CABECALHO_PAGINA_LICENCA As String
    Private RH29_BODY_A_LICENCA As String
    Private RH29_BODY_B_LICENCA As String
    Private RH29_RODAPE_PAGINA_LICENCA As String
    Private RH29_RODAPE_RELATORIO_LICENCA As String
    Private RH29_LEGENDA As String
    Private RH29_LEI As String
    Private RH29_IN_RETIFICACAO As Boolean
    Private RH29_IN_AFASTAMENTO As Boolean
    Private RH29_IN_TEMPO_FUNCAO As Boolean

    Public Property Codigo() As Integer
        Get
            Return RH29_ID_TIPO_LICENCA
        End Get
        Set(value As Integer)
            RH29_ID_TIPO_LICENCA = value
        End Set
    End Property

    Public Property IdRegraValidacaoLicenca() As Integer
        Get
            Return RH75_ID_REGRA_VALIDACAO_LICENCA
        End Get
        Set(value As Integer)
            RH75_ID_REGRA_VALIDACAO_LICENCA = value
        End Set
    End Property

    Public Property TipoLicenca() As String
        Get
            Return RH29_NM_TIPO_LICENCA
        End Get
        Set(value As String)
            RH29_NM_TIPO_LICENCA = value
        End Set
    End Property

    Public Property CabecalhoRelatorio() As String
        Get
            Return RH29_CABECALHO_RELATORIO_LICENCA
        End Get
        Set(value As String)
            RH29_CABECALHO_RELATORIO_LICENCA = value
        End Set
    End Property

    Public Property CabecalhoPagina() As String
        Get
            Return RH29_CABECALHO_PAGINA_LICENCA
        End Get
        Set(value As String)
            RH29_CABECALHO_PAGINA_LICENCA = value
        End Set
    End Property

    Public Property DetalhesA() As String
        Get
            Return RH29_BODY_A_LICENCA
        End Get
        Set(value As String)
            RH29_BODY_A_LICENCA = value
        End Set
    End Property

    Public Property DetalhesB() As String
        Get
            Return RH29_BODY_B_LICENCA
        End Get
        Set(value As String)
            RH29_BODY_B_LICENCA = value
        End Set
    End Property

    Public Property RodapePagina() As String
        Get
            Return RH29_RODAPE_PAGINA_LICENCA
        End Get
        Set(value As String)
            RH29_RODAPE_PAGINA_LICENCA = value
        End Set
    End Property

    Public Property RodapeRelatorio() As String
        Get
            Return RH29_RODAPE_RELATORIO_LICENCA
        End Get
        Set(value As String)
            RH29_RODAPE_RELATORIO_LICENCA = value
        End Set
    End Property

    Public Property Legenda() As String
        Get
            Return RH29_LEGENDA
        End Get
        Set(value As String)
            RH29_LEGENDA = value
        End Set
    End Property

    Public Property Lei() As String
        Get
            Return RH29_LEI
        End Get
        Set(value As String)
            RH29_LEI = value
        End Set
    End Property

    Public Property Retificacao() As Boolean
        Get
            Return RH29_IN_RETIFICACAO
        End Get
        Set(value As Boolean)
            RH29_IN_RETIFICACAO = value
        End Set
    End Property

    Public Property Afastamento() As Boolean
        Get
            Return RH29_IN_AFASTAMENTO
        End Get
        Set(value As Boolean)
            RH29_IN_AFASTAMENTO = value
        End Set
    End Property

    Public Property TempoFuncao() As Boolean
        Get
            Return RH29_IN_TEMPO_FUNCAO
        End Get
        Set(value As Boolean)
            RH29_IN_TEMPO_FUNCAO = value
        End Set
    End Property

    Public Sub New(Optional ByVal IdTipoLicenca As Integer = 0)
        If IdTipoLicenca > 0 Then
            Obter(IdTipoLicenca)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH29_TIPO_LICENCA")
        strSQL.Append(" where RH29_ID_TIPO_LICENCA = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH29_ID_TIPO_LICENCA") = ProBanco(RH29_ID_TIPO_LICENCA, eTipoValor.CHAVE)
        'dr("RH75_ID_REGRA_VALIDACAO_LICENCA") = ProBanco(RH75_ID_REGRA_VALIDACAO_LICENCA, eTipoValor.CHAVE)
        dr("RH29_NM_TIPO_LICENCA") = ProBanco(RH29_NM_TIPO_LICENCA, eTipoValor.TEXTO_LIVRE)
        dr("RH29_CABECALHO_RELATORIO_LICENCA") = ProBanco(RH29_CABECALHO_RELATORIO_LICENCA, eTipoValor.TEXTO_LIVRE)
        dr("RH29_CABECALHO_PAGINA_LICENCA") = ProBanco(RH29_CABECALHO_PAGINA_LICENCA, eTipoValor.TEXTO_LIVRE)
        dr("RH29_BODY_A_LICENCA") = ProBanco(RH29_BODY_A_LICENCA, eTipoValor.TEXTO_LIVRE)
        dr("RH29_BODY_B_LICENCA") = ProBanco(RH29_BODY_B_LICENCA, eTipoValor.TEXTO_LIVRE)
        dr("RH29_RODAPE_PAGINA_LICENCA") = ProBanco(RH29_RODAPE_PAGINA_LICENCA, eTipoValor.TEXTO_LIVRE)
        dr("RH29_RODAPE_RELATORIO_LICENCA") = ProBanco(RH29_RODAPE_RELATORIO_LICENCA, eTipoValor.TEXTO_LIVRE)
        dr("RH29_LEGENDA") = ProBanco(RH29_LEGENDA, eTipoValor.TEXTO)
        dr("RH29_LEI") = ProBanco(RH29_LEI, eTipoValor.TEXTO)
        dr("RH29_IN_RETIFICACAO") = ProBanco(RH29_IN_RETIFICACAO, eTipoValor.BOOLEANO)
        dr("RH29_IN_AFASTAMENTO") = ProBanco(RH29_IN_AFASTAMENTO, eTipoValor.BOOLEANO)
        dr("RH29_IN_TEMPO_FUNCAO") = ProBanco(RH29_IN_TEMPO_FUNCAO, eTipoValor.BOOLEANO)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal IdTipoLicenca As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH29_TIPO_LICENCA")
        strSQL.Append(" where RH29_ID_TIPO_LICENCA = " & IdTipoLicenca)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH29_ID_TIPO_LICENCA = DoBanco(dr("RH29_ID_TIPO_LICENCA"), eTipoValor.CHAVE)
            'RH75_ID_REGRA_VALIDACAO_LICENCA = DoBanco(dr("RH75_ID_REGRA_VALIDACAO_LICENCA"), eTipoValor.CHAVE)
            RH29_NM_TIPO_LICENCA = DoBanco(dr("RH29_NM_TIPO_LICENCA"), eTipoValor.TEXTO_LIVRE)
            RH29_IN_RETIFICACAO = DoBanco(dr("RH29_IN_RETIFICACAO"), eTipoValor.BOOLEANO)
            RH29_IN_AFASTAMENTO = DoBanco(dr("RH29_IN_AFASTAMENTO"), eTipoValor.BOOLEANO)
            RH29_IN_TEMPO_FUNCAO = DoBanco(dr("RH29_IN_TEMPO_FUNCAO"), eTipoValor.BOOLEANO)

            RH29_CABECALHO_RELATORIO_LICENCA = DoBanco(dr("RH29_CABECALHO_RELATORIO_LICENCA"), eTipoValor.TEXTO_LIVRE)
            RH29_CABECALHO_PAGINA_LICENCA = DoBanco(dr("RH29_CABECALHO_PAGINA_LICENCA"), eTipoValor.TEXTO_LIVRE)
            RH29_BODY_A_LICENCA = DoBanco(dr("RH29_BODY_A_LICENCA"), eTipoValor.TEXTO_LIVRE)
            RH29_BODY_B_LICENCA = DoBanco(dr("RH29_BODY_B_LICENCA"), eTipoValor.TEXTO_LIVRE)
            RH29_RODAPE_PAGINA_LICENCA = DoBanco(dr("RH29_RODAPE_PAGINA_LICENCA"), eTipoValor.TEXTO_LIVRE)
            RH29_RODAPE_RELATORIO_LICENCA = DoBanco(dr("RH29_RODAPE_RELATORIO_LICENCA"), eTipoValor.TEXTO_LIVRE)
            RH29_LEGENDA = DoBanco(dr("RH29_LEGENDA"), eTipoValor.TEXTO_LIVRE)
            RH29_LEI = DoBanco(dr("RH29_LEI"), eTipoValor.TEXTO_LIVRE)

        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "",
                              Optional IdTipoLicenca As Integer = 0,
                              Optional IdRegraValidacao As Integer = 0,
                              Optional TipoLicenca As String = "") As DataTable

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select *")
        strSQL.Append(" from RH29_TIPO_LICENCA")
        strSQL.Append(" where RH29_ID_TIPO_LICENCA is not null")

        If IdTipoLicenca > 0 Then
            strSQL.Append(" and RH29_ID_TIPO_LICENCA = " & IdTipoLicenca)
        End If

        If IdRegraValidacao > 0 Then
            strSQL.Append(" and RH75_ID_REGRA_VALIDACAO_LICENCA = " & IdRegraValidacao)
        End If

        If TipoLicenca <> "" Then
            strSQL.Append(" and upper(RH29_NM_TIPO_LICENCA) like '%" & TipoLicenca.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH29_ID_TIPO_LICENCA", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select RH29_ID_TIPO_LICENCA as CODIGO, RH29_NM_TIPO_LICENCA as DESCRICAO")
        strSQL.Append(" from RH29_TIPO_LICENCA")
        strSQL.Append(" order by 2 ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return dt

    End Function

    Public Function ObterUltimo() As Integer

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(RH29_ID_TIPO_LICENCA) from RH29_TIPO_LICENCA")

        With cnn.AbrirDataTable(strSQL.ToString)
            If Not IsDBNull(.Rows(0)(0)) Then
                CodigoUltimo = .Rows(0)(0)
            Else
                CodigoUltimo = 0
            End If
        End With

        cnn.FecharBanco()
        cnn = Nothing

        Return CodigoUltimo

    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
