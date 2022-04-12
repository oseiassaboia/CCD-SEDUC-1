Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class Licenca
    Implements IDisposable

    Private RH30_ID_LICENCA As Integer
    Private RH02_ID_SERVIDOR As Integer
    Private RH29_ID_TIPO_LICENCA As Integer
    Private RH30_NR_PROCESSO As Integer 'VERIFICAR COM BRUNO
    Private RH30_NR_ANO_PROCESSO As Integer
    Private RH30_DT_INICIO_AQUISICAO As String
    Private RH30_DT_TERMINO_AQUISICAO As String
    Private RH30_DT_INICIO_GOZO As String
    Private RH30_DT_TERMINO_GOZO As String
    Private RH30_NR_ANO_PORTARIA As Integer
    Private RH30_NR_PORTARIA As Integer
    Private RH30_DT_EMISSAO_PORTARIA As String
    Private RH30_DS_EMBASAMENTO As String
    Private CA04_ID_USUARIO As Integer
    Private RH30_DH_CADASTRO As String

    Private disposedValue As Boolean

#Region "Getters e  Setters"
    Public Property IdLicenca() As Integer
        Get
            Return RH30_ID_LICENCA
        End Get
        Set(value As Integer)
            RH30_ID_LICENCA = value
        End Set
    End Property

    Public Property IdServidor() As Integer
        Get
            Return RH02_ID_SERVIDOR
        End Get
        Set(value As Integer)
            RH02_ID_SERVIDOR = value
        End Set
    End Property

    Public Property TipoLicenca() As Integer
        Get
            Return RH29_ID_TIPO_LICENCA
        End Get
        Set(value As Integer)
            RH29_ID_TIPO_LICENCA = value
        End Set
    End Property

    Public Property NumeroProcesso() As Integer
        Get
            Return RH30_NR_PROCESSO
        End Get
        Set(value As Integer)
            RH30_NR_PROCESSO = value
        End Set
    End Property

    Public Property AnoProcesso() As Integer
        Get
            Return RH30_NR_ANO_PROCESSO
        End Get
        Set(value As Integer)
            RH30_NR_ANO_PROCESSO = value
        End Set
    End Property

    Public Property DataInicioAquisicao() As String
        Get
            Return RH30_DT_INICIO_AQUISICAO
        End Get
        Set(value As String)
            RH30_DT_INICIO_AQUISICAO = value
        End Set
    End Property

    Public Property DataTerminoAquisicao() As String
        Get
            Return RH30_DT_TERMINO_AQUISICAO
        End Get
        Set(value As String)
            RH30_DT_TERMINO_AQUISICAO = value
        End Set
    End Property

    Public Property DataTerminoGozo() As String
        Get
            Return RH30_DT_TERMINO_GOZO
        End Get
        Set(value As String)
            RH30_DT_TERMINO_GOZO = value
        End Set
    End Property

    Public Property DataInicioGozo() As String
        Get
            Return RH30_DT_INICIO_GOZO
        End Get
        Set(value As String)
            RH30_DT_INICIO_GOZO = value
        End Set
    End Property

    Public Property AnoPortaria() As Integer
        Get
            Return RH30_NR_ANO_PORTARIA
        End Get
        Set(value As Integer)
            RH30_NR_ANO_PORTARIA = value
        End Set
    End Property

    Public Property NumeroPortaria() As Integer
        Get
            Return RH30_NR_PORTARIA
        End Get
        Set(value As Integer)
            RH30_NR_PORTARIA = value
        End Set
    End Property

    Public Property DataEmissaoPortaria() As String
        Get
            Return RH30_DT_EMISSAO_PORTARIA
        End Get
        Set(value As String)
            RH30_DT_EMISSAO_PORTARIA = value
        End Set
    End Property

    Public Property Embasamento() As String
        Get
            Return RH30_DS_EMBASAMENTO
        End Get
        Set(value As String)
            RH30_DS_EMBASAMENTO = value
        End Set
    End Property

    Public Property IdUsuario() As Integer
        Get
            Return CA04_ID_USUARIO
        End Get
        Set(value As Integer)
            CA04_ID_USUARIO = value
        End Set
    End Property

    Public Property DataCadastro() As String
        Get
            Return RH30_DH_CADASTRO
        End Get
        Set(value As String)
            RH30_DH_CADASTRO = value
        End Set
    End Property
#End Region
    Public Sub New(Optional ByVal IdLicenca As Integer = 0)
        If IdLicenca > 0 Then
            Obter(IdLicenca)
        End If
    End Sub

    Public Sub Salvar(Optional ByRef transacao As Transacao = Nothing)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH30_LICENCA")
        strSQL.Append(" where RH30_ID_LICENCA  = " & IdLicenca)

        dt = cnn.EditarDataTable(strSQL.ToString, transacao)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH30_ID_LICENCA") = ProBanco(RH30_ID_LICENCA, eTipoValor.CHAVE)
        dr("RH02_ID_SERVIDOR") = ProBanco(RH02_ID_SERVIDOR, eTipoValor.CHAVE)
        dr("RH29_ID_TIPO_LICENCA") = ProBanco(RH29_ID_TIPO_LICENCA, eTipoValor.CHAVE)
        dr("RH30_NR_PROCESSO") = ProBanco(RH30_NR_PROCESSO, eTipoValor.NUMERO_INTEIRO)
        dr("RH30_NR_ANO_PROCESSO") = ProBanco(RH30_NR_ANO_PROCESSO, eTipoValor.NUMERO_INTEIRO)
        dr("RH30_NR_ANO_PORTARIA") = ProBanco(RH30_NR_ANO_PORTARIA, eTipoValor.NUMERO_INTEIRO)
        dr("RH30_DT_INICIO_AQUISICAO") = ProBanco(RH30_DT_INICIO_AQUISICAO, eTipoValor.DATA)
        dr("RH30_DT_TERMINO_AQUISICAO") = ProBanco(RH30_DT_TERMINO_AQUISICAO, eTipoValor.DATA)
        dr("RH30_DT_INICIO_GOZO") = ProBanco(RH30_DT_INICIO_GOZO, eTipoValor.DATA)
        dr("RH30_DT_TERMINO_GOZO") = ProBanco(RH30_DT_TERMINO_GOZO, eTipoValor.DATA)
        dr("RH30_NR_PORTARIA") = ProBanco(RH30_NR_PORTARIA, eTipoValor.NUMERO_INTEIRO)
        dr("RH30_DT_EMISSAO_PORTARIA") = ProBanco(RH30_DT_EMISSAO_PORTARIA, eTipoValor.DATA)
        dr("RH30_DS_EMBASAMENTO") = ProBanco(RH30_DS_EMBASAMENTO, eTipoValor.TEXTO)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("RH30_DH_CADASTRO") = ProBanco(RH30_DH_CADASTRO, eTipoValor.DATA_COMPLETA)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal IdLicenca As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH30_LICENCA")
        strSQL.Append(" where RH30_ID_LICENCA = " & IdLicenca)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH30_ID_LICENCA = DoBanco(dr("RH30_ID_LICENCA"), eTipoValor.CHAVE)
            RH02_ID_SERVIDOR = DoBanco(dr("RH02_ID_SERVIDOR"), eTipoValor.CHAVE)
            RH29_ID_TIPO_LICENCA = DoBanco(dr("RH29_ID_TIPO_LICENCA"), eTipoValor.CHAVE)
            RH30_NR_PROCESSO = DoBanco(dr("RH30_NR_PROCESSO"), eTipoValor.NUMERO_INTEIRO)
            RH30_NR_ANO_PROCESSO = DoBanco(dr("RH30_NR_ANO_PROCESSO"), eTipoValor.NUMERO_INTEIRO)
            RH30_NR_ANO_PORTARIA = DoBanco(dr("RH30_NR_ANO_PORTARIA"), eTipoValor.NUMERO_INTEIRO)
            RH30_NR_PORTARIA = DoBanco(dr("RH30_NR_PORTARIA"), eTipoValor.NUMERO_INTEIRO)
            RH30_DT_INICIO_AQUISICAO = DoBanco(dr("RH30_DT_INICIO_AQUISICAO"), eTipoValor.DATA)
            RH30_DT_TERMINO_AQUISICAO = DoBanco(dr("RH30_DT_TERMINO_AQUISICAO"), eTipoValor.DATA)
            RH30_DT_INICIO_GOZO = DoBanco(dr("RH30_DT_INICIO_GOZO"), eTipoValor.DATA)
            RH30_DT_TERMINO_GOZO = DoBanco(dr("RH30_DT_TERMINO_GOZO"), eTipoValor.DATA)
            RH30_DS_EMBASAMENTO = DoBanco(dr("RH30_DS_EMBASAMENTO"), eTipoValor.TEXTO)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            RH30_DT_EMISSAO_PORTARIA = DoBanco(dr("RH30_DT_EMISSAO_PORTARIA"), eTipoValor.DATA)
            RH30_DH_CADASTRO = DoBanco(dr("RH30_DH_CADASTRO"), eTipoValor.DATA_COMPLETA)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function PesquisarPorSolicitacaoCorrecDoc(Optional ByVal Sort As String = "",
                                        Optional Situacao As String = "1",
                                        Optional CodigoLotacao As Integer = 0,
                                        Optional Ativo As Integer = 1,
                                        Optional NomeServidor As String = "") As DataTable

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT  RH30.RH30_ID_LICENCA, " & vbCrLf)
        strSQL.Append("         RH01.RH01_NM_PESSOA,  " & vbCrLf)
        strSQL.Append("         RH29.RH29_NM_TIPO_LICENCA,  " & vbCrLf)
        strSQL.Append("         RH30.RH30_NR_PROCESSO,  " & vbCrLf)
        strSQL.Append("         RH30.RH30_NR_ANO_PROCESSO,  " & vbCrLf)
        strSQL.Append("         PC01.RH02_ID_SERVIDOR_BENEF,  " & vbCrLf)
        strSQL.Append("         DI09.DI09_RESUMO  " & vbCrLf)
        strSQL.Append("    FROM RH30_LICENCA                                                        RH30  " & vbCrLf)
        strSQL.Append("    JOIN RH02_SERVIDOR                                                       RH02 ON RH30.RH02_ID_SERVIDOR     = RH02.RH02_ID_SERVIDOR  " & vbCrLf)
        strSQL.Append("    JOIN RH01_PESSOA                                                         RH01 ON RH01.RH01_ID_PESSOA       = RH02.RH01_ID_PESSOA  " & vbCrLf)
        strSQL.Append("    JOIN RH29_TIPO_LICENCA                                                   RH29 ON RH29.RH29_ID_TIPO_LICENCA = RH30.RH29_ID_TIPO_LICENCA  " & vbCrLf)
        strSQL.Append("    JOIN RH92_LICENCA_PROCESSO                                               RH92 ON RH92.RH30_ID_LICENCA      = RH30.RH30_ID_LICENCA  " & vbCrLf)
        strSQL.Append("    JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[DBO].PC01_PROCESSO           PC01 ON PC01.PC01_ID_PROCESSO     = RH92.PC01_ID_PROCESSO  " & vbCrLf)
        strSQL.Append("    JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[DBO].PC02_MOVIMENTO_PROCESSO PC02 ON PC02.PC01_ID_PROCESSO     = PC01.PC01_ID_PROCESSO  " & vbCrLf)
        strSQL.Append("    JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[DBO].DI09_STATUS             DI09 ON DI09.DI09_COD_STATUS      = PC02.DI09_ID_STATUS  " & vbCrLf)
        strSQL.Append("   WHERE RH30.RH30_ID_LICENCA IS NOT NULL " & vbCrLf)
        strSQL.Append("     AND DI09.RH36_ID_LOTACAO = " & CodigoLotacao & vbCrLf)
        strSQL.Append("     AND PC02.PC02_IN_ATIVO = " & Ativo & vbCrLf)
        strSQL.Append("   UNION " & vbCrLf)
        strSQL.Append("  SELECT RH30.RH30_ID_LICENCA, " & vbCrLf)
        strSQL.Append("         RH01.RH01_NM_PESSOA,  " & vbCrLf)
        strSQL.Append("         RH29.RH29_NM_TIPO_LICENCA,  " & vbCrLf)
        strSQL.Append("         RH30.RH30_NR_PROCESSO, " & vbCrLf)
        strSQL.Append("         PC01.RH02_ID_SERVIDOR_BENEF,  " & vbCrLf)
        strSQL.Append("         'HÁ ALGUMA PENDÊNCIA NA DOCUMENTAÇÃ' AS DI09_RESUMO " & vbCrLf)
        strSQL.Append("    FROM RH100_SOLIC_CORREC_DOC RH100 " & vbCrLf)
        strSQL.Append("    JOIN RH96_LICENCA_DOC                                                    RH96 ON RH96.RH96_ID_LICENCA_DOC     = RH100.RH96_ID_LICENCA_DOC  " & vbCrLf)
        strSQL.Append("    JOIN RH30_LICENCA                                                        RH30 ON RH30.RH30_ID_LICENCA         = RH96.RH30_ID_LICENCA  " & vbCrLf)
        strSQL.Append("    JOIN RH29_TIPO_LICENCA                                                   RH29 ON RH29.RH29_ID_TIPO_LICENCA    = RH30.RH29_ID_TIPO_LICENCA " & vbCrLf)
        strSQL.Append("    JOIN RH02_SERVIDOR                                                       RH02 ON RH30.RH02_ID_SERVIDOR        = RH02.RH02_ID_SERVIDOR  " & vbCrLf)
        strSQL.Append("    JOIN RH01_PESSOA                                                         RH01 ON RH01.RH01_ID_PESSOA          = RH02.RH01_ID_PESSOA  " & vbCrLf)
        strSQL.Append("    JOIN RH92_LICENCA_PROCESSO                                               RH92 ON RH92.RH30_ID_LICENCA         = RH30.RH30_ID_LICENCA  " & vbCrLf)
        strSQL.Append("    JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.DBO.DI02_TIPO_DOCUMENTO         DI02 ON DI02.DI02_COD_TIPO_DOCUMENTO = RH96.DI02_ID_TIPO_DOCUMENTO  " & vbCrLf)
        strSQL.Append("    JOIN RH95_STATUS_CATEG_TIPO_DOC                                          RH95 ON RH95.DI02_ID_TIPO_DOCUMENTO  = DI02.DI02_COD_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("    JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.[DBO].[DI09_STATUS]             DI09 ON DI09.DI09_COD_STATUS         = RH95.DI09_ID_STATUS    " & vbCrLf)
        strSQL.Append("    JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[DBO].PC01_PROCESSO           PC01 ON PC01.PC01_ID_PROCESSO        = RH92.PC01_ID_PROCESSO  " & vbCrLf)
        strSQL.Append("    JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[DBO].PC02_MOVIMENTO_PROCESSO PC02 ON PC02.PC01_ID_PROCESSO        = PC01.PC01_ID_PROCESSO  " & vbCrLf)
        strSQL.Append("   WHERE RH30.RH30_ID_LICENCA IS NOT NULL " & vbCrLf)
        strSQL.Append("     AND RH100.RH36_ID_LOTACAO = " & CodigoLotacao & vbCrLf)
        strSQL.Append("     AND RH100.RH100_ST_SOLIC_CORREC_DOC in (" & Situacao & ")" & vbCrLf)
        strSQL.Append("     AND PC02.PC02_IN_ATIVO = " & Ativo & vbCrLf)
        strSQL.Append("   UNION " & vbCrLf)
        strSQL.Append("  SELECT RH30.RH30_ID_LICENCA, " & vbCrLf)
        strSQL.Append("         RH01.RH01_NM_PESSOA,  " & vbCrLf)
        strSQL.Append("         RH29.RH29_NM_TIPO_LICENCA,  " & vbCrLf)
        strSQL.Append("         RH30.RH30_NR_PROCESSO, " & vbCrLf)
        strSQL.Append("         PC01.RH02_ID_SERVIDOR_BENEF,  " & vbCrLf)
        strSQL.Append("         'HÁ ALGUMA PENDÊNCIA NA DOCUMENTAÇÃO' AS DI09_RESUMO " & vbCrLf)
        strSQL.Append("    FROM RH102_SOLIC_CAD_DOC                                         RH102     " & vbCrLf)
        strSQL.Append("    JOIN RH30_LICENCA                                                RH30 ON RH30.RH30_ID_LICENCA         = RH102.RH30_ID_LICENCA  " & vbCrLf)
        strSQL.Append("    JOIN RH02_SERVIDOR                                               RH02 ON RH30.RH02_ID_SERVIDOR        = RH02.RH02_ID_SERVIDOR  " & vbCrLf)
        strSQL.Append("    JOIN RH01_PESSOA                                                 RH01 ON RH01.RH01_ID_PESSOA          = RH02.RH01_ID_PESSOA " & vbCrLf)
        strSQL.Append("    JOIN RH29_TIPO_LICENCA                                           RH29 ON RH29.RH29_ID_TIPO_LICENCA    = RH30.RH29_ID_TIPO_LICENCA " & vbCrLf)
        strSQL.Append("    JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.DBO.DI02_TIPO_DOCUMENTO DI02 ON DI02.DI02_COD_TIPO_DOCUMENTO = RH102.DI02_ID_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("    JOIN RH95_STATUS_CATEG_TIPO_DOC                                  RH95 ON RH95.DI02_ID_TIPO_DOCUMENTO  = DI02.DI02_COD_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("    JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.[DBO].[DI09_STATUS]     DI09 ON DI09.DI09_COD_STATUS         = RH95.DI09_ID_STATUS " & vbCrLf)
        strSQL.Append("    JOIN RH36_LOTACAO                                                RH36 ON RH36.RH36_ID_LOTACAO         = DI09.RH36_ID_LOTACAO  " & vbCrLf)
        strSQL.Append("    JOIN RH92_LICENCA_PROCESSO                                       RH92 ON RH92.RH30_ID_LICENCA         = RH30.RH30_ID_LICENCA " & vbCrLf)
        strSQL.Append("    JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[DBO].PC01_PROCESSO   PC01 ON PC01.PC01_ID_PROCESSO        = RH92.PC01_ID_PROCESSO " & vbCrLf)
        strSQL.Append("    JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[DBO].PC02_MOVIMENTO_PROCESSO PC02 ON PC02.PC01_ID_PROCESSO        = PC01.PC01_ID_PROCESSO " & vbCrLf)
        strSQL.Append("   WHERE RH102.RH102_ID_SOLIC_CAD_DOC IS NOT NULL  " & vbCrLf)
        strSQL.Append("     AND RH102.RH102_ST_SOLIC_CAD_DOC in (" & Situacao & ")" & vbCrLf)
        strSQL.Append("     AND DI09.RH36_ID_LOTACAO = " & CodigoLotacao & vbCrLf)
        strSQL.Append(" 	AND PC02.PC02_IN_ATIVO = " & Ativo & vbCrLf)

        If NomeServidor <> "" Then
            strSQL.Append(" AND UPPER(RH01.RH01_NM_PESSOA) LIKE '%" & NomeServidor.ToUpper & "%'" & vbCrLf)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH30.RH30_ID_LICENCA", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function Pesquisar(Optional ByVal Sort As String = "",
                              Optional IdLicenca As Integer = 0,
                              Optional IdServidor As Integer = 0,
                              Optional IdTipoLicenca As Integer = 0,
                              Optional NrProcesso As Integer = 0,
                              Optional IdPessoa As Integer = 0,
                              Optional NrAnoProcesso As Integer = 0,
                              Optional CodigoLotacao As String = "",
                              Optional Ativo As Integer = 1,
                              Optional NomeServidor As String = "") As DataTable

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT RH30.RH30_ID_LICENCA," & vbCrLf)
        strSQL.Append("        RH01.RH01_NM_PESSOA, " & vbCrLf)
        strSQL.Append("        RH29.RH29_NM_TIPO_LICENCA, " & vbCrLf)
        strSQL.Append("        RH30.RH30_NR_PROCESSO, " & vbCrLf)
        strSQL.Append("        RH30.RH30_NR_ANO_PROCESSO, " & vbCrLf)
        strSQL.Append("        PC01.RH02_ID_SERVIDOR_BENEF, " & vbCrLf)
        strSQL.Append("        DI09.DI09_RESUMO " & vbCrLf)
        strSQL.Append("   FROM RH30_LICENCA                                                        RH30 " & vbCrLf)
        strSQL.Append("   JOIN RH02_SERVIDOR                                                       RH02 ON RH30.RH02_ID_SERVIDOR     = RH02.RH02_ID_SERVIDOR " & vbCrLf)
        strSQL.Append("   JOIN RH01_PESSOA                                                         RH01 ON RH01.RH01_ID_PESSOA       = RH02.RH01_ID_PESSOA " & vbCrLf)
        strSQL.Append("   JOIN RH29_TIPO_LICENCA                                                   RH29 ON RH29.RH29_ID_TIPO_LICENCA = RH30.RH29_ID_TIPO_LICENCA " & vbCrLf)
        strSQL.Append("   JOIN RH92_LICENCA_PROCESSO                                               RH92 ON RH92.RH30_ID_LICENCA      = RH30.RH30_ID_LICENCA " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[DBO].PC01_PROCESSO           PC01 ON PC01.PC01_ID_PROCESSO     = RH92.PC01_ID_PROCESSO " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[DBO].PC02_MOVIMENTO_PROCESSO PC02 ON PC02.PC01_ID_PROCESSO     = PC01.PC01_ID_PROCESSO " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[DBO].DI09_STATUS             DI09 ON DI09.DI09_COD_STATUS      = PC02.DI09_ID_STATUS " & vbCrLf)
        strSQL.Append("  WHERE RH30.RH30_ID_LICENCA IS NOT NULL" & vbCrLf)

        If IdLicenca > 0 Then
            strSQL.Append(" AND RH30.RH30_ID_LICENCA = " & IdLicenca & vbCrLf)
        End If

        If IdServidor > 0 Then
            strSQL.Append(" AND RH30.RH02_ID_SERVIDOR = " & IdServidor & vbCrLf)
        End If

        If IdTipoLicenca > 0 Then
            strSQL.Append(" AND RH30.RH29_ID_TIPO_LICENCA = " & IdTipoLicenca & vbCrLf)
        End If

        If NrProcesso > 0 Then
            strSQL.Append(" AND RH30.RH30_NR_PROCESSO = " & NrProcesso & vbCrLf)
        End If

        If IdPessoa > 0 Then
            strSQL.Append(" AND RH02.RH01_ID_PESSOA = " & IdPessoa & vbCrLf)
        End If

        If NrAnoProcesso > 0 Then
            strSQL.Append(" AND RH30.RH30_NR_ANO_PROCESSO = " & NrAnoProcesso & vbCrLf)
        End If

        If CodigoLotacao <> "" Then
            strSQL.Append(" AND DI09.RH36_ID_LOTACAO in (" & CodigoLotacao & ")" & vbCrLf)
        End If

        If Ativo > 0 Then
            strSQL.Append(" AND PC02.PC02_IN_ATIVO = " & Ativo & vbCrLf)
        End If

        If NomeServidor <> "" Then
            strSQL.Append(" AND UPPER(RH01.RH01_NM_PESSOA) LIKE '%" & NomeServidor.ToUpper & "%'" & vbCrLf)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH30.RH30_ID_LICENCA", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select RH30_ID_LICENCA as CODIGO, RH30_ID_LICENCA as DESCRICAO")
        strSQL.Append(" from RH30_LICENCA")
        strSQL.Append(" order by 2 ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return dt

    End Function

    Public Function ObterUltimo(Optional ByRef transacao As Transacao = Nothing) As Integer

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" SELECT MAX(RH30_ID_LICENCA) FROM RH30_LICENCA")

        With cnn.AbrirDataTable(strSQL.ToString, transacao)
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

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' Tarefa pendente: descartar o estado gerenciado (objetos gerenciados)
            End If

            ' Tarefa pendente: liberar recursos não gerenciados (objetos não gerenciados) e substituir o finalizador
            ' Tarefa pendente: definir campos grandes como nulos
            disposedValue = True
        End If
    End Sub

    ' ' Tarefa pendente: substituir o finalizador somente se 'Dispose(disposing As Boolean)' tiver o código para liberar recursos não gerenciados
    ' Protected Overrides Sub Finalize()
    '     ' Não altere este código. Coloque o código de limpeza no método 'Dispose(disposing As Boolean)'
    '     Dispose(disposing:=False)
    '     MyBase.Finalize()
    ' End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Não altere este código. Coloque o código de limpeza no método 'Dispose(disposing As Boolean)'
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
