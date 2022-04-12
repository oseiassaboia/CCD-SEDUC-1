Imports Microsoft.VisualBasic
Imports System.Data

Public Class SolicCorrecaoDoc

    Private RH100_ID_SOLIC_CORREC_DOC As Integer
    Private RH96_ID_LICENCA_DOC As Integer
    Private RH36_ID_LOTACAO As Integer
    Private CA04_ID_USUARIO As Integer
    Private CA04_ID_USUARIO_ALT As Integer
    Private RH100_ST_SOLIC_CORREC_DOC As Integer
    Private RH100_DH_ST_SOLIC_CORREC_DOC As String
    Private RH100_DH_CADASTRO As String
    Private RH100_DH_ALTERACAO As String
    Private RH100_DS_OBSERVACAO As String

    Public Property IdSolicCorrecaoDoc() As Integer
        Get
            Return RH100_ID_SOLIC_CORREC_DOC
        End Get
        Set(ByVal value As Integer)
            RH100_ID_SOLIC_CORREC_DOC = value
        End Set
    End Property

    Public Property LicencaDoc() As Integer
        Get
            Return RH96_ID_LICENCA_DOC
        End Get
        Set(ByVal value As Integer)
            RH96_ID_LICENCA_DOC = value
        End Set
    End Property

    Public Property LotocaoSolicCorrec() As Integer
        Get
            Return RH36_ID_LOTACAO
        End Get
        Set(ByVal value As Integer)
            RH36_ID_LOTACAO = value
        End Set
    End Property

    Public Property Usuario() As Integer
        Get
            Return CA04_ID_USUARIO
        End Get
        Set(ByVal value As Integer)
            CA04_ID_USUARIO = value
        End Set
    End Property

    Public Property UsuarioAlt() As Integer
        Get
            Return CA04_ID_USUARIO_ALT
        End Get
        Set(ByVal value As Integer)
            CA04_ID_USUARIO_ALT = value
        End Set
    End Property

    Public Property SituacaoSolicCorrecaoDoc() As Integer
        Get
            Return RH100_ST_SOLIC_CORREC_DOC
        End Get
        Set(ByVal value As Integer)
            RH100_ST_SOLIC_CORREC_DOC = value
        End Set
    End Property

    Public Property DataSituacao() As String
        Get
            Return RH100_DH_ST_SOLIC_CORREC_DOC
        End Get
        Set(ByVal value As String)
            RH100_DH_ST_SOLIC_CORREC_DOC = value
        End Set
    End Property

    Public Property DataCadastro() As String
        Get
            Return RH100_DH_CADASTRO
        End Get
        Set(ByVal value As String)
            RH100_DH_CADASTRO = value
        End Set
    End Property

    Public Property DataAlteracao() As String
        Get
            Return RH100_DH_ALTERACAO
        End Get
        Set(ByVal value As String)
            RH100_DH_ALTERACAO = value
        End Set
    End Property

    Public Property Obsevacao() As String
        Get
            Return RH100_DS_OBSERVACAO
        End Get
        Set(ByVal value As String)
            RH100_DS_OBSERVACAO = value
        End Set
    End Property

    Public Sub New(Optional ByVal codigo As Integer = 0)
        If codigo > 0 Then
            Obter(codigo)
        End If
    End Sub

    Public Sub Salvar(Optional ByRef transacao As Transacao = Nothing)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT * ")
        strSQL.Append("   FROM RH100_SOLIC_CORREC_DOC")
        strSQL.Append("  WHERE RH100_ID_SOLIC_CORREC_DOC = " & IdSolicCorrecaoDoc)

        dt = cnn.EditarDataTable(strSQL.ToString, transacao)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH100_ID_SOLIC_CORREC_DOC") = ProBanco(RH100_ID_SOLIC_CORREC_DOC, eTipoValor.CHAVE)
        dr("RH96_ID_LICENCA_DOC") = ProBanco(RH96_ID_LICENCA_DOC, eTipoValor.CHAVE)
        dr("RH36_ID_LOTACAO") = ProBanco(RH36_ID_LOTACAO, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO_ALT") = ProBanco(CA04_ID_USUARIO_ALT, eTipoValor.CHAVE)
        dr("RH100_ST_SOLIC_CORREC_DOC") = ProBanco(RH100_ST_SOLIC_CORREC_DOC, eTipoValor.NUMERO_INTEIRO)
        dr("RH100_DH_ST_SOLIC_CORREC_DOC") = ProBanco(RH100_DH_ST_SOLIC_CORREC_DOC, eTipoValor.DATA_COMPLETA)
        dr("RH100_DH_CADASTRO") = ProBanco(RH100_DH_CADASTRO, eTipoValor.DATA_COMPLETA)
        dr("RH100_DH_ALTERACAO") = ProBanco(RH100_DH_ALTERACAO, eTipoValor.DATA_COMPLETA)
        dr("RH100_DS_OBSERVACAO") = ProBanco(RH100_DS_OBSERVACAO, eTipoValor.TEXTO_LIVRE)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal SoliccadastroDoc As String)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT * ")
        strSQL.Append("   FROM RH100_SOLIC_CORREC_DOC")
        strSQL.Append("  WHERE RH100_ID_SOLIC_CORREC_DOC =" & SoliccadastroDoc)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH100_ID_SOLIC_CORREC_DOC = DoBanco(dr("RH100_ID_SOLIC_CORREC_DOC"), eTipoValor.CHAVE)
            RH96_ID_LICENCA_DOC = DoBanco(dr("RH96_ID_LICENCA_DOC"), eTipoValor.CHAVE)
            RH36_ID_LOTACAO = DoBanco(dr("RH36_ID_LOTACAO"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            CA04_ID_USUARIO_ALT = DoBanco(dr("CA04_ID_USUARIO_ALT"), eTipoValor.CHAVE)
            RH100_ST_SOLIC_CORREC_DOC = DoBanco(dr("RH100_ST_SOLIC_CORREC_DOC"), eTipoValor.NUMERO_INTEIRO)
            RH100_DH_ST_SOLIC_CORREC_DOC = DoBanco(dr("RH100_DH_ST_SOLIC_CORREC_DOC"), eTipoValor.DATA_COMPLETA)
            RH100_DH_CADASTRO = DoBanco(dr("RH100_DH_CADASTRO"), eTipoValor.DATA_COMPLETA)
            RH100_DH_ALTERACAO = DoBanco(dr("RH100_DH_ALTERACAO"), eTipoValor.DATA_COMPLETA)
            RH100_DS_OBSERVACAO = DoBanco(dr("RH100_DS_OBSERVACAO"), eTipoValor.TEXTO_LIVRE)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function PesquisarSituacaoSetor(Optional ByVal Sort As String = "",
                              Optional Licenca As Integer = 0,
                              Optional SituacaoDoc As String = "1",
                              Optional NomeCategaria As String = "TRÂMITE",
                              Optional CodigoLotacao As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT DI09.DI09_DESCRICAO,   " & vbCrLf)
        strSQL.Append("        DI09.DI09_RESUMO,   " & vbCrLf)
        strSQL.Append(" 	   DI02.DI02_DESCRICAO,   " & vbCrLf)
        strSQL.Append(" 	   RH95.DI02_ID_TIPO_DOCUMENTO,   " & vbCrLf)
        strSQL.Append(" 	   RH95.RH95_ID_STATUS_CATEG_TIPO_DOC,  " & vbCrLf)
        strSQL.Append(" 	   RH36.RH36_ID_LOTACAO,  " & vbCrLf)
        strSQL.Append(" 	   RH36.RH36_NM_LOTACAO,  " & vbCrLf)
        strSQL.Append(" 	   RH100.RH100_ST_SOLIC_CORREC_DOC   " & vbCrLf)
        strSQL.Append("   FROM RH96_LICENCA_DOC                                            RH96  " & vbCrLf)
        strSQL.Append("   JOIN RH95_STATUS_CATEG_TIPO_DOC                                  RH95 ON RH95.DI02_ID_TIPO_DOCUMENTO  = RH96.DI02_ID_TIPO_DOCUMENTO  " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.DBO.DI02_TIPO_DOCUMENTO DI02 ON DI02.DI02_COD_TIPO_DOCUMENTO = RH95.DI02_ID_TIPO_DOCUMENTO  " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.DBO.DI09_STATUS         DI09 ON DI09.DI09_COD_STATUS         = RH95.DI09_ID_STATUS  " & vbCrLf)
        strSQL.Append("   JOIN RH94_CATEGORIA_DOC                                          RH94 ON RH94.RH94_ID_CATEGORIA_DOC   = RH95.RH94_ID_CATEGORIA_DOC  " & vbCrLf)
        strSQL.Append("   JOIN RH36_LOTACAO                                                RH36 ON RH36.RH36_ID_LOTACAO         = DI09.RH36_ID_LOTACAO  " & vbCrLf)
        strSQL.Append("   JOIN RH100_SOLIC_CORREC_DOC                                     RH100 ON RH100.RH96_ID_LICENCA_DOC    = RH96.RH96_ID_LICENCA_DOC   " & vbCrLf)
        strSQL.Append("  WHERE RH100.RH100_ID_SOLIC_CORREC_DOC IS NOT NULL" & vbCrLf)
        strSQL.Append("    AND RH94.RH94_NM_CATEGORIA_DOC LIKE '%" & NomeCategaria & "%'" & vbCrLf)
        strSQL.Append("    AND RH100.RH100_ST_SOLIC_CORREC_DOC in (" & SituacaoDoc & ")" & vbCrLf)

        If Licenca > 0 Then
            strSQL.Append(" AND RH96.RH30_ID_LICENCA = " & Licenca & vbCrLf)
        End If

        If CodigoLotacao > 0 Then
            strSQL.Append(" AND RH36.RH36_ID_LOTACAO = " & CodigoLotacao & vbCrLf)
        End If

        strSQL.Append(" ORDER BY " & IIf(Sort = "", "RH100.RH100_ID_SOLIC_CORREC_DOC", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarSituacao(Optional ByVal Sort As String = "",
                              Optional SolicCorrecaoDoc As Integer = 0,
                              Optional Licenca As Integer = 0,
                              Optional SituacaoDoc As String = "1",
                              Optional NomeCategaria As String = "servidor",
                              Optional CodigoServidor As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT RH100.RH100_ID_SOLIC_CORREC_DOC, " & vbCrLf)
        strSQL.Append("        RH100.RH100_ST_SOLIC_CORREC_DOC, " & vbCrLf)
        strSQL.Append("        RH96.RH30_ID_LICENCA,  " & vbCrLf)
        strSQL.Append("        RH97.RH97_NM_TITULO_ARQUIVO, " & vbCrLf)
        strSQL.Append("        RH100.RH100_DS_OBSERVACAO " & vbCrLf)
        strSQL.Append("   FROM RH100_SOLIC_CORREC_DOC                                  RH100" & vbCrLf)
        strSQL.Append("   JOIN RH97_LICENCA_DOC_ANEXO                                  RH97 ON RH97.RH96_ID_LICENCA_DOC    = RH100.RH96_ID_LICENCA_DOC " & vbCrLf)
        strSQL.Append("   JOIN RH96_LICENCA_DOC                                        RH96 ON RH96.RH96_ID_LICENCA_DOC    = RH100.RH96_ID_LICENCA_DOC " & vbCrLf)
        strSQL.Append("   JOIN RH95_STATUS_CATEG_TIPO_DOC                              RH95 ON RH95.DI02_ID_TIPO_DOCUMENTO = RH96.DI02_ID_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("   JOIN RH94_CATEGORIA_DOC                                      RH94 ON RH94.RH94_ID_CATEGORIA_DOC  = RH95.RH94_ID_CATEGORIA_DOC " & vbCrLf)
        strSQL.Append("   JOIN RH30_LICENCA                                            RH30 ON RH30.RH30_ID_LICENCA        = RH96.RH30_ID_LICENCA " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[DBO].DI09_STATUS DI09 ON DI09.DI09_COD_STATUS        = RH95.DI09_ID_STATUS  " & vbCrLf)
        strSQL.Append("   JOIN RH36_LOTACAO                                            RH36 ON RH36.RH36_ID_LOTACAO        = DI09.RH36_ID_LOTACAO " & vbCrLf)
        strSQL.Append("  WHERE RH100.RH100_ID_SOLIC_CORREC_DOC IS NOT NULL" & vbCrLf)
        strSQL.Append("    AND RH94.RH94_NM_CATEGORIA_DOC LIKE '%" & NomeCategaria & "%'" & vbCrLf)
        strSQL.Append("    AND RH100.RH100_ST_SOLIC_CORREC_DOC in (" & SituacaoDoc & ")" & vbCrLf)

        If SolicCorrecaoDoc > 0 Then
            strSQL.Append(" AND RH100.RH100_ID_SOLIC_CORREC_DOC = " & SolicCorrecaoDoc & vbCrLf)
        End If

        If Licenca > 0 Then
            strSQL.Append(" AND RH96.RH30_ID_LICENCA = " & Licenca & vbCrLf)
        End If

        If CodigoServidor > 0 Then
            strSQL.Append(" AND RH30.RH02_ID_SERVIDOR = " & CodigoServidor & vbCrLf)
        End If

        strSQL.Append(" ORDER BY " & IIf(Sort = "", "RH100.RH100_ID_SOLIC_CORREC_DOC", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function Pesquisar(Optional ByVal Sort As String = "",
                              Optional SolicCorrecaoDoc As Integer = 0,
                              Optional Licenca As Integer = 0,
                              Optional LicencaDoc As Integer = 0,
                              Optional Usuario As Integer = 0,
                              Optional UsuarioAlt As Integer = 0,
                              Optional SituacaoDoc As String = "",
                              Optional Lotacao As Integer = 0,
                              Optional Lotacao2 As Integer = 0,
                              Optional DataHoraSitSolicCorrecaoDoc As String = "",
                              Optional DataHoraCadastro As String = "",
                              Optional DataHoraAlteracao As String = "",
                              Optional CodigoServidor As Integer = 0,
                              Optional Observacao As String = "",
                              Optional Setor As Boolean = True) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append("  SELECT RH100.RH100_ID_SOLIC_CORREC_DOC,  " & vbCrLf)
        strSQL.Append("         RH96.RH96_ID_LICENCA_DOC,  " & vbCrLf)
        strSQL.Append("         RH36.RH36_NM_LOTACAO,  " & vbCrLf)

        If Setor Then 'PARA O SETOR JURIDICO E PROTOCOLO
            strSQL.Append("          CASE RH100.RH100_ST_SOLIC_CORREC_DOC  " & vbCrLf)
            strSQL.Append("            WHEN 1 THEN 'EM ABERTO' " & vbCrLf)
            strSQL.Append("            WHEN 2 THEN 'EXECUTADA' " & vbCrLf)
            strSQL.Append("            WHEN 3 THEN 'CONTESTADA' " & vbCrLf)
            strSQL.Append("            WHEN 4 THEN 'FINALIZADA' " & vbCrLf)
            strSQL.Append("            WHEN 5 THEN 'CANCELADA' " & vbCrLf)
            strSQL.Append("          END AS NM_SITUACAO,   " & vbCrLf)

        Else ' PARA OS DEMAIS SETORES E SERVIDORES
            strSQL.Append("          CASE RH100.RH100_ST_SOLIC_CORREC_DOC  " & vbCrLf)
            strSQL.Append("            WHEN 1 THEN 'PENDENTE' " & vbCrLf)
            strSQL.Append("            WHEN 2 THEN 'EM ANÁLISE' " & vbCrLf)
            strSQL.Append("            WHEN 3 THEN 'EM ANÁLISE' " & vbCrLf)
            strSQL.Append("            WHEN 4 THEN 'VÁLIDA' " & vbCrLf)
            strSQL.Append("            WHEN 5 THEN 'CANCELADA' " & vbCrLf)
            strSQL.Append("          END AS NM_SITUACAO, " & vbCrLf)
        End If

        strSQL.Append("       RH96.RH30_ID_LICENCA, " & vbCrLf)
        strSQL.Append("       RH96.DI02_ID_TIPO_DOCUMENTO, " & vbCrLf)
        strSQL.Append("       RH97.RH97_ID_LICENCA_DOC_ANEXO, " & vbCrLf)
        strSQL.Append("       RH97.RH97_NM_CAMINHO_ARQUIVO,  " & vbCrLf)
        strSQL.Append("       RH97.RH97_SG_EXTENSAO_ARQUIVO, " & vbCrLf)
        strSQL.Append("       RH100.RH100_DS_OBSERVACAO,   " & vbCrLf)
        strSQL.Append("       DI02.DI02_DESCRICAO, " & vbCrLf)
        strSQL.Append("       RH100.RH100_DH_CADASTRO " & vbCrLf)
        strSQL.Append("  FROM RH100_SOLIC_CORREC_DOC                                      RH100 " & vbCrLf)
        strSQL.Append("  JOIN RH96_LICENCA_DOC                                            RH96  ON RH96.RH96_ID_LICENCA_DOC     = RH100.RH96_ID_LICENCA_DOC   " & vbCrLf)
        strSQL.Append("  JOIN RH97_LICENCA_DOC_ANEXO                                      RH97  ON RH97.RH96_ID_LICENCA_DOC     = RH100.RH96_ID_LICENCA_DOC   " & vbCrLf)
        strSQL.Append("  JOIN RH95_STATUS_CATEG_TIPO_DOC                                  RH95  ON RH95.DI02_ID_TIPO_DOCUMENTO  = RH96.DI02_ID_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("  JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.DBO.DI02_TIPO_DOCUMENTO DI02  ON DI02.DI02_COD_TIPO_DOCUMENTO = RH96.DI02_ID_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("  JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.DBO.DI09_STATUS         DI09  ON DI09.DI09_COD_STATUS         = RH95.DI09_ID_STATUS  " & vbCrLf)
        strSQL.Append("  JOIN RH36_LOTACAO                                                RH36  ON RH36.RH36_ID_LOTACAO         = DI09.RH36_ID_LOTACAO " & vbCrLf)
        strSQL.Append("  JOIN RH30_LICENCA                                                RH30  ON RH30.RH30_ID_LICENCA         = RH96.RH30_ID_LICENCA " & vbCrLf)
        strSQL.Append(" WHERE RH100.RH100_ID_SOLIC_CORREC_DOC IS NOT NULL " & vbCrLf)

        If SolicCorrecaoDoc > 0 Then
            strSQL.Append(" AND RH100.RH100_ID_SOLIC_CORREC_DOC = " & SolicCorrecaoDoc & vbCrLf)
        End If

        If Licenca > 0 Then
            strSQL.Append(" AND RH96.RH30_ID_LICENCA = " & Licenca & vbCrLf)
        End If

        If LicencaDoc > 0 Then
            strSQL.Append(" AND RH100.RH96_ID_LICENCA_DOC = " & LicencaDoc & vbCrLf)
        End If

        If Usuario > 0 Then
            strSQL.Append(" AND RH100.CA04_ID_USUARIO = " & Usuario & vbCrLf)
        End If

        If UsuarioAlt > 0 Then
            strSQL.Append(" AND RH100.CA04_ID_USUARIO_ALT = " & UsuarioAlt & vbCrLf)
        End If

        If SituacaoDoc <> "" Then
            strSQL.Append(" AND RH100.RH100_ST_SOLIC_CORREC_DOC in (" & SituacaoDoc & ")" & vbCrLf) ' campo in (1) - filtro padrão.
        End If

        If Lotacao > 0 Then
            strSQL.Append(" AND RH36.RH36_ID_LOTACAO in (" & Lotacao & "," & Lotacao2 & ")" & vbCrLf)
        End If

        If DataHoraSitSolicCorrecaoDoc <> "" Then
            strSQL.Append(" AND RH100.RH100_DH_ST_SOLIC_CORREC_DOC = " & DataHoraSitSolicCorrecaoDoc & vbCrLf)
        End If

        If DataHoraCadastro <> "" Then
            strSQL.Append(" AND RH100.RH100_DH_CADASTRO = " & DataHoraCadastro & vbCrLf)
        End If

        If DataHoraAlteracao <> "" Then
            strSQL.Append(" AND RH100.RH100_DH_ALTERACAO = " & DataHoraAlteracao & vbCrLf)
        End If

        If CodigoServidor > 0 Then
            strSQL.Append(" AND RH30.RH02_ID_SERVIDOR = " & CodigoServidor & vbCrLf)
        End If

        If Observacao <> "" Then
            strSQL.Append(" AND UPPER(RH100.RH100_DS_OBSERVACAO) LIKE '%" & Observacao.ToUpper & "%'" & vbCrLf)
        End If

        strSQL.Append(" ORDER BY " & IIf(Sort = "", "RH100.RH100_ID_SOLIC_CORREC_DOC", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarPorSolicitacao(Optional ByVal Sort As String = "",
                              Optional SolicCorrecaoDoc As Integer = 0,
                              Optional LicencaDoc As Integer = 0,
                              Optional Lotacao As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT TOP 10 RH100.RH100_ID_SOLIC_CORREC_DOC," & vbCrLf)
        strSQL.Append("        RH100.RH96_ID_LICENCA_DOC " & vbCrLf)
        strSQL.Append("   FROM RH100_SOLIC_CORREC_DOC RH100" & vbCrLf)
        strSQL.Append("   JOIN RH96_LICENCA_DOC       RH96 ON RH96.RH96_ID_LICENCA_DOC = RH100.RH96_ID_LICENCA_DOC " & vbCrLf)
        strSQL.Append("   JOIN RH36_LOTACAO           RH36 ON RH36.RH36_ID_LOTACAO     = RH100.RH36_ID_LOTACAO " & vbCrLf)
        strSQL.Append("  WHERE RH100_ID_SOLIC_CORREC_DOC IS NOT NULL " & vbCrLf)

        If SolicCorrecaoDoc > 0 Then
            strSQL.Append(" AND RH100.RH100_ID_SOLIC_CORREC_DOC = " & SolicCorrecaoDoc & vbCrLf)
        End If

        If LicencaDoc > 0 Then
            strSQL.Append(" AND RH100.RH96_ID_LICENCA_DOC = " & LicencaDoc & vbCrLf)
        End If

        If Lotacao > 0 Then
            strSQL.Append(" AND RH36.RH36_ID_LOTACAO = " & Lotacao & vbCrLf)
        End If

        strSQL.Append(" ORDER BY " & IIf(Sort = "", "RH100.RH100_ID_SOLIC_CORREC_DOC", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT RH100_ID_SOLIC_CORREC_DOC as CODIGO, RH96_ID_LICENCA_DOC as DESCRICAO")
        strSQL.Append("   FROM RH100_SOLIC_CORREC_DOC")
        strSQL.Append("  ORDER BY ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return dt
    End Function


    Public Function ObterLotacaoSolicCorrec(Codigolotacao As Integer, Optional ByRef transacao As Transacao = Nothing) As Boolean
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim retorno As Boolean

        strSQL.Append(" SELECT RH36.RH36_NM_LOTACAO  " & vbCrLf)
        strSQL.Append("   FROM RH100_SOLIC_CORREC_DOC                                  RH100 " & vbCrLf)
        strSQL.Append("   JOIN RH97_LICENCA_DOC_ANEXO                                  RH97 ON RH97.RH96_ID_LICENCA_DOC    = RH100.RH96_ID_LICENCA_DOC  " & vbCrLf)
        strSQL.Append("   JOIN RH96_LICENCA_DOC                                        RH96 ON RH96.RH96_ID_LICENCA_DOC    = RH100.RH96_ID_LICENCA_DOC  " & vbCrLf)
        strSQL.Append("   JOIN RH95_STATUS_CATEG_TIPO_DOC                              RH95 ON RH95.DI02_ID_TIPO_DOCUMENTO = RH96.DI02_ID_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("   JOIN RH94_CATEGORIA_DOC                                      RH94 ON RH94.RH94_ID_CATEGORIA_DOC  = RH95.RH94_ID_CATEGORIA_DOC  " & vbCrLf)
        strSQL.Append("   JOIN RH30_LICENCA                                            RH30 ON RH30.RH30_ID_LICENCA        = RH96.RH30_ID_LICENCA  " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[DBO].DI09_STATUS DI09 ON DI09.DI09_COD_STATUS        = RH95.DI09_ID_STATUS  " & vbCrLf)
        strSQL.Append("   JOIN RH36_LOTACAO                                            RH36 ON RH36.RH36_ID_LOTACAO        = DI09.RH36_ID_LOTACAO " & vbCrLf)
        strSQL.Append("  WHERE RH100.RH100_ID_SOLIC_CORREC_DOC IS NOT NULL " & vbCrLf)
        strSQL.Append("    AND RH94.RH94_NM_CATEGORIA_DOC LIKE '%TRÂMITE%' " & vbCrLf)
        strSQL.Append("    AND RH100.RH100_ST_SOLIC_CORREC_DOC in (1,2,3) " & vbCrLf)
        strSQL.Append("    AND RH36.RH36_ID_LOTACAO = " & Codigolotacao & vbCrLf)
        strSQL.Append("  ORDER BY RH100.RH100_ID_SOLIC_CORREC_DOC " & vbCrLf)

        Try

            With cnn.AbrirDataTable(strSQL.ToString)
                If .Rows.Count > 0 Then
                    retorno = True
                Else
                    retorno = False
                End If
            End With

        Catch ex As Exception
            Dim erro As String = ex.ToString
            retorno = False
        End Try

        cnn = Nothing
        Return retorno
    End Function

    Public Function ObterUltimo(Optional ByRef transacao As Transacao = Nothing) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" SELECT MAX(RH100_ID_SOLIC_CORREC_DOC) ")
        strSQL.Append("   FROM RH100_SOLIC_CORREC_DOC")

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

    Public Function ObterUltimoPorUsuario(CodigoUsuario As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" SELECT MAX(RH100_ID_SOLIC_CORREC_DOC) ")
        strSQL.Append("   FROM RH100_SOLIC_CORREC_DOC")
        strSQL.Append("   WHERE CA04_ID_USUARIO = " & CodigoUsuario)

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

    Public Function Excluir(ByVal SolicCorrecaoDoc As String) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" DELECT ")
        strSQL.Append("   FROM RH100_SOLIC_CORREC_DOC")
        strSQL.Append("  WHERE RH100_ID_SOLIC_CORREC_DOC = " & SolicCorrecaoDoc)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

    Public Function Ultimo() As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim retorno As Integer

        strSQL.Append(" SELECT IDENT_CURRENT('RH100_SOLIC_CORREC_DOC')")

        retorno = cnn.AbrirDataTable(strSQL.ToString)(0)(0)

        cnn = Nothing
        strSQL = Nothing

        Return retorno
    End Function

End Class
