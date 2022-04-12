Imports Microsoft.VisualBasic
Imports System.Data

Public Class SolicitarCadastradoDoc

    Private RH102_ID_SOLIC_CAD_DOC As Integer
    Private RH30_ID_LICENCA As Integer
    Private DI02_ID_TIPO_DOCUMENTO As Integer
    Private CA04_ID_USUARIO As Integer
    Private CA04_ID_USUARIO_ALT As Integer
    Private RH102_DH_CADASTRO As String
    Private RH102_ST_SOLIC_CAD_DOC As Integer
    Private RH102_DH_ST_SOLIC_CAD_DOC As String
    Private RH102_DS_OBSERVACAO As String


#Region "Getters e Setters"
    Public Property IdSolicCadastroDoc() As Integer
        Get
            Return RH102_ID_SOLIC_CAD_DOC
        End Get
        Set(ByVal value As Integer)
            RH102_ID_SOLIC_CAD_DOC = value
        End Set
    End Property

    Public Property IdLicenca() As Integer
        Get
            Return RH30_ID_LICENCA
        End Get
        Set(ByVal value As Integer)
            RH30_ID_LICENCA = value
        End Set
    End Property

    Public Property TipoDocumento() As Integer
        Get
            Return DI02_ID_TIPO_DOCUMENTO
        End Get
        Set(ByVal value As Integer)
            DI02_ID_TIPO_DOCUMENTO = value
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

    Public Property SituacaoSolicCadastroDoc() As Integer
        Get
            Return RH102_ST_SOLIC_CAD_DOC
        End Get
        Set(ByVal value As Integer)
            RH102_ST_SOLIC_CAD_DOC = value
        End Set
    End Property

    Public Property DataHoraCadastro() As String
        Get
            Return RH102_DH_CADASTRO
        End Get
        Set(ByVal value As String)
            RH102_DH_CADASTRO = value
        End Set
    End Property

    Public Property DataHoraSituacao() As String
        Get
            Return RH102_DH_ST_SOLIC_CAD_DOC
        End Get
        Set(ByVal value As String)
            RH102_DH_ST_SOLIC_CAD_DOC = value
        End Set
    End Property

    Public Property Obsevacao() As String
        Get
            Return RH102_DS_OBSERVACAO
        End Get
        Set(ByVal value As String)
            RH102_DS_OBSERVACAO = value
        End Set
    End Property

#End Region

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
        strSQL.Append("   FROM RH102_SOLIC_CAD_DOC")
        strSQL.Append("  WHERE RH102_ID_SOLIC_CAD_DOC = " & IdSolicCadastroDoc)

        dt = cnn.EditarDataTable(strSQL.ToString, transacao)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH102_ID_SOLIC_CAD_DOC") = ProBanco(RH102_ID_SOLIC_CAD_DOC, eTipoValor.CHAVE)
        dr("RH30_ID_LICENCA") = ProBanco(RH30_ID_LICENCA, eTipoValor.CHAVE)
        dr("DI02_ID_TIPO_DOCUMENTO") = ProBanco(DI02_ID_TIPO_DOCUMENTO, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("RH102_DH_CADASTRO") = ProBanco(RH102_DH_CADASTRO, eTipoValor.DATA_COMPLETA)
        dr("CA04_ID_USUARIO_ALT") = ProBanco(CA04_ID_USUARIO_ALT, eTipoValor.CHAVE)
        dr("RH102_ST_SOLIC_CAD_DOC") = ProBanco(RH102_ST_SOLIC_CAD_DOC, eTipoValor.NUMERO_INTEIRO)
        dr("RH102_DH_ST_SOLIC_CAD_DOC") = ProBanco(RH102_DH_ST_SOLIC_CAD_DOC, eTipoValor.DATA_COMPLETA)
        dr("RH102_DS_OBSERVACAO") = ProBanco(RH102_DS_OBSERVACAO, eTipoValor.TEXTO_LIVRE)

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
        strSQL.Append("   FROM RH102_SOLIC_CAD_DOC")
        strSQL.Append("  WHERE RH102_ID_SOLIC_CAD_DOC = " & SoliccadastroDoc)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH102_ID_SOLIC_CAD_DOC = DoBanco(dr("RH102_ID_SOLIC_CAD_DOC"), eTipoValor.CHAVE)
            RH30_ID_LICENCA = DoBanco(dr("RH30_ID_LICENCA"), eTipoValor.CHAVE)
            DI02_ID_TIPO_DOCUMENTO = DoBanco(dr("DI02_ID_TIPO_DOCUMENTO"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            RH102_DH_CADASTRO = DoBanco(dr("RH102_DH_CADASTRO"), eTipoValor.DATA_COMPLETA)
            CA04_ID_USUARIO_ALT = DoBanco(dr("CA04_ID_USUARIO_ALT"), eTipoValor.CHAVE)
            RH102_ST_SOLIC_CAD_DOC = DoBanco(dr("RH102_ST_SOLIC_CAD_DOC"), eTipoValor.NUMERO_INTEIRO)
            RH102_DH_ST_SOLIC_CAD_DOC = DoBanco(dr("RH102_DH_ST_SOLIC_CAD_DOC"), eTipoValor.DATA_COMPLETA)
            RH102_DS_OBSERVACAO = DoBanco(dr("RH102_DS_OBSERVACAO"), eTipoValor.TEXTO_LIVRE)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional Sort As String = "",
                              Optional CodigoDocSolic As Integer = 0,
                              Optional IdLicenca As Integer = 0,
                              Optional Lotacao As Integer = 0,
                              Optional SituacaoDoc As String = "",
                              Optional Setor As Boolean = True) As DataTable

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT RH102.RH102_ID_SOLIC_CAD_DOC, " & vbCrLf)
        strSQL.Append("        RH102.DI02_ID_TIPO_DOCUMENTO, " & vbCrLf)

        If Setor Then 'PARA O SETOR JURIDICO 
            strSQL.Append("          CASE RH102.RH102_ST_SOLIC_CAD_DOC  " & vbCrLf)
            strSQL.Append("            WHEN 1 THEN 'EM ABERTO'  " & vbCrLf)
            strSQL.Append("            WHEN 2 THEN 'EXECUTADA'  " & vbCrLf)
            strSQL.Append("            WHEN 3 THEN 'CONTESTADA'  " & vbCrLf)
            strSQL.Append("            WHEN 4 THEN 'FINALIZADA'  " & vbCrLf)
            strSQL.Append("            WHEN 5 THEN 'CANCELADA'  " & vbCrLf)
            strSQL.Append("          END AS NM_SITUACAO,   " & vbCrLf)

        Else ' PARA OS DEMAIS SETORES E SERVIDORES
            strSQL.Append("          CASE RH102.RH102_ST_SOLIC_CAD_DOC  " & vbCrLf)
            strSQL.Append("            WHEN 1 THEN 'PENDENTE'  " & vbCrLf)
            strSQL.Append("            WHEN 2 THEN 'EM ANÁLISE'  " & vbCrLf)
            strSQL.Append("            WHEN 3 THEN 'EM ANÁLISE'  " & vbCrLf)
            strSQL.Append("            WHEN 4 THEN 'VÁLIDA'  " & vbCrLf)
            strSQL.Append("            WHEN 5 THEN 'CANCELADA'  " & vbCrLf)
            strSQL.Append("          END AS NM_SITUACAO,   " & vbCrLf)
        End If

        strSQL.Append("        RH30.RH30_ID_LICENCA, " & vbCrLf)
        strSQL.Append(" 	   DI02.DI02_DESCRICAO, " & vbCrLf)
        strSQL.Append("        RH102.RH102_DS_OBSERVACAO, " & vbCrLf)
        strSQL.Append("        RH102.RH102_DH_CADASTRO " & vbCrLf)
        strSQL.Append("   FROM RH102_SOLIC_CAD_DOC                                         RH102  " & vbCrLf)
        strSQL.Append("   JOIN RH30_LICENCA                                                RH30  ON RH30.RH30_ID_LICENCA         = RH102.RH30_ID_LICENCA " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.DBO.DI02_TIPO_DOCUMENTO DI02  ON DI02.DI02_COD_TIPO_DOCUMENTO = RH102.DI02_ID_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("   JOIN RH95_STATUS_CATEG_TIPO_DOC                                  RH95 ON RH95.DI02_ID_TIPO_DOCUMENTO   = DI02.DI02_COD_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.[DBO].[DI09_STATUS]     DI09 ON DI09.DI09_COD_STATUS          = RH95.DI09_ID_STATUS  " & vbCrLf)
        strSQL.Append("   JOIN RH36_LOTACAO                                                RH36 ON RH36.RH36_ID_LOTACAO          = DI09.RH36_ID_LOTACAO  " & vbCrLf)
        strSQL.Append("  WHERE RH102.RH102_ID_SOLIC_CAD_DOC IS NOT NULL " & vbCrLf)

        If CodigoDocSolic > 0 Then
            strSQL.Append(" AND RH102.RH102_ID_SOLIC_CAD_DOC = " & CodigoDocSolic & vbCrLf)
        End If

        If IdLicenca > 0 Then
            strSQL.Append(" AND RH30.RH30_ID_LICENCA = " & IdLicenca & vbCrLf)
        End If

        If Lotacao > 0 Then
            strSQL.Append(" AND RH36.RH36_ID_LOTACAO= " & Lotacao & vbCrLf)
        End If

        If SituacaoDoc <> "" Then
            strSQL.Append(" AND RH102.RH102_ST_SOLIC_CAD_DOC in (" & SituacaoDoc & ")" & vbCrLf) ' campo in (1) - filtro padrão.
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH102.RH102_ID_SOLIC_CAD_DOC", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function PesquisarSituacao(Optional ByVal Sort As String = "",
                              Optional SolicitacaoDoc As Integer = 0,
                              Optional Licenca As Integer = 0,
                              Optional SituacaoDoc As String = "1",
                              Optional Lotacao As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder


        strSQL.Append(" SELECT DI02.DI02_DESCRICAO, " & vbCrLf)
        strSQL.Append("        RH102.RH102_DS_OBSERVACAO, " & vbCrLf)
        strSQL.Append(" 	   RH30.RH30_ID_LICENCA, " & vbCrLf)
        strSQL.Append("        RH36.RH36_NM_LOTACAO   " & vbCrLf)
        strSQL.Append("   FROM RH102_SOLIC_CAD_DOC                                         RH102    " & vbCrLf)
        strSQL.Append("   JOIN RH30_LICENCA                                                RH30  ON RH30.RH30_ID_LICENCA         = RH102.RH30_ID_LICENCA   " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.DBO.DI02_TIPO_DOCUMENTO DI02  ON DI02.DI02_COD_TIPO_DOCUMENTO = RH102.DI02_ID_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("   JOIN RH95_STATUS_CATEG_TIPO_DOC                                  RH95 ON RH95.DI02_ID_TIPO_DOCUMENTO   = DI02.DI02_COD_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.[DBO].[DI09_STATUS]     DI09 ON DI09.DI09_COD_STATUS          = RH95.DI09_ID_STATUS   " & vbCrLf)
        strSQL.Append("   JOIN RH36_LOTACAO                                                RH36 ON RH36.RH36_ID_LOTACAO          = DI09.RH36_ID_LOTACAO    " & vbCrLf)
        strSQL.Append("  WHERE RH102.RH102_ID_SOLIC_CAD_DOC IS NOT NULL " & vbCrLf)
        strSQL.Append("    AND RH102.RH102_ST_SOLIC_CAD_DOC in (" & SituacaoDoc & ")" & vbCrLf)

        If SolicitacaoDoc > 0 Then
            strSQL.Append(" AND RH102.RH102_ID_SOLIC_CAD_DOC = " & SolicitacaoDoc & vbCrLf)
        End If

        If Licenca > 0 Then
            strSQL.Append(" AND RH30.RH30_ID_LICENCA = " & Licenca & vbCrLf)
        End If

        If Lotacao > 0 Then
            strSQL.Append("  AND RH36.RH36_ID_LOTACAO = " & Lotacao & vbCrLf)
        End If

        strSQL.Append(" ORDER BY " & IIf(Sort = "", "RH102.RH102_ST_SOLIC_CAD_DOC", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarPorSolicitacaoDoc(Optional ByVal Sort As String = "",
                              Optional SolicitarDoc As Integer = 0,
                              Optional TipoDoc As Integer = 0,
                              Optional Licenca As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT TOP 1 RH102.RH102_ID_SOLIC_CAD_DOC   " & vbCrLf)
        strSQL.Append("   FROM RH102_SOLIC_CAD_DOC                                         RH102  " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.DBO.DI02_TIPO_DOCUMENTO DI02  ON DI02.DI02_COD_TIPO_DOCUMENTO = RH102.DI02_ID_TIPO_DOCUMENTO  " & vbCrLf)
        strSQL.Append("   JOIN RH30_LICENCA                                                RH30  ON RH30.RH30_ID_LICENCA         = RH102.RH30_ID_LICENCA      " & vbCrLf)
        strSQL.Append("  WHERE RH102.RH102_ID_SOLIC_CAD_DOC IS NOT NULL  " & vbCrLf)

        If SolicitarDoc > 0 Then
            strSQL.Append(" AND RH102.RH102_ID_SOLIC_CAD_DOC = " & SolicitarDoc & vbCrLf)
        End If

        If Licenca > 0 Then
            strSQL.Append("    AND RH30.RH30_ID_LICENCA = " & Licenca & vbCrLf)
        End If

        If TipoDoc > 0 Then
            strSQL.Append("    AND RH102.DI02_ID_TIPO_DOCUMENTO = " & TipoDoc & vbCrLf)
        End If

        strSQL.Append(" ORDER BY " & IIf(Sort = "", "RH102.RH102_ID_SOLIC_CAD_DOC", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT RH102_ID_SOLIC_CAD_DOC as CODIGO, RH07_NM_SITUACAO_SERVIDOR as DESCRICAO")
        strSQL.Append("   FROM RH102_SOLIC_CAD_DOC")
        strSQL.Append("  ORDER BY ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return dt
    End Function

    Public Function ObterLotacaoSolicDoc(Codigolotacao As Integer, Optional ByRef transacao As Transacao = Nothing) As Boolean
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim retorno As Boolean

        strSQL.Append(" SELECT RH36.RH36_NM_LOTACAO  " & vbCrLf)
        strSQL.Append("   FROM RH102_SOLIC_CAD_DOC                                         RH102   " & vbCrLf)
        strSQL.Append("   JOIN RH30_LICENCA                                                RH30  ON RH30.RH30_ID_LICENCA         = RH102.RH30_ID_LICENCA  " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.DBO.DI02_TIPO_DOCUMENTO DI02  ON DI02.DI02_COD_TIPO_DOCUMENTO = RH102.DI02_ID_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("   JOIN RH95_STATUS_CATEG_TIPO_DOC                                  RH95 ON RH95.DI02_ID_TIPO_DOCUMENTO   = DI02.DI02_COD_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.[DBO].[DI09_STATUS]     DI09 ON DI09.DI09_COD_STATUS          = RH95.DI09_ID_STATUS  " & vbCrLf)
        strSQL.Append("   JOIN RH36_LOTACAO                                                RH36 ON RH36.RH36_ID_LOTACAO          = DI09.RH36_ID_LOTACAO   " & vbCrLf)
        strSQL.Append("  WHERE RH102.RH102_ID_SOLIC_CAD_DOC IS NOT NULL  " & vbCrLf)
        strSQL.Append("    AND RH102.RH102_ST_SOLIC_CAD_DOC in (1,2,3)  " & vbCrLf)
        strSQL.Append("    AND RH36.RH36_ID_LOTACAO = " & Codigolotacao & vbCrLf)
        strSQL.Append("  ORDER BY RH102.RH102_ST_SOLIC_CAD_DOC " & vbCrLf)

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

        strSQL.Append(" select max(RH102_ID_SOLIC_CAD_DOC) from RH102_SOLIC_CAD_DOC")

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

        strSQL.Append(" SELECT MAX(RH102_ID_SOLIC_CAD_DOC) ")
        strSQL.Append("   FROM RH102_SOLIC_CAD_DOC")
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

    Public Function Excluir(ByVal SoliccadastroDoc As String) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" DELECT ")
        strSQL.Append("   FROM RH102_SOLIC_CAD_DOC")
        strSQL.Append("  WHERE RH102_ID_SOLIC_CAD_DOC = " & SoliccadastroDoc)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class
