Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class DocumentoLicencaAnexo
    Implements IDisposable
    Private RH97_ID_LICENCA_DOC_ANEXO As Integer
    Private RH96_ID_LICENCA_DOC As Integer
    Private CA04_ID_USUARIO As Integer
    Private RH97_DH_CADASTRO As String
    Private RH97_NM_CAMINHO_ARQUIVO As String
    Private RH97_SG_EXTENSAO_ARQUIVO As String
    Private RH97_NM_TITULO_ARQUIVO As String
    Private RH97_DS_OBSERVACAO As String
    Private disposedValue As Boolean

#Region "Getters e  Setters"
    Public Property Codigo As Integer
        Get
            Return RH97_ID_LICENCA_DOC_ANEXO
        End Get
        Set(value As Integer)
            RH97_ID_LICENCA_DOC_ANEXO = value
        End Set
    End Property

    Public Property IdLicencaDoc As Integer
        Get
            Return RH96_ID_LICENCA_DOC
        End Get
        Set(value As Integer)
            RH96_ID_LICENCA_DOC = value
        End Set
    End Property

    Public Property DataCadastro As String
        Get
            Return RH97_DH_CADASTRO
        End Get
        Set(value As String)
            RH97_DH_CADASTRO = value
        End Set
    End Property

    Public Property CaminhoArquivo As String
        Get
            Return RH97_NM_CAMINHO_ARQUIVO
        End Get
        Set(value As String)
            RH97_NM_CAMINHO_ARQUIVO = value
        End Set
    End Property

    Public Property ExtensaoArquivo As String
        Get
            Return RH97_SG_EXTENSAO_ARQUIVO
        End Get
        Set(value As String)
            RH97_SG_EXTENSAO_ARQUIVO = value
        End Set
    End Property

    Public Property TituloArquivo As String
        Get
            Return RH97_NM_TITULO_ARQUIVO
        End Get
        Set(value As String)
            RH97_NM_TITULO_ARQUIVO = value
        End Set
    End Property

    Public Property Observacao As String
        Get
            Return RH97_DS_OBSERVACAO
        End Get
        Set(value As String)
            RH97_DS_OBSERVACAO = value
        End Set
    End Property

    Public Property IdUsuario As Integer
        Get
            Return CA04_ID_USUARIO
        End Get
        Set(value As Integer)
            CA04_ID_USUARIO = value
        End Set
    End Property

#End Region

    Public Sub New(Optional ByVal Codigo As Integer = 0)
        If Codigo > 0 Then
            Obter(Codigo)
        End If
    End Sub

    Private Sub Obter(codigo As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH97_LICENCA_DOC_ANEXO")
        strSQL.Append(" where RH97_ID_LICENCA_DOC_ANEXO = " & codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH97_ID_LICENCA_DOC_ANEXO = DoBanco(dr("RH97_ID_LICENCA_DOC_ANEXO"), eTipoValor.CHAVE)
            RH96_ID_LICENCA_DOC = DoBanco(dr("RH96_ID_LICENCA_DOC"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            RH97_DH_CADASTRO = DoBanco(dr("RH97_DH_CADASTRO"), eTipoValor.CHAVE)
            RH97_NM_CAMINHO_ARQUIVO = DoBanco(dr("RH97_NM_CAMINHO_ARQUIVO"), eTipoValor.TEXTO_LIVRE)
            RH97_SG_EXTENSAO_ARQUIVO = DoBanco(dr("RH97_SG_EXTENSAO_ARQUIVO"), eTipoValor.TEXTO_LIVRE)
            RH97_NM_TITULO_ARQUIVO = DoBanco(dr("RH97_NM_TITULO_ARQUIVO"), eTipoValor.TEXTO_LIVRE)
            RH97_DS_OBSERVACAO = DoBanco(dr("RH97_DS_OBSERVACAO"), eTipoValor.TEXTO_LIVRE)

        End If

        cnn.FecharBanco()
    End Sub

    Public Function CarregaGridDocumentoSolicitado(Optional Sort As String = "",
                                                   Optional CodigoDocSolic As Integer = 0,
                                                   Optional IdLicenca As Integer = 0,
                                                   Optional Lotacao As Integer = 0) As DataTable

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT RH102.RH102_ID_SOLIC_CAD_DOC, " & vbCrLf)
        strSQL.Append("        RH102.DI02_ID_TIPO_DOCUMENTO, " & vbCrLf)
        strSQL.Append("        CASE RH102.RH102_ST_SOLIC_CAD_DOC " & vbCrLf)
        strSQL.Append("           WHEN 1 THEN 'ENVIADA'  " & vbCrLf)
        strSQL.Append("           WHEN 2 THEN 'EXECUTADA'  " & vbCrLf)
        strSQL.Append("           WHEN 3 THEN 'INVIÁVEL'  " & vbCrLf)
        strSQL.Append("           WHEN 4 THEN 'NÃO PROCEDE'  " & vbCrLf)
        strSQL.Append("           WHEN 5 THEN 'REENVIADA'  " & vbCrLf)
        strSQL.Append("           WHEN 6 THEN 'CANCELADA'  " & vbCrLf)
        strSQL.Append("         END AS NM_SITUACAO, " & vbCrLf)
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
            strSQL.Append(" AND RH102.RH102_ID_SOLIC_CAD_DOC = " & CodigoDocSolic)
        End If

        If IdLicenca > 0 Then
            strSQL.Append(" AND RH30.RH30_ID_LICENCA = " & IdLicenca)
        End If

        If Lotacao > 0 Then
            strSQL.Append(" AND RH36.RH36_ID_LOTACAO= " & Lotacao)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH102.RH102_ID_SOLIC_CAD_DOC", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function VerificaAnexoDoc(ByVal Licenca As Integer, ByVal TipoDoc As String) As Boolean
        Dim retorno As Boolean = False
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT RH96.RH96_ID_LICENCA_DOC,   " & vbCrLf)
        strSQL.Append(" 	   RH97.RH97_NM_TITULO_ARQUIVO, " & vbCrLf)
        strSQL.Append(" 	   RH95.DI09_ID_STATUS " & vbCrLf)
        strSQL.Append("   FROM RH97_LICENCA_DOC_ANEXO RH97 " & vbCrLf)
        strSQL.Append("   JOIN RH96_LICENCA_DOC                                                RH96 ON RH96.RH96_ID_LICENCA_DOC    = RH97.RH96_ID_LICENCA_DOC " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.[DBO].[DI02_TIPO_DOCUMENTO] DI02 ON RH96.DI02_ID_TIPO_DOCUMENTO = DI02.DI02_COD_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("   JOIN RH95_STATUS_CATEG_TIPO_DOC                                      RH95 ON RH95.DI02_ID_TIPO_DOCUMENTO = DI02.DI02_COD_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.[DBO].[DI09_STATUS]         DI09 ON DI09.DI09_COD_STATUS        = RH95.DI09_ID_STATUS    " & vbCrLf)
        strSQL.Append("  WHERE RH97.RH97_ID_LICENCA_DOC_ANEXO IS NOT NULL " & vbCrLf)
        strSQL.Append("    AND RH96.RH30_ID_LICENCA = " & Licenca & vbCrLf)
        strSQL.Append("    AND RH96.DI02_ID_TIPO_DOCUMENTO in (" & TipoDoc & ")" & vbCrLf)
        strSQL.Append("  Order By RH97.RH97_ID_LICENCA_DOC_ANEXO")

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

    Public Function Pesquisar(Optional Sort As String = "",
                              Optional Codigo As Integer = 0,
                              Optional IdLicencaDoc As Integer = 0,
                              Optional IdLicenca As Integer = 0,
                              Optional IdLotacaoDoc As Integer = 0,
                              Optional Caminho As String = "",
                              Optional Extensao As String = "",
                              Optional Titulo As String = "",
                              Optional Observacao As String = "",
                              Optional NomeCategaria As String = "servidor") As DataTable

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT RH97.RH97_ID_LICENCA_DOC_ANEXO, " & vbCrLf)
        strSQL.Append("        RH96.RH96_ID_LICENCA_DOC, " & vbCrLf)
        strSQL.Append("        DI02.DI02_DESCRICAO, " & vbCrLf)
        strSQL.Append("        RH97.RH97_NM_CAMINHO_ARQUIVO, " & vbCrLf)
        strSQL.Append("        RH97.RH97_SG_EXTENSAO_ARQUIVO " & vbCrLf)
        strSQL.Append("   FROM RH97_LICENCA_DOC_ANEXO                                          RH97" & vbCrLf)
        strSQL.Append("   JOIN RH96_LICENCA_DOC                                                RH96 ON RH96.RH96_ID_LICENCA_DOC    = RH97.RH96_ID_LICENCA_DOC " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.[DBO].[DI02_TIPO_DOCUMENTO] DI02 ON RH96.DI02_ID_TIPO_DOCUMENTO = DI02.DI02_COD_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("   JOIN RH95_STATUS_CATEG_TIPO_DOC                                      RH95 ON RH95.DI02_ID_TIPO_DOCUMENTO = DI02.DI02_COD_TIPO_DOCUMENTO " & vbCrLf)
        strSQL.Append("   JOIN RH94_CATEGORIA_DOC                                              RH94 ON RH94.RH94_ID_CATEGORIA_DOC  = RH95.RH94_ID_CATEGORIA_DOC " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.[DBO].[DI09_STATUS]         DI09 ON DI09.DI09_COD_STATUS        = RH95.DI09_ID_STATUS " & vbCrLf)
        strSQL.Append("  WHERE RH97.RH97_ID_LICENCA_DOC_ANEXO is not null")

        If Codigo > 0 Then
            strSQL.Append(" AND RH97.RH97_ID_LICENCA_DOC_ANEXO = " & Codigo & vbCrLf)
        End If

        If IdLicencaDoc > 0 Then
            strSQL.Append(" AND RH97.RH96_ID_LICENCA_DOC = " & IdLicencaDoc & vbCrLf)
        End If

        If IdLicenca > 0 Then
            strSQL.Append(" AND RH96.RH30_ID_LICENCA = " & IdLicenca & vbCrLf)
        End If

        If IdLotacaoDoc > 0 Then
            strSQL.Append(" AND DI09.RH36_ID_LOTACAO = " & IdLotacaoDoc & vbCrLf)
        End If

        If Caminho <> "" Then
            strSQL.Append(" AND UPPER(RH97.RH97_NM_CAMINHO_ARQUIVO) LIKE '%" & Caminho & "%'" & vbCrLf)
        End If

        If Extensao <> "" Then
            strSQL.Append(" AND UPPER(RH97.RH97_SG_EXTENSAO_ARQUIVO) LIKE '%" & Extensao & "%'" & vbCrLf)
        End If

        If Titulo <> "" Then
            strSQL.Append(" AND UPPER(RH97_NM_TITULO_ARQUIVO) LIKE '%" & Titulo & "%'" & vbCrLf)
        End If

        If Observacao <> "" Then
            strSQL.Append(" AND UPPER(RH97.RH97_DS_OBSERVACAO) LIKE '%" & Observacao & "%'" & vbCrLf)
        End If

        If NomeCategaria <> "" Then
            strSQL.Append(" AND RH94.RH94_NM_CATEGORIA_DOC LIKE '%" & NomeCategaria & "%'" & vbCrLf)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH97.RH97_ID_LICENCA_DOC_ANEXO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Sub Salvar(Optional ByRef transacao As Transacao = Nothing)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH97_LICENCA_DOC_ANEXO")
        strSQL.Append(" where RH97_ID_LICENCA_DOC_ANEXO = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString, transacao)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH97_ID_LICENCA_DOC_ANEXO") = ProBanco(RH97_ID_LICENCA_DOC_ANEXO, eTipoValor.CHAVE)
        dr("RH96_ID_LICENCA_DOC") = ProBanco(RH96_ID_LICENCA_DOC, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("RH97_DH_CADASTRO") = ProBanco(RH97_DH_CADASTRO, eTipoValor.DATA_COMPLETA)
        dr("RH97_NM_CAMINHO_ARQUIVO") = ProBanco(RH97_NM_CAMINHO_ARQUIVO, eTipoValor.TEXTO_LIVRE)
        dr("RH97_SG_EXTENSAO_ARQUIVO") = ProBanco(RH97_SG_EXTENSAO_ARQUIVO, eTipoValor.TEXTO_LIVRE)
        dr("RH97_NM_TITULO_ARQUIVO") = ProBanco(RH97_NM_TITULO_ARQUIVO, eTipoValor.TEXTO_LIVRE)
        dr("RH97_DS_OBSERVACAO") = ProBanco(RH97_DS_OBSERVACAO, eTipoValor.TEXTO_LIVRE)

        cnn.SalvarDataTable(dr)

        dt.Dispose()

        cnn.FecharBanco()
    End Sub

    Public Function ObterUltimo(Optional ByRef transacao As Transacao = Nothing) As Integer

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(RH97_ID_LICENCA_DOC_ANEXO) from RH97_LICENCA_DOC_ANEXO")

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

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects)
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override finalizer
            ' TODO: set large fields to null
            disposedValue = True
        End If
    End Sub

    ' ' TODO: override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
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
End Class
