Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class DocumentoLicenca
    Implements IDisposable
    Private RH96_ID_LICENCA_DOC As Integer
    Private RH30_ID_LICENCA As Integer
    Private DI02_ID_TIPO_DOCUMENTO As Integer
    Private TG43_ID_ORGAO_EMISSOR As Integer
    Private CA04_ID_USUARIO As Integer
    Private CA04_ID_USUARIO_ALT As Integer
    Private RH96_NU_DOCUMENTO As String
    Private RH96_SG_UF_ORGAO_EMISSOR As String
    Private RH96_DT_EMISSAO As String
    Private RH96_DT_VALIDADE As String
    Private RH96_DH_CADASTRO As String
    Private RH96_DH_ALTERACAO As String

    Private disposedValue As Boolean

#Region "Getters e  Setters"

    Public Property Codigo As Integer
        Get
            Return RH96_ID_LICENCA_DOC
        End Get
        Set(value As Integer)
            RH96_ID_LICENCA_DOC = value
        End Set
    End Property

    Public Property IdLicenca As Integer
        Get
            Return RH30_ID_LICENCA
        End Get
        Set(value As Integer)
            RH30_ID_LICENCA = value
        End Set
    End Property

    Public Property TipoDocumento As Integer
        Get
            Return DI02_ID_TIPO_DOCUMENTO
        End Get
        Set(value As Integer)
            DI02_ID_TIPO_DOCUMENTO = value
        End Set
    End Property

    Public Property OrgaoEmissor As Integer
        Get
            Return TG43_ID_ORGAO_EMISSOR
        End Get
        Set(value As Integer)
            TG43_ID_ORGAO_EMISSOR = value
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

    Public Property IdUsuarioAlt As Integer
        Get
            Return CA04_ID_USUARIO_ALT
        End Get
        Set(value As Integer)
            CA04_ID_USUARIO_ALT = value
        End Set
    End Property

    Public Property NumeroDocumento As String
        Get
            Return RH96_NU_DOCUMENTO
        End Get
        Set(value As String)
            RH96_NU_DOCUMENTO = value
        End Set
    End Property

    Public Property SgOrgaoEmissor As String
        Get
            Return RH96_SG_UF_ORGAO_EMISSOR
        End Get
        Set(value As String)
            RH96_SG_UF_ORGAO_EMISSOR = value
        End Set
    End Property

    Public Property DataEmissao As String
        Get
            Return RH96_DT_EMISSAO
        End Get
        Set(value As String)
            RH96_DT_EMISSAO = value
        End Set
    End Property

    Public Property DataValidade As String
        Get
            Return RH96_DT_VALIDADE
        End Get
        Set(value As String)
            RH96_DT_VALIDADE = value
        End Set
    End Property

    Public Property DataCadastro As String
        Get
            Return RH96_DH_CADASTRO
        End Get
        Set(value As String)
            RH96_DH_CADASTRO = value
        End Set
    End Property

    Public Property DataAlteracao As String
        Get
            Return RH96_DH_ALTERACAO
        End Get
        Set(value As String)
            RH96_DH_ALTERACAO = value
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
        strSQL.Append(" from RH96_LICENCA_DOC")
        strSQL.Append(" where RH96_ID_LICENCA_DOC = " & codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH96_ID_LICENCA_DOC = DoBanco(dr("RH96_ID_LICENCA_DOC"), eTipoValor.CHAVE)
            RH30_ID_LICENCA = DoBanco(dr("RH30_ID_LICENCA"), eTipoValor.CHAVE)
            DI02_ID_TIPO_DOCUMENTO = DoBanco(dr("DI02_ID_TIPO_DOCUMENTO"), eTipoValor.CHAVE)
            TG43_ID_ORGAO_EMISSOR = DoBanco(dr("TG43_ID_ORGAO_EMISSOR"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            CA04_ID_USUARIO_ALT = DoBanco(dr("CA04_ID_USUARIO_ALT"), eTipoValor.CHAVE)
            RH96_NU_DOCUMENTO = DoBanco(dr("RH96_NU_DOCUMENTO"), eTipoValor.TEXTO_LIVRE)
            RH96_SG_UF_ORGAO_EMISSOR = DoBanco(dr("RH96_SG_UF_ORGAO_EMISSOR"), eTipoValor.TEXTO)
            RH96_DT_EMISSAO = DoBanco(dr("RH96_DT_EMISSAO"), eTipoValor.DATA)
            RH96_DT_VALIDADE = DoBanco(dr("RH96_DT_VALIDADE"), eTipoValor.DATA)
            RH96_DH_CADASTRO = DoBanco(dr("RH96_DH_CADASTRO"), eTipoValor.DATA_COMPLETA)
            RH96_DH_ALTERACAO = DoBanco(dr("RH96_DH_ALTERACAO"), eTipoValor.DATA_COMPLETA)

        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "",
                              Optional Codigo As Integer = 0,
                              Optional IdLicenca As Integer = 0,
                              Optional IdTipoDocumento As Integer = 0,
                              Optional IdOrgaoEmissor As Integer = 0,
                              Optional IdUsuario As Integer = 0,
                              Optional IdUsuarioAlt As Integer = 0,
                              Optional NumDocumento As String = "",
                              Optional SgUfOrgaoEmissor As String = "") As DataTable

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH96_LICENCA_DOC")
        strSQL.Append(" where RH96_ID_LICENCA_DOC is not null")

        If Codigo > 0 Then
            strSQL.Append(" AND RH96_ID_LICENCA_DOC = " & Codigo)
        End If

        If IdLicenca > 0 Then
            strSQL.Append(" AND RH30_ID_LICENCA = " & IdLicenca)
        End If

        If IdTipoDocumento > 0 Then
            strSQL.Append(" AND DI02_ID_TIPO_DOCUMENTO = " & IdTipoDocumento)
        End If

        If IdOrgaoEmissor > 0 Then
            strSQL.Append(" AND TG43_ID_ORGAO_EMISSOR = " & IdOrgaoEmissor)
        End If

        If IdUsuario > 0 Then
            strSQL.Append(" AND CA04_ID_USUARIO = " & IdUsuario)
        End If

        If IdUsuarioAlt > 0 Then
            strSQL.Append(" AND CA04_ID_USUARIO_ALT = " & IdUsuarioAlt)
        End If

        If NumDocumento <> "" Then
            strSQL.Append(" AND RH96_NU_DOCUMENTO = " & NumDocumento)
        End If

        If SgUfOrgaoEmissor <> "" Then
            strSQL.Append(" AND RH96_SG_UF_ORGAO_EMISSOR = " & SgUfOrgaoEmissor)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH96_ID_LICENCA_DOC", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Sub Salvar(Optional ByRef transacao As Transacao = Nothing)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH96_LICENCA_DOC")
        strSQL.Append(" where RH96_ID_LICENCA_DOC = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString, transacao)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH96_ID_LICENCA_DOC") = ProBanco(RH96_ID_LICENCA_DOC, eTipoValor.CHAVE)
        dr("RH30_ID_LICENCA") = ProBanco(RH30_ID_LICENCA, eTipoValor.CHAVE)
        dr("DI02_ID_TIPO_DOCUMENTO") = ProBanco(DI02_ID_TIPO_DOCUMENTO, eTipoValor.CHAVE)
        dr("TG43_ID_ORGAO_EMISSOR") = ProBanco(TG43_ID_ORGAO_EMISSOR, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO_ALT") = ProBanco(CA04_ID_USUARIO_ALT, eTipoValor.CHAVE)
        dr("RH96_NU_DOCUMENTO") = ProBanco(RH96_NU_DOCUMENTO, eTipoValor.TEXTO_LIVRE)
        dr("RH96_SG_UF_ORGAO_EMISSOR") = ProBanco(RH96_SG_UF_ORGAO_EMISSOR, eTipoValor.TEXTO)
        dr("RH96_DT_EMISSAO") = ProBanco(RH96_DT_EMISSAO, eTipoValor.DATA)
        dr("RH96_DT_VALIDADE") = ProBanco(RH96_DT_VALIDADE, eTipoValor.DATA)
        dr("RH96_DH_CADASTRO") = ProBanco(RH96_DH_CADASTRO, eTipoValor.DATA)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function ObterUltimo(Optional ByRef transacao As Transacao = Nothing) As Integer

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(RH96_ID_LICENCA_DOC) from RH96_LICENCA_DOC")

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

    Public Function ObterLicencaDoc(Optional ByVal Sort As String = "",
                                    Optional Codigo As Integer = 0,
                                    Optional IdLicenca As Integer = 0,
                                    Optional IdTipoDocumento As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT RH102.RH102_ID_SOLIC_CAD_DOC,   ")
        strSQL.Append("        RH30.RH30_ID_LICENCA,   ")
        strSQL.Append(" 	   RH102.DI02_ID_TIPO_DOCUMENTO,   ")
        strSQL.Append(" 	   DI02.DI02_DESCRICAO,   ")
        strSQL.Append(" 	   RH102.RH102_DS_OBSERVACAO   ")
        strSQL.Append("   FROM RH102_SOLIC_CAD_DOC RH102   ")
        strSQL.Append("   JOIN RH30_LICENCA        RH30  ON RH30.RH30_ID_LICENCA = RH102.RH30_ID_LICENCA    ")
        strSQL.Append("   JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.[DBO].[DI02_TIPO_DOCUMENTO] DI02  ON DI02.DI02_COD_TIPO_DOCUMENTO = RH102.DI02_ID_TIPO_DOCUMENTO   ")
        strSQL.Append("  WHERE RH102.RH102_ID_SOLIC_CAD_DOC IS NOT NULL   ")

        If Codigo > 0 Then
            strSQL.Append(" AND RH102.RH102_ID_SOLIC_CAD_DOC = " & Codigo)
        End If

        If IdLicenca > 0 Then
            strSQL.Append(" AND RH30.RH30_ID_LICENCA = " & IdLicenca)
        End If

        If IdTipoDocumento > 0 Then
            strSQL.Append(" AND DI02.DI02_ID_TIPO_DOCUMENTO = " & IdTipoDocumento)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH102.RH102_ID_SOLIC_CAD_DOC", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela(Optional NomeLicenca As String = "", Optional idLicenca As Integer = 1) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT DI02_ID_TIPO_DOCUMENTO as CODIGO, DI02_DESCRICAO as DESCRICAO   " & vbCrLf)
        strSQL.Append("   FROM RH95_STATUS_CATEG_TIPO_DOC                                      RH95   " & vbCrLf)
        strSQL.Append("   JOIN RH94_CATEGORIA_DOC                                              RH94 ON RH94.RH94_ID_CATEGORIA_DOC          = RH95.RH94_ID_CATEGORIA_DOC   " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[dbo].DI02_TIPO_DOCUMENTO DI02 ON RH95.DI02_ID_TIPO_DOCUMENTO         = DI02.DI02_COD_TIPO_DOCUMENTO  " & vbCrLf)
        strSQL.Append("   JOIN [172.16.2.71].[DBDIARIAS_TREINAMENTO].[dbo].DI09_STATUS         DI09 ON DI09.DI09_COD_STATUS                = RH95.DI09_ID_STATUS   " & vbCrLf)
        strSQL.Append("   JOIN RH93_TIPO_LICENCA_PROCESSO                                      RH93 ON RH93.RH93_ID_TIPO_LICENCA_PROCESSO  = RH94.RH93_ID_TIPO_LICENCA_PROCESSO   " & vbCrLf)
        strSQL.Append("   JOIN RH36_LOTACAO                                                    RH36 ON RH36.RH36_ID_LOTACAO                = DI09.RH36_ID_LOTACAO  " & vbCrLf)
        strSQL.Append("   JOIN RH29_TIPO_LICENCA                                               RH29 ON RH29.RH29_ID_TIPO_LICENCA           = RH93.RH29_ID_TIPO_LICENCA   " & vbCrLf)
        strSQL.Append("  WHERE RH94.RH94_NM_CATEGORIA_DOC like '%trâmite%'    " & vbCrLf)


        If NomeLicenca <> "" Then
            strSQL.Append("  AND UPPER(RH29.RH29_NM_TIPO_LICENCA) LIKE '%" & NomeLicenca.ToUpper & "%'" & vbCrLf)
        End If

        If idLicenca > 0 Then
            strSQL.Append("  AND RH29.RH29_ID_TIPO_LICENCA = " & idLicenca & vbCrLf)
        End If

        strSQL.Append("    ORDER BY RH95.RH95_ID_STATUS_CATEG_TIPO_DOC    " & vbCrLf)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return dt
    End Function

    Public Function ObterLicencaDocumento(ByVal TipoDocumento As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoLotacao As Integer

        strSQL.Append(" SELECT TOP 1 RH96_ID_LICENCA_DOC   " & vbCrLf)
        strSQL.Append("   FROM RH96_LICENCA_DOC " & vbCrLf)
        strSQL.Append("  WHERE RH96_ID_LICENCA_DOC IS NOT NULL " & vbCrLf)
        strSQL.Append("   AND DI02_ID_TIPO_DOCUMENTO  = " & TipoDocumento)

        With cnn.AbrirDataTable(strSQL.ToString)
            If Not IsDBNull(.Rows(0)(0)) Then
                CodigoLotacao = .Rows(0)(0)
            Else
                CodigoLotacao = 0
            End If
        End With

        cnn.FecharBanco()
        cnn = Nothing

        Return CodigoLotacao

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
