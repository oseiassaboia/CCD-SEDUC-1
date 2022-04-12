Imports System.Data
Imports Microsoft.VisualBasic

Public Class TipoDocumento
    Implements IDisposable

    Private DI02_COD_TIPO_DOCUMENTO As Integer
    Private DI02_DESCRICAO As String
    Private DI02_DARE As Boolean
    Private DI02_IND_EXIGE_DT_EMISSAO As Boolean
    Private DI02_IND_EXIGE_DT_VALIDADE As Boolean

    Private disposedValue As Boolean

    Public Property Codigo() As Integer
        Get
            Return DI02_COD_TIPO_DOCUMENTO
        End Get
        Set(value As Integer)
            DI02_COD_TIPO_DOCUMENTO = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return DI02_DESCRICAO
        End Get
        Set(value As String)
            DI02_DESCRICAO = value
        End Set
    End Property

    Public Property Dare() As Boolean
        Get
            Return DI02_DARE
        End Get
        Set(value As Boolean)
            DI02_DARE = value
        End Set
    End Property

    Public Property ExigeDtEmissao() As Boolean
        Get
            Return DI02_IND_EXIGE_DT_EMISSAO
        End Get
        Set(value As Boolean)
            DI02_IND_EXIGE_DT_EMISSAO = value
        End Set
    End Property

    Public Property ExigeDtValidade() As Boolean
        Get
            Return DI02_IND_EXIGE_DT_VALIDADE
        End Get
        Set(value As Boolean)
            DI02_IND_EXIGE_DT_VALIDADE = value
        End Set
    End Property

    Public Sub New(Optional IdTipoDocumento As Integer = 0)
        If IdTipoDocumento > 0 Then
            Obter(IdTipoDocumento)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DI02_TIPO_DOCUMENTO")
        strSQL.Append(" where DI02_COD_TIPO_DOCUMENTO  = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("DI02_COD_TIPO_DOCUMENTO") = ProBanco(DI02_COD_TIPO_DOCUMENTO, eTipoValor.CHAVE)
        dr("DI02_DESCRICAO") = ProBanco(DI02_DESCRICAO, eTipoValor.TEXTO)
        dr("DI02_DARE") = ProBanco(DI02_DARE, eTipoValor.BOOLEANO)
        dr("DI02_IND_EXIGE_DT_EMISSAO") = ProBanco(DI02_IND_EXIGE_DT_EMISSAO, eTipoValor.BOOLEANO)
        dr("DI02_IND_EXIGE_DT_VALIDADE") = ProBanco(DI02_IND_EXIGE_DT_VALIDADE, eTipoValor.BOOLEANO)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(idTipoProcesso As Integer)
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append("   from DI02_TIPO_DOCUMENTO")
        strSQL.Append("  where DI02_COD_TIPO_DOCUMENTO = " & idTipoProcesso)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            DI02_COD_TIPO_DOCUMENTO = DoBanco(dr("DI02_COD_TIPO_DOCUMENTO"), eTipoValor.CHAVE)
            DI02_DESCRICAO = DoBanco(dr("DI02_DESCRICAO"), eTipoValor.TEXTO)
            DI02_DARE = DoBanco(dr("DI02_DARE"), eTipoValor.BOOLEANO)
            DI02_IND_EXIGE_DT_EMISSAO = DoBanco(dr("DI02_IND_EXIGE_DT_EMISSAO"), eTipoValor.BOOLEANO)
            DI02_IND_EXIGE_DT_VALIDADE = DoBanco(dr("DI02_IND_EXIGE_DT_VALIDADE"), eTipoValor.BOOLEANO)

        End If

        cnn.FecharBanco()
    End Sub

    Public Function Pesquisar(Optional Sort As String = "",
                              Optional IdTipoDocumento As Integer = 0,
                              Optional Descricao As String = "",
                              Optional Dare As Boolean = Nothing,
                              Optional ExigeDtEmissao As Boolean = Nothing,
                              Optional ExigeDtValidade As Boolean = Nothing) As DataTable

        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT *")
        strSQL.Append(" FROM DI02_TIPO_DOCUMENTO")
        strSQL.Append(" WHERE DI02_COD_TIPO_DOCUMENTO IS NOT NULL")

        If IdTipoDocumento > 0 Then
            strSQL.Append(" and DI02_COD_TIPO_DOCUMENTO = " & IdTipoDocumento)
        End If

        If Descricao <> "" Then
            strSQL.Append(" and upper(DI02_DESCRICAO) like '%" & Descricao.ToUpper & "%'")
        End If

        If Dare <> Nothing Then
            strSQL.Append(" and  DI02_DARE = " & Dare)
        End If

        If ExigeDtEmissao <> Nothing Then
            strSQL.Append(" and  DI02_IND_EXIGE_DT_EMISSAO = " & ExigeDtEmissao)
        End If

        If ExigeDtValidade <> Nothing Then
            strSQL.Append(" and  DI02_IND_EXIGE_DT_VALIDADE = " & ExigeDtValidade)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "DI02_COD_TIPO_DOCUMENTO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function ObterDocumentos(ListDocumento As String) As DataTable

        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT *")
        strSQL.Append(" FROM DI02_TIPO_DOCUMENTO")
        strSQL.Append(" WHERE DI02_COD_TIPO_DOCUMENTO IS NOT NULL")
        If ListDocumento <> "" Then
            strSQL.Append(" AND DI02_COD_TIPO_DOCUMENTO IN ( " & ListDocumento & " )")
        Else
            strSQL.Append(" AND DI02_COD_TIPO_DOCUMENTO IN ( 0 )")
        End If


        strSQL.Append(" Order By DI02_COD_TIPO_DOCUMENTO")

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function ObterTabela(Optional ByVal Documento As Integer = 0) As DataTable
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT DI02_COD_TIPO_DOCUMENTO CODIGO, DI02_DESCRICAO DESCRICAO")
        strSQL.Append("   FROM DI02_TIPO_DOCUMENTO")
        strSQL.Append("  WHERE DI02_COD_TIPO_DOCUMENTO IS NOT NULL ")
        If Documento > 0 Then
            strSQL.Append("  AND DI02_COD_TIPO_DOCUMENTO = " & Documento)
        End If
        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return dt

    End Function

    'Public Function ObterTabela(Optional ByVal Documento As Integer = 0) As DataTable
    '    Dim cnn As New Conexao("StringConexao")
    '    Dim dt As DataTable
    '    Dim strSQL As New StringBuilder

    '    strSQL.Append(" SELECT DI02.DI02_COD_TIPO_DOCUMENTO CODIGO, DI02.DI02_DESCRICAO DESCRICAO")
    '    strSQL.Append("   FROM RH95_STATUS_CATEG_TIPO_DOC RH95 ")
    '    strSQL.Append("   JOIN [172.16.2.71].DBDIARIAS_TREINAMENTO.DBO.DI02_TIPO_DOCUMENTO DI02 ON DI02.DI02_COD_TIPO_DOCUMENTO = RH95.DI02_ID_TIPO_DOCUMENTO")
    '    strSQL.Append("  WHERE RH95.DI02_ID_TIPO_DOCUMENTO IS NOT NULL ")
    '    If Documento > 0 Then
    '        strSQL.Append("  AND RH95.DI02_ID_TIPO_DOCUMENTO = " & Documento)
    '    End If
    '    dt = cnn.AbrirDataTable(strSQL.ToString)

    '    cnn.FecharBanco()
    '    cnn = Nothing

    '    Return dt

    'End Function

    Public Function ObterUltimo() As Integer

        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(DI17_COD_TIPO_PROCESSO) from DI17_TIPO_PROCESSO")

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
