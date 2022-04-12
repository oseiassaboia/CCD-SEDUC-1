Imports System.Data
Imports Microsoft.VisualBasic

Public Class CategoriaDocumento
    Implements IDisposable

    Private RH94_ID_CATEGORIA_DOC As Integer
    Private RH93_ID_TIPO_LICENCA_PROCESSO As Integer
    Private RH94_NM_CATEGORIA_DOC As String

    Private disposedValue As Boolean

    Public Property Codigo() As Integer
        Get
            Return RH94_ID_CATEGORIA_DOC
        End Get
        Set(value As Integer)
            RH94_ID_CATEGORIA_DOC = value
        End Set
    End Property

    Public Property TipoLicencaProcesso() As Integer
        Get
            Return RH93_ID_TIPO_LICENCA_PROCESSO
        End Get
        Set(value As Integer)
            RH93_ID_TIPO_LICENCA_PROCESSO = value
        End Set
    End Property

    Public Property NomeCategoria() As String
        Get
            Return RH94_NM_CATEGORIA_DOC
        End Get
        Set(value As String)
            RH94_NM_CATEGORIA_DOC = value
        End Set
    End Property

    Public Sub New(Optional IdCategoriaDocumento As Integer = 0)
        If IdCategoriaDocumento > 0 Then
            Obter(IdCategoriaDocumento)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH94_CATEGORIA_DOC")
        strSQL.Append(" where RH94_ID_CATEGORIA_DOC = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH94_ID_CATEGORIA_DOC") = ProBanco(RH94_ID_CATEGORIA_DOC, eTipoValor.CHAVE)
        dr("RH93_ID_TIPO_LICENCA_PROCESSO") = ProBanco(RH93_ID_TIPO_LICENCA_PROCESSO, eTipoValor.CHAVE)
        dr("RH94_NM_CATEGORIA_DOC") = ProBanco(RH94_NM_CATEGORIA_DOC, eTipoValor.TEXTO_LIVRE)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(CodigoCategoriaDocumento As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH94_CATEGORIA_DOC")
        strSQL.Append(" where RH94_ID_CATEGORIA_DOC = " & CodigoCategoriaDocumento)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH94_ID_CATEGORIA_DOC = DoBanco(dr("RH94_ID_CATEGORIA_DOC"), eTipoValor.CHAVE)
            RH93_ID_TIPO_LICENCA_PROCESSO = DoBanco(dr("RH93_ID_TIPO_LICENCA_PROCESSO"), eTipoValor.CHAVE)
            RH94_NM_CATEGORIA_DOC = DoBanco(dr("RH94_NM_CATEGORIA_DOC"), eTipoValor.TEXTO_LIVRE)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional Sort As String = "",
                              Optional IdCategoriaDocumento As Integer = 0,
                              Optional IdTipoLicencaProcesso As Integer = 0,
                              Optional NomeCategoriaDocumento As String = "") As DataTable

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select *")
        strSQL.Append(" from RH94_CATEGORIA_DOC")
        strSQL.Append(" where RH94_ID_CATEGORIA_DOC is not null")

        If IdCategoriaDocumento > 0 Then
            strSQL.Append(" and RH94_ID_CATEGORIA_DOC = " & IdCategoriaDocumento)
        End If

        If IdTipoLicencaProcesso > 0 Then
            strSQL.Append(" and RH93_ID_TIPO_LICENCA_PROCESSO = " & IdTipoLicencaProcesso)
        End If

        If NomeCategoriaDocumento <> "" Then
            strSQL.Append(" and upper(RH94_NM_CATEGORIA_DOC) like '%" & NomeCategoriaDocumento.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH94_ID_CATEGORIA_DOC", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function ObterTabela(Optional IdTipoLicencaProcesso As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select RH94_ID_CATEGORIA_DOC as CODIGO, RH94_NM_CATEGORIA_DOC as DESCRICAO")
        strSQL.Append(" from RH94_CATEGORIA_DOC")

        If IdTipoLicencaProcesso > 0 Then
            strSQL.Append(" WHERE RH93_ID_TIPO_LICENCA_PROCESSO = " & IdTipoLicencaProcesso)
        End If

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

        strSQL.Append(" select max(RH94_ID_CATEGORIA_DOC) from RH94_CATEGORIA_DOC")

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
