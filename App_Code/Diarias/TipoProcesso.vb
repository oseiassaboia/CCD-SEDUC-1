Imports System.Data
Imports Microsoft.VisualBasic

Public Class TipoProcesso
    Implements IDisposable

    Private DI17_COD_TIPO_PROCESSO As Integer
    Private DI17_COD_EPROCESSO As Integer
    Private DI17_DESCRICAO As String
    Private FKDI17PT25_COD_TIPO_PROCESSO As Integer

    Private disposedValue As Boolean

    Public Property Codigo() As Integer
        Get
            Return DI17_COD_TIPO_PROCESSO
        End Get
        Set(value As Integer)
            DI17_COD_TIPO_PROCESSO = value
        End Set
    End Property

    Public Property CodEprocesso() As Integer
        Get
            Return DI17_COD_EPROCESSO
        End Get
        Set(value As Integer)
            DI17_COD_EPROCESSO = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return DI17_DESCRICAO
        End Get
        Set(value As String)
            DI17_DESCRICAO = value
        End Set
    End Property

    Public Property CodTipoEprocesso() As Integer
        Get
            Return FKDI17PT25_COD_TIPO_PROCESSO
        End Get
        Set(value As Integer)
            FKDI17PT25_COD_TIPO_PROCESSO = value
        End Set
    End Property

    Public Sub New(Optional IdTipoProcesso As Integer = 0)
        If IdTipoProcesso > 0 Then
            Obter(IdTipoProcesso)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DI17_TIPO_PROCESSO")
        strSQL.Append(" where DI17_COD_TIPO_PROCESSO  = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("DI17_COD_TIPO_PROCESSO") = ProBanco(DI17_COD_TIPO_PROCESSO, eTipoValor.CHAVE)
        dr("DI17_COD_EPROCESSO") = ProBanco(DI17_COD_EPROCESSO, eTipoValor.NUMERO_INTEIRO)
        dr("DI17_DESCRICAO") = ProBanco(DI17_DESCRICAO, eTipoValor.TEXTO_LIVRE)
        dr("FKDI17PT25_COD_TIPO_PROCESSO") = ProBanco(FKDI17PT25_COD_TIPO_PROCESSO, eTipoValor.CHAVE)

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
        strSQL.Append(" from DI17_TIPO_PROCESSO")
        strSQL.Append(" where DI17_COD_TIPO_PROCESSO = " & idTipoProcesso)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            DI17_COD_TIPO_PROCESSO = DoBanco(dr("DI17_COD_TIPO_PROCESSO"), eTipoValor.CHAVE)
            DI17_COD_EPROCESSO = DoBanco(dr("DI17_COD_EPROCESSO"), eTipoValor.NUMERO_INTEIRO)
            DI17_DESCRICAO = DoBanco(dr("DI17_DESCRICAO"), eTipoValor.TEXTO_LIVRE)
            FKDI17PT25_COD_TIPO_PROCESSO = DoBanco(dr("FKDI17PT25_COD_TIPO_PROCESSO"), eTipoValor.NUMERO_INTEIRO)

        End If

        cnn.FecharBanco()
    End Sub

    Public Function Pesquisar(Optional Sort As String = "",
                              Optional IdTipoLicenca As Integer = 0,
                              Optional CodEprocesso As Integer = 0,
                              Optional Descricao As String = "",
                              Optional FKDI17PT25_COD_TIPO_PROCESSO As Integer = 0) As DataTable

        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT *")
        strSQL.Append(" FROM DI17_TIPO_PROCESSO")
        strSQL.Append(" WHERE DI17_COD_TIPO_PROCESSO IS NOT NULL")

        If IdTipoLicenca > 0 Then
            strSQL.Append(" and DI17_COD_TIPO_PROCESSO = " & IdTipoLicenca)
        End If

        If CodEprocesso > 0 Then
            strSQL.Append(" and DI17_COD_EPROCESSO = " & CodEprocesso)
        End If

        If Descricao <> "" Then
            strSQL.Append(" and upper(DI17_DESCRICAO) like '%" & Descricao.ToUpper & "%'")
        End If

        If FKDI17PT25_COD_TIPO_PROCESSO > 0 Then
            strSQL.Append(" and FKDI17PT25_COD_TIPO_PROCESSO = " & FKDI17PT25_COD_TIPO_PROCESSO)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "DI17_COD_TIPO_PROCESSO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function ObterTabela(Optional CodTipoProcesso As Integer = 0) As DataTable
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select DI17_COD_TIPO_PROCESSO as CODIGO, DI17_DESCRICAO as DESCRICAO")
        strSQL.Append(" from DI17_TIPO_PROCESSO")

        If (CodTipoProcesso > 0) Then
            strSQL.Append(" WHERE FKDI17PT25_COD_TIPO_PROCESSO =" & CodTipoProcesso)
        End If

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return dt

    End Function

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
