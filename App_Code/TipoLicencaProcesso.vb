Imports System.Data
Imports Microsoft.VisualBasic

Public Class TipoLicencaProcesso
    Implements IDisposable

    Private RH93_ID_TIPO_LICENCA_PROCESSO As Integer
    Private DI17_ID_TIPO_PROCESSO As Integer
    Private RH29_ID_TIPO_LICENCA As Integer

    Private disposedValue As Boolean

    Public Property Codigo() As Integer
        Get
            Return RH93_ID_TIPO_LICENCA_PROCESSO
        End Get
        Set(value As Integer)
            RH93_ID_TIPO_LICENCA_PROCESSO = value
        End Set
    End Property

    Public Property TipoProcesso() As Integer
        Get
            Return DI17_ID_TIPO_PROCESSO
        End Get
        Set(value As Integer)
            DI17_ID_TIPO_PROCESSO = value
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

    Public Sub New(Optional ByVal IdTipoLicencaProcesso As Integer = 0)
        If IdTipoLicencaProcesso > 0 Then
            Obter(IdTipoLicencaProcesso)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT * ")
        strSQL.Append("   FROM RH93_TIPO_LICENCA_PROCESSO")
        strSQL.Append("  WHERE RH93_ID_TIPO_LICENCA_PROCESSO  = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH93_ID_TIPO_LICENCA_PROCESSO") = ProBanco(RH93_ID_TIPO_LICENCA_PROCESSO, eTipoValor.CHAVE)
        dr("DI17_ID_TIPO_PROCESSO") = ProBanco(DI17_ID_TIPO_PROCESSO, eTipoValor.CHAVE)
        dr("RH29_ID_TIPO_LICENCA") = ProBanco(RH29_ID_TIPO_LICENCA, eTipoValor.CHAVE)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(idTipoLicencaProcesso As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH93_TIPO_LICENCA_PROCESSO")
        strSQL.Append(" where RH93_ID_TIPO_LICENCA_PROCESSO = " & idTipoLicencaProcesso)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH93_ID_TIPO_LICENCA_PROCESSO = DoBanco(dr("RH93_ID_TIPO_LICENCA_PROCESSO"), eTipoValor.CHAVE)
            DI17_ID_TIPO_PROCESSO = DoBanco(dr("DI17_ID_TIPO_PROCESSO"), eTipoValor.CHAVE)
            RH29_ID_TIPO_LICENCA = DoBanco(dr("RH29_ID_TIPO_LICENCA"), eTipoValor.CHAVE)

        End If

        cnn.FecharBanco()
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "",
                              Optional IdTipoLicencaProcesso As Integer = 0,
                              Optional IdTipoProcesso As Integer = 0,
                              Optional IdTipoLicenca As Integer = 0,
                              Optional NomeTipoLIcenca As String = "") As DataTable

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT * " & vbCrLf)
        strSQL.Append("   FROM RH93_TIPO_LICENCA_PROCESSO RH93  " & vbCrLf)
        strSQL.Append("   JOIN RH29_TIPO_LICENCA		  RH29 ON RH29.RH29_ID_TIPO_LICENCA = RH93.RH29_ID_TIPO_LICENCA " & vbCrLf)
        strSQL.Append("  WHERE RH93.RH93_ID_TIPO_LICENCA_PROCESSO IS NOT NULL " & vbCrLf)

        If IdTipoLicencaProcesso > 0 Then
            strSQL.Append(" and RH93_ID_TIPO_LICENCA_PROCESSO = " & IdTipoLicencaProcesso)
        End If

        If IdTipoProcesso > 0 Then
            strSQL.Append(" and DI17_ID_TIPO_PROCESSO = " & IdTipoProcesso)
        End If

        If IdTipoLicenca > 0 Then
            strSQL.Append(" and RH29.RH29_ID_TIPO_LICENCA = " & IdTipoLicenca)
        End If

        If NomeTipoLIcenca <> "" Then
            strSQL.Append(" and upper(RH29.RH29_NM_TIPO_LICENCA) like '%" & NomeTipoLIcenca.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH93_ID_TIPO_LICENCA_PROCESSO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT RH93_ID_TIPO_LICENCA_PROCESSO as CODIGO, RH29_NM_TIPO_LICENCA as DESCRICAO  " & vbCrLf)
        strSQL.Append("   FROM RH93_TIPO_LICENCA_PROCESSO RH93  " & vbCrLf)
        strSQL.Append("   JOIN RH29_TIPO_LICENCA          RH29 ON RH29.RH29_ID_TIPO_LICENCA = RH93.RH29_ID_TIPO_LICENCA " & vbCrLf)
        strSQL.Append("  ORDER bY 2 ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return dt

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
