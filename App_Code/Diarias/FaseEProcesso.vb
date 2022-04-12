Imports System.Data
Imports Microsoft.VisualBasic

Public Class FaseEProcesso
    Implements IDisposable

    Private DI21_COD_FASE_EPROCESSO As Integer
    Private DI21_DESCRICAO As String

    Private disposedValue As Boolean

    Public Property Codigo() As Integer
        Get
            Return DI21_COD_FASE_EPROCESSO
        End Get
        Set(value As Integer)
            DI21_COD_FASE_EPROCESSO = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return DI21_DESCRICAO
        End Get
        Set(value As String)
            DI21_DESCRICAO = value
        End Set
    End Property

    Public Sub New(Optional ByVal Codigo As Integer = 0)
        If Codigo > 0 Then
            Obter(Codigo)
        End If
    End Sub



    Public Sub Obter(Codigo As Integer)
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DI21_FASE_EPROCESSO")
        strSQL.Append(" where DI21_COD_FASE_EPROCESSO = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            DI21_COD_FASE_EPROCESSO = DoBanco(dr("DI21_COD_FASE_EPROCESSO"), eTipoValor.CHAVE)
            DI21_DESCRICAO = DoBanco(dr("DI21_DESCRICAO"), eTipoValor.TEXTO_LIVRE)

        End If

        cnn.FecharBanco()
    End Sub

    Public Function ObterUltimo() As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(DI21_COD_FASE_EPROCESSO) from DI21_FASE_EPROCESSO")

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

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select DI21_COD_FASE_EPROCESSO as CODIGO, DI21_DESCRICAO as DESCRICAO")
        strSQL.Append(" from DI21_FASE_EPROCESSO")

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
