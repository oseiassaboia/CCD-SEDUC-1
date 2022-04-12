Imports System.Data
Imports Microsoft.VisualBasic

Public Class SequenciaStatus
    Implements IDisposable
    Private DI10_COD_SEQUENCIA_STATUS As Integer
    Private FKDI10DI09_COD_STATUS_ATUAL As Integer
    Private FKDI10DI09_COD_STATUS_PROXIMO As Integer

    Private disposedValue As Boolean

    Public Property Codigo() As Integer
        Get
            Return DI10_COD_SEQUENCIA_STATUS
        End Get
        Set(value As Integer)
            DI10_COD_SEQUENCIA_STATUS = value
        End Set
    End Property

    Public Property StatusAtual() As Integer
        Get
            Return FKDI10DI09_COD_STATUS_ATUAL
        End Get
        Set(value As Integer)
            FKDI10DI09_COD_STATUS_ATUAL = value
        End Set
    End Property

    Public Property StatusSeguinte() As Integer
        Get
            Return FKDI10DI09_COD_STATUS_PROXIMO
        End Get
        Set(value As Integer)
            FKDI10DI09_COD_STATUS_PROXIMO = value
        End Set
    End Property

    Public Sub New(Optional ByVal IdSequenciaStatus As Integer = 0)
        If IdSequenciaStatus > 0 Then
            Obter(IdSequenciaStatus)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DI10_SEQUENCIA_STATUS")
        strSQL.Append(" where DI10_COD_SEQUENCIA_STATUS = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("DI10_COD_SEQUENCIA_STATUS") = ProBanco(DI10_COD_SEQUENCIA_STATUS, eTipoValor.CHAVE)
        dr("FKDI10DI09_COD_STATUS_ATUAL") = ProBanco(FKDI10DI09_COD_STATUS_ATUAL, eTipoValor.CHAVE)
        dr("FKDI10DI09_COD_STATUS_PROXIMO") = ProBanco(FKDI10DI09_COD_STATUS_PROXIMO, eTipoValor.CHAVE)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(IdSequenciaStatus As Integer)
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DI10_SEQUENCIA_STATUS")
        strSQL.Append(" where DI10_COD_SEQUENCIA_STATUS = " & IdSequenciaStatus)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            DI10_COD_SEQUENCIA_STATUS = DoBanco(dr("DI10_COD_SEQUENCIA_STATUS"), eTipoValor.CHAVE)
            FKDI10DI09_COD_STATUS_ATUAL = DoBanco(dr("FKDI10DI09_COD_STATUS_ATUAL"), eTipoValor.CHAVE)
            FKDI10DI09_COD_STATUS_PROXIMO = DoBanco(dr("FKDI10DI09_COD_STATUS_PROXIMO"), eTipoValor.CHAVE)

        End If

        cnn.FecharBanco()
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "",
                              Optional IdSequenciaStatus As Integer = 0,
                              Optional IdStatusAtual As Integer = 0,
                              Optional IdStatusSeguinte As Integer = 0) As DataTable

        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT *")
        strSQL.Append(" FROM DI10_SEQUENCIA_STATUS")

        strSQL.Append(" WHERE DI10_COD_SEQUENCIA_STATUS IS NOT NULL")

        If IdSequenciaStatus > 0 Then
            strSQL.Append(" and DI10_COD_SEQUENCIA_STATUS = " & IdSequenciaStatus)
        End If

        If IdStatusAtual > 0 Then
            strSQL.Append(" and FKDI10DI09_COD_STATUS_ATUAL = " & IdStatusAtual)
        End If

        If IdStatusSeguinte > 0 Then
            strSQL.Append(" and FKDI10DI09_COD_STATUS_PROXIMO = " & IdStatusSeguinte)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "DI10_COD_SEQUENCIA_STATUS", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function MontarGridSequencia(Optional ByVal Sort As String = "",
                              Optional IdSequenciaStatus As Integer = 0,
                              Optional IdStatusAtual As Integer = 0,
                              Optional IdStatusProximo As Integer = 0) As DataTable

        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim strSQL As New StringBuilder

        strSQL.Append(" SELECT DI10_SEQUENCIA_STATUS.DI10_COD_SEQUENCIA_STATUS, ")
        strSQL.Append(" ATUAL.DI09_DESCRICAO + ' - ' + ATUAL.DI09_RESUMO AS ATUAL, ")
        strSQL.Append(" PROXIMO.DI09_DESCRICAO + ' - ' + PROXIMO.DI09_RESUMO AS PROXIMO")
        strSQL.Append(" FROM DI10_SEQUENCIA_STATUS")
        strSQL.Append(" JOIN DI09_STATUS ATUAL ON DI10_SEQUENCIA_STATUS.FKDI10DI09_COD_STATUS_ATUAL = ATUAL.DI09_COD_STATUS")
        strSQL.Append(" JOIN DI09_STATUS PROXIMO ON DI10_SEQUENCIA_STATUS.FKDI10DI09_COD_STATUS_PROXIMO = PROXIMO.DI09_COD_STATUS")
        strSQL.Append(" WHERE DI10_COD_SEQUENCIA_STATUS IS NOT NULL")


        If IdSequenciaStatus > 0 Then
            strSQL.Append(" and DI10_COD_SEQUENCIA_STATUS = " & IdSequenciaStatus)
        End If

        If IdStatusAtual > 0 Then
            strSQL.Append(" and DI10_SEQUENCIA_STATUS.FKDI10DI09_COD_STATUS_ATUAL = " & IdStatusAtual)
        End If

        If IdStatusProximo > 0 Then
            strSQL.Append(" and DI10_SEQUENCIA_STATUS.FKDI10DI09_COD_STATUS_PROXIMO = " & IdStatusProximo)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "DI10_COD_SEQUENCIA_STATUS", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function ObterTabela(Optional IdStatusAtual As Integer = 0) As DataTable

        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim strSQL As New StringBuilder

        strSQL.Append(" select FKDI10DI09_COD_STATUS_PROXIMO as CODIGO, Proximo.DI09_DESCRICAO + '-' + Proximo.DI09_RESUMO as DESCRICAO")
        strSQL.Append(" from DI10_SEQUENCIA_STATUS")
        strSQL.Append(" join di09_status Proximo on DI10_SEQUENCIA_STATUS.FKDI10DI09_COD_STATUS_PROXIMO = Proximo.DI09_COD_STATUS")

        strSQL.Append(" WHERE DI10_COD_SEQUENCIA_STATUS IS NOT NULL")

        If IdStatusAtual > 0 Then
            strSQL.Append(" and FKDI10DI09_COD_STATUS_ATUAL = " & IdStatusAtual)
        End If


        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function ObterUltimo() As Integer
        Dim cnn As New Conexao("StringConexaoDiarias")
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(DI10_COD_SEQUENCIA_STATUS) from DI10_SEQUENCIA_STATUS")

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
