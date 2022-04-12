Imports System.Data
Imports Microsoft.VisualBasic

Public Class PeriodoFerias

    Implements IDisposable

    Private RH87_ID_PERIODO_FERIAS As Integer
    Private RH87_NR_ANO_REFERENCIA As Integer
    Private RH87_NR_ANO_FRUICAO As Integer
    Private RH87_DT_INICIO_LANCAMENTO As String
    Private RH87_DT_TERMINO_LANCAMENTO As String
    Private RH87_DT_LIMITE_LANCAMENTO As String
    Private RH87_ST_PERIODO_FERIAS As String
    Private RH87_DH_ST_PERIODO_FERIAS As String
    Private CA04_ID_USUARIO As Integer

    Public Property IdPeriodoFerias() As Integer
        Get
            Return RH87_ID_PERIODO_FERIAS
        End Get
        Set(value As Integer)
            RH87_ID_PERIODO_FERIAS = value
        End Set
    End Property

    Public Property AnoReferencia() As Integer
        Get
            Return RH87_NR_ANO_REFERENCIA
        End Get
        Set(value As Integer)
            RH87_NR_ANO_REFERENCIA = value
        End Set
    End Property

    Public Property AnoFruicao() As Integer
        Get
            Return RH87_NR_ANO_FRUICAO
        End Get
        Set(value As Integer)
            RH87_NR_ANO_FRUICAO = value
        End Set
    End Property

    Public Property DataInicioLancamento As String
        Get
            Return RH87_DT_INICIO_LANCAMENTO
        End Get
        Set(value As String)
            RH87_DT_INICIO_LANCAMENTO = value
        End Set
    End Property

    Public Property DataFimLancamento As String
        Get
            Return RH87_DT_TERMINO_LANCAMENTO
        End Get
        Set(value As String)
            RH87_DT_TERMINO_LANCAMENTO = value
        End Set
    End Property

    Public Property DataLimiteLancamento As String
        Get
            Return RH87_DT_LIMITE_LANCAMENTO
        End Get
        Set(value As String)
            RH87_DT_LIMITE_LANCAMENTO = value
        End Set
    End Property

    Public Property SituacaoPeriodoFerias As String
        Get
            Return RH87_ST_PERIODO_FERIAS
        End Get
        Set(value As String)
            RH87_ST_PERIODO_FERIAS = value
        End Set
    End Property

    Public Property DataHoraSituacaoPeriodoFerias As String
        Get
            Return RH87_DH_ST_PERIODO_FERIAS
        End Get
        Set(value As String)
            RH87_DH_ST_PERIODO_FERIAS = value
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

    Public Sub New(Optional ByVal IdPeriodoFerias As Integer = 0)
        If IdPeriodoFerias > 0 Then
            Obter(IdPeriodoFerias)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH87_PERIODO_FERIAS")
        strSQL.Append(" where RH87_ID_PERIODO_FERIAS = " & IdPeriodoFerias)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH87_NR_ANO_REFERENCIA") = ProBanco(RH87_NR_ANO_REFERENCIA, eTipoValor.CHAVE)
        dr("RH87_NR_ANO_FRUICAO") = ProBanco(RH87_NR_ANO_FRUICAO, eTipoValor.CHAVE)
        dr("RH87_DT_INICIO_LANCAMENTO") = ProBanco(RH87_DT_INICIO_LANCAMENTO, eTipoValor.DATA)
        dr("RH87_DT_TERMINO_LANCAMENTO") = ProBanco(RH87_DT_TERMINO_LANCAMENTO, eTipoValor.DATA)
        dr("RH87_DT_LIMITE_LANCAMENTO") = ProBanco(RH87_DT_LIMITE_LANCAMENTO, eTipoValor.DATA_COMPLETA)
        dr("RH87_ST_PERIODO_FERIAS") = ProBanco(RH87_ST_PERIODO_FERIAS, eTipoValor.TEXTO)
        dr("RH87_DH_ST_PERIODO_FERIAS") = ProBanco(RH87_DH_ST_PERIODO_FERIAS, eTipoValor.DATA_COMPLETA)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)


        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal IdPeriodoFerias As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH87_PERIODO_FERIAS")
        strSQL.Append(" where RH87_ID_PERIODO_FERIAS = " & IdPeriodoFerias)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH87_ID_PERIODO_FERIAS = DoBanco(dr("RH87_ID_PERIODO_FERIAS"), eTipoValor.CHAVE)
            RH87_NR_ANO_REFERENCIA = DoBanco(dr("RH87_NR_ANO_REFERENCIA"), eTipoValor.CHAVE)
            RH87_NR_ANO_FRUICAO = DoBanco(dr("RH87_NR_ANO_FRUICAO"), eTipoValor.CHAVE)
            RH87_DT_INICIO_LANCAMENTO = DoBanco(dr("RH87_DT_INICIO_LANCAMENTO"), eTipoValor.DATA)
            RH87_DT_TERMINO_LANCAMENTO = DoBanco(dr("RH87_DT_TERMINO_LANCAMENTO"), eTipoValor.DATA)
            RH87_DT_LIMITE_LANCAMENTO = DoBanco(dr("RH87_DT_LIMITE_LANCAMENTO"), eTipoValor.DATA)
            RH87_ST_PERIODO_FERIAS = DoBanco(dr("RH87_ST_PERIODO_FERIAS"), eTipoValor.TEXTO)
            RH87_DH_ST_PERIODO_FERIAS = DoBanco(dr("RH87_DH_ST_PERIODO_FERIAS"), eTipoValor.DATA_COMPLETA)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)

        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub
    Public Function PesquisarPeriodoAtivo() As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select RH87_ID_PERIODO_FERIAS,RH87_NR_ANO_REFERENCIA,RH87_NR_ANO_FRUICAO  ")
        strSQL.Append(" ,RH87_DT_INICIO_LANCAMENTO,RH87_DT_TERMINO_LANCAMENTO,RH87_DT_LIMITE_LANCAMENTO,RH87_ST_PERIODO_FERIAS  ")
        strSQL.Append(" from RH87_PERIODO_FERIAS  ")
        strSQL.Append(" where RH87_ST_PERIODO_FERIAS = 'A'  ")
        strSQL.Append(" and RH87_DT_LIMITE_LANCAMENTO >=   convert(date,getdate())    ")

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional PeriodoId As Integer = 0, Optional ByVal Fruicao As Integer = 0, Optional ByVal Referencia As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select *, convert(varchar,RH87_NR_ANO_REFERENCIA)+ ' / '+ convert(varchar,RH87_NR_ANO_FRUICAO) as ReferenciaFruicao ")
        strSQL.Append(" from RH87_PERIODO_FERIAS")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where RH87_ID_PERIODO_FERIAS is not null")

        If Fruicao > 0 Then
            strSQL.Append(" where RH87_ID_PERIODO_FERIAS =" & Fruicao)
        End If

        If Referencia > 0 Then
            strSQL.Append(" where RH87_NR_ANO_REFERENCIA =" & Referencia)
        End If

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function ObterUltimo() As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(RH87_ID_PERIODO_FERIAS) from RH87_PERIODO_FERIAS")

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

    Public Function Excluir(ByVal PeriodoFeriasId As String) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from RH87_PERIODO_FERIAS")
        strSQL.Append(" where RH87_ID_PERIODO_FERIAS = " & PeriodoFeriasId)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' Para detectar chamadas redundantes

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: descartar estado gerenciado (objetos gerenciados).
            End If

            ' TODO: liberar recursos não gerenciados (objetos não gerenciados) e substituir um Finalize() abaixo.
            ' TODO: definir campos grandes como nulos.
        End If
        disposedValue = True
    End Sub

    ' TODO: substituir Finalize() somente se Dispose(disposing As Boolean) acima tiver o código para liberar recursos não gerenciados.
    'Protected Overrides Sub Finalize()
    '    ' Não altere este código. Coloque o código de limpeza em Dispose(disposing As Boolean) acima.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' Código adicionado pelo Visual Basic para implementar corretamente o padrão descartável.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Não altere este código. Coloque o código de limpeza em Dispose(disposing As Boolean) acima.
        Dispose(True)
        ' TODO: remover marca de comentário da linha a seguir se Finalize() for substituído acima.
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
