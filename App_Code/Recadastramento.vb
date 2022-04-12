Imports Microsoft.VisualBasic
Imports System.Data

Public Class Recadastramento

    Implements IDisposable
    Private RH86_ID_RECADASTRAMENTO As Integer
    Private RH01_ID_PESSOA As Integer
    Private CA04_ID_USUARIO As Integer
    Private RH86_DH_CADASTRO As String
    Private RH86_NR_ANO_RECADASTRAMENTO As String

    Public Property RecadastramentoId As Integer
        Get
            Return RH86_ID_RECADASTRAMENTO
        End Get
        Set(value As Integer)
            RH86_ID_RECADASTRAMENTO = value
        End Set
    End Property

    Public Property PessoaID As Integer
        Get
            Return RH01_ID_PESSOA
        End Get
        Set(value As Integer)
            RH01_ID_PESSOA = value
        End Set
    End Property

    Public Property UsuarioId As Integer
        Get
            Return CA04_ID_USUARIO
        End Get
        Set(value As Integer)
            CA04_ID_USUARIO = value
        End Set
    End Property

    Public Property DataHoraCadastro As String
        Get
            Return RH86_DH_CADASTRO
        End Get
        Set(value As String)
            RH86_DH_CADASTRO = value
        End Set
    End Property

    Public Property AnoRecadastramento As String
        Get
            Return RH86_NR_ANO_RECADASTRAMENTO
        End Get
        Set(value As String)
            RH86_NR_ANO_RECADASTRAMENTO = value
        End Set
    End Property

    Public Sub New(Optional ByVal RecadastramentoId As Integer = 0, Optional ByVal Ano As Integer = 0)
        If RecadastramentoId > 0 And Ano > 0 Then
            Obter(RecadastramentoId, Ano)
        End If

    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH86_RECADASTRAMENTO")
        strSQL.Append(" where RH86_ID_RECADASTRAMENTO = " & RecadastramentoId)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH01_ID_PESSOA") = ProBanco(RH01_ID_PESSOA, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("RH86_DH_CADASTRO") = ProBanco(RH86_DH_CADASTRO, eTipoValor.DATA_COMPLETA)
        dr("RH86_NR_ANO_RECADASTRAMENTO") = ProBanco(RH86_NR_ANO_RECADASTRAMENTO, eTipoValor.NUMERO_INTEIRO)
        'dr("RH16_IN_COMISSIONADO") = ProBanco(RH16_IN_COMISSIONADO, eTipoValor.CHAVE)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub
    Public Sub Obter(ByVal PessoaID As String, ByVal Ano As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH86_RECADASTRAMENTO")
        strSQL.Append(" where RH01_ID_PESSOA = " & PessoaID)
        strSQL.Append(" and year(RH86_DH_CADASTRO)  =" & Ano)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH86_ID_RECADASTRAMENTO = DoBanco(dr("RH86_ID_RECADASTRAMENTO"), eTipoValor.CHAVE)
            RH01_ID_PESSOA = DoBanco(dr("RH01_ID_PESSOA"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            RH86_DH_CADASTRO = DoBanco(dr("RH86_DH_CADASTRO"), eTipoValor.DATA_COMPLETA)
            RH86_NR_ANO_RECADASTRAMENTO = DoBanco(dr("RH86_NR_ANO_RECADASTRAMENTO"), eTipoValor.NUMERO_INTEIRO)

        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function ObterUltimo() As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(RH86_ID_RECADASTRAMENTO) from RH86_RECADASTRAMENTO")

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

    Public Function Excluir(ByVal RecadastramentoId As String) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from RH86_RECADASTRAMENTO")
        strSQL.Append(" where RH86_ID_RECADASTRAMENTO = " & RecadastramentoId)

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
