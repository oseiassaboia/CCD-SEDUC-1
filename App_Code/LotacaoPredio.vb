Imports Microsoft.VisualBasic
Imports System.Data

Public Class LotacaoPredio
    Implements IDisposable
    Private RH70_ID_LOTACAO_PREDIO As Integer
    Private RH36_ID_LOTACAO As Integer
    Private TG59_ID_PREDIO As Integer
    Private CA04_ID_USUARIO As Integer
    Private RH70_DH_CADASTRO As String
    Private RH70_DH_DESATIVACAO As String

    Public Property IdLotacaoPredio() As Integer
        Get
            Return RH70_ID_LOTACAO_PREDIO
        End Get
        Set(ByVal Value As Integer)
            RH70_ID_LOTACAO_PREDIO = Value
        End Set
    End Property
    Public Property IdLotacao() As Integer
        Get
            Return RH36_ID_LOTACAO
        End Get
        Set(ByVal Value As Integer)
            RH36_ID_LOTACAO = Value
        End Set
    End Property
    Public Property idPredio() As Integer
        Get
            Return TG59_ID_PREDIO
        End Get
        Set(ByVal Value As Integer)
            TG59_ID_PREDIO = Value
        End Set
    End Property
    Public Property IdUsuario() As Integer
        Get
            Return CA04_ID_USUARIO
        End Get
        Set(ByVal Value As Integer)
            CA04_ID_USUARIO = Value
        End Set
    End Property
    Public Property DataHoraCadastro() As String
        Get
            Return RH70_DH_CADASTRO
        End Get
        Set(ByVal Value As String)
            RH70_DH_CADASTRO = Value
        End Set
    End Property
    Public Property DataHoraDesativacao() As String
        Get
            Return RH70_DH_DESATIVACAO
        End Get
        Set(ByVal Value As String)
            RH70_DH_DESATIVACAO = Value
        End Set
    End Property

    Public Sub New(Optional ByVal IdLotacaoPredio As Integer = 0)
        If IdLotacaoPredio > 0 Then
            Obter(IdLotacaoPredio)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH70_LOTACAO_PREDIO")
        strSQL.Append(" where RH70_ID_LOTACAO_PREDIO = " & IdLotacaoPredio)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH36_ID_LOTACAO") = ProBanco(RH36_ID_LOTACAO, eTipoValor.CHAVE)
        dr("TG59_ID_PREDIO") = ProBanco(TG59_ID_PREDIO, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("RH70_DH_CADASTRO") = ProBanco(RH70_DH_CADASTRO, eTipoValor.DATA)
        dr("RH70_DH_DESATIVACAO") = ProBanco(RH70_DH_DESATIVACAO, eTipoValor.DATA)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal IdLotacaoPredio As String)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH70_LOTACAO_PREDIO")
        strSQL.Append(" where RH70_ID_LOTACAO_PREDIO = " & IdLotacaoPredio)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH70_ID_LOTACAO_PREDIO = DoBanco(dr("RH70_ID_LOTACAO_PREDIO"), eTipoValor.CHAVE)
            RH36_ID_LOTACAO = DoBanco(dr("RH36_ID_LOTACAO"), eTipoValor.CHAVE)
            TG59_ID_PREDIO = DoBanco(dr("TG59_ID_PREDIO"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            RH70_DH_CADASTRO = DoBanco(dr("RH70_DH_CADASTRO"), eTipoValor.DATA)
            RH70_DH_DESATIVACAO = DoBanco(dr("RH70_DH_DESATIVACAO"), eTipoValor.DATA)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional IdLotacaoPredio As Integer = 0, Optional IdLotacao As Integer = 0, Optional idPredio As Integer = 0, Optional IdUsuario As Integer = 0, Optional DataHoraCadastro As String = "", Optional DataHoraDesativacao As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH70_LOTACAO_PREDIO")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where RH70_ID_LOTACAO_PREDIO is not null")

        If IdLotacaoPredio > 0 Then
            strSQL.Append(" and RH70_ID_LOTACAO_PREDIO = " & IdLotacaoPredio)
        End If

        If IdLotacao > 0 Then
            strSQL.Append(" and RH36_ID_LOTACAO = " & IdLotacao)
        End If

        If idPredio > 0 Then
            strSQL.Append(" and TG59_ID_PREDIO = " & idPredio)
        End If

        If IdUsuario > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & IdUsuario)
        End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" and RH70_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
        End If

        If IsDate(DataHoraDesativacao) Then
            strSQL.Append(" and RH70_DH_DESATIVACAO = Convert(DateTime, '" & DataHoraDesativacao & "', 103)")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH70_ID_LOTACAO_PREDIO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select RH70_ID_LOTACAO_PREDIO as CODIGO, RH36_ID_LOTACAO as DESCRICAO")
        strSQL.Append(" from RH70_LOTACAO_PREDIO")
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

        strSQL.Append(" select max(RH70_ID_LOTACAO_PREDIO) from RH70_LOTACAO_PREDIO")

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
    Public Function Excluir(ByVal IdLotacaoPredio As String) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from RH70_LOTACAO_PREDIO")
        strSQL.Append(" where RH70_ID_LOTACAO_PREDIO = " & IdLotacaoPredio)

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

'******************************************************************************
'*                                 23/09/2019                                 *
'*                                                                            *
'*          ESTE CÓDIGO FOI GERADO PELO GERA CODIGO VERSÃO 4.0                *
'*    SUPORTE PARA ASP.NET 2.0, AJAX, SQL SERVER COM ENTERPRISE LIBRARY       *
'*                                                                            *
'*  O Gera-Codigo gera um MODELO de código Página, Interface, Classe e Css    *
'*  cabe a cada programador fazer as adaptações quando NECESSÁRIAS.           *
'*                                                                            *
'*  Esta ferramenta é TOTALMENTE GRATUITA, por favor, não remova os créditos  *
'*                                                                            *
'*  O autor não se responsabiliza por qualquer evento acontecido com o uso    *
'*  desta ferramenta ou do sistema que ela vier a gerar.                      *
'*                                                                            *
'*          Desenvolvido por Nírondes Anglada Casanovas Tavares               *
'*                  E-Mail/MSN: nirondes@hotmail.com                          *
'*                                                                            *
'******************************************************************************

