Imports Microsoft.VisualBasic
Imports System.Data

Public Class EncaminhamentoCarta
    Implements IDisposable

    Private RH75_ID_ENCAMINHAMENTO_CARTA As Integer
    Private RH54_ID_CARTA_APRESENTACAO As Integer
    Private RH36_ID_LOTACAO As Integer
    Private CA04_ID_USUARIO As Integer
    Private RH75_DH_CADASTRO As String
    Private RH75_IN_ATIVO As String

    Public Property EncaminhamentoId() As Integer
        Get
            Return RH75_ID_ENCAMINHAMENTO_CARTA
        End Get
        Set(ByVal Value As Integer)
            RH75_ID_ENCAMINHAMENTO_CARTA = Value
        End Set
    End Property
    Public Property CartaApresentacaoId() As Integer
        Get
            Return RH54_ID_CARTA_APRESENTACAO
        End Get
        Set(ByVal Value As Integer)
            RH54_ID_CARTA_APRESENTACAO = Value
        End Set
    End Property
    Public Property LotacaoId() As Integer
        Get
            Return RH36_ID_LOTACAO
        End Get
        Set(ByVal Value As Integer)
            RH36_ID_LOTACAO = Value
        End Set
    End Property
    Public Property UsuarioId() As Integer
        Get
            Return CA04_ID_USUARIO
        End Get
        Set(ByVal Value As Integer)
            CA04_ID_USUARIO = Value
        End Set
    End Property
    Public Property DataHoraCadastro() As String
        Get
            Return RH75_DH_CADASTRO
        End Get
        Set(ByVal Value As String)
            RH75_DH_CADASTRO = Value
        End Set
    End Property
    Public Property Ativo() As String
        Get
            Return RH75_IN_ATIVO
        End Get
        Set(ByVal Value As String)
            RH75_IN_ATIVO = Value
        End Set
    End Property

    Public Sub New(Optional ByVal EncaminhamentoId As Integer = 0, Optional ByVal CartaApresentacao As Integer = 0)
        If EncaminhamentoId > 0 Or CartaApresentacao > 0 Then
            Obter(EncaminhamentoId, CartaApresentacao)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH75_ENCAMINHAMENTO_CARTA")
        strSQL.Append(" where RH75_ID_ENCAMINHAMENTO_CARTA = " & EncaminhamentoId)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH54_ID_CARTA_APRESENTACAO") = ProBanco(RH54_ID_CARTA_APRESENTACAO, eTipoValor.CHAVE)
        dr("RH36_ID_LOTACAO") = ProBanco(RH36_ID_LOTACAO, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("RH75_DH_CADASTRO") = ProBanco(RH75_DH_CADASTRO, eTipoValor.DATA)
        dr("RH75_IN_ATIVO") = ProBanco(RH75_IN_ATIVO, eTipoValor.BOOLEANO)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(Optional ByVal EncaminhamentoId As Integer = 0, Optional ByVal CartaApresentacao As Integer = 0)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH75_ENCAMINHAMENTO_CARTA")
        strSQL.Append(" where RH75_ID_ENCAMINHAMENTO_CARTA > 0")

        If EncaminhamentoId > 0 Then
            strSQL.Append(" and  RH75_ID_ENCAMINHAMENTO_CARTA = " & EncaminhamentoId)
        End If

        If CartaApresentacao > 0 Then
            strSQL.Append(" and RH54_ID_CARTA_APRESENTACAO = " & CartaApresentacao)
        End If

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH75_ID_ENCAMINHAMENTO_CARTA = DoBanco(dr("RH75_ID_ENCAMINHAMENTO_CARTA"), eTipoValor.CHAVE)
            RH54_ID_CARTA_APRESENTACAO = DoBanco(dr("RH54_ID_CARTA_APRESENTACAO"), eTipoValor.CHAVE)
            RH36_ID_LOTACAO = DoBanco(dr("RH36_ID_LOTACAO"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            RH75_DH_CADASTRO = DoBanco(dr("RH75_DH_CADASTRO"), eTipoValor.DATA)
            RH75_IN_ATIVO = DoBanco(dr("RH75_IN_ATIVO"), eTipoValor.BOOLEANO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional EncaminhamentoId As Integer = 0, Optional CartaApresentacaoId As Integer = 0, Optional LotacaoId As String = "", Optional UsuarioId As Integer = 0, Optional DataHoraCadastro As String = "", Optional Ativo As String = "", Optional SituacaoId As Integer = 0, Optional PessoaId As Integer = 0, Optional ByVal SituacaoCarta As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select *,convert(varchar,DATEADD(DAY, 3,CartaApresentacao.RH55_DH_SITUACAO_CARTA),103) as DATA_LIMITE, LotacaoEcaminhamento.RH36_SG_LOTACAO as LotacaoEcaminhamento, Lotacao.RH36_NM_LOTACAO as Lotacao")
        strSQL.Append(" from RH75_ENCAMINHAMENTO_CARTA EncaminhamentoCarta")
        strSQL.Append(" inner join RH54_CARTA_APRESENTACAO CartaApresentacao on CartaApresentacao.RH54_ID_CARTA_APRESENTACAO = EncaminhamentoCarta.RH54_ID_CARTA_APRESENTACAO ")
        strSQL.Append(" inner join RH02_SERVIDOR Servidor on Servidor.RH02_ID_SERVIDOR = CartaApresentacao.RH02_ID_SERVIDOR ")
        strSQL.Append(" inner join RH36_LOTACAO LotacaoEcaminhamento on LotacaoEcaminhamento.RH36_ID_LOTACAO = EncaminhamentoCarta.RH36_ID_LOTACAO  ")
        strSQL.Append(" left join RH06_FUNCAO FUNCAO on FUNCAO.RH06_ID_FUNCAO = CartaApresentacao.RH06_ID_FUNCAO ")
        strSQL.Append(" Left Join RH36_LOTACAO Lotacao on Lotacao.RH36_ID_LOTACAO = CartaApresentacao.RH36_ID_LOTACAO  ")

        strSQL.Append(" where RH75_ID_ENCAMINHAMENTO_CARTA is not null")

        If EncaminhamentoId > 0 Then
            strSQL.Append(" and RH75_ID_ENCAMINHAMENTO_CARTA = " & EncaminhamentoId)
        End If

        If CartaApresentacaoId > 0 Then
            strSQL.Append(" and RH54_ID_CARTA_APRESENTACAO = " & CartaApresentacaoId)
        End If

        If IsNumeric(LotacaoId.Replace(".", "")) Then
            strSQL.Append(" and RH36_ID_LOTACAO = " & LotacaoId.Replace(".", "").Replace(",", "."))
        End If

        If UsuarioId > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & UsuarioId)
        End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" and RH75_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
        End If

        If Ativo <> "" Then
            strSQL.Append(" and upper(RH75_IN_ATIVO) like '%" & Ativo.ToUpper & "%'")
        End If

        If SituacaoId > 0 Then
            strSQL.Append(" and Servidor.RH07_ID_SITUACAO_SERVIDOR in (1,11,10)")
        End If

        If PessoaId > 0 Then
            strSQL.Append(" and Servidor.RH01_ID_Pessoa = " & PessoaId)
        End If

        If SituacaoCarta > 0 then 
            strSQL.Append(" and CartaApresentacao.RH55_ID_SITUACAO_CARTA =" & SituacaoCarta)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH75_ID_ENCAMINHAMENTO_CARTA", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select RH75_ID_ENCAMINHAMENTO_CARTA as CODIGO, RH36_ID_LOTACAO as DESCRICAO")
        strSQL.Append(" from RH75_ENCAMINHAMENTO_CARTA")
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

        strSQL.Append(" select max(RH75_ID_ENCAMINHAMENTO_CARTA) from RH75_ENCAMINHAMENTO_CARTA")

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
    Public Function Excluir(ByVal EncaminhamentoId As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from RH75_ENCAMINHAMENTO_CARTA")
        strSQL.Append(" where RH75_ID_ENCAMINHAMENTO_CARTA = " & EncaminhamentoId)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

'******************************************************************************
'*                                 18/01/2019                                 *
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

