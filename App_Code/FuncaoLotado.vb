Imports Microsoft.VisualBasic
Imports System.Data

Public Class FuncaoLotado

    Implements IDisposable

    Private RH53_ID_FUNCAO_LOTADO As Integer
    Private RH14_ID_LOTACAO_SERVIDOR As Integer
    Private RH52_ID_FUNCAO_MAPEAMENTO As integer
    Private TG06_ID_TURNO As Integer
    Private CA04_ID_USUARIO As Integer
    Private RH53_DT_LOTACAO As String
    Private RH53_DH_CADASTRO As String
    Private RH53_DH_DESATIVACAO As String
    Private CA04_ID_USUARIO_ALT As Integer
    Private RH53_DS_OBSERVACAO As String

    Public Property IdFuncaoLotado() As Integer
        Get
            Return RH53_ID_FUNCAO_LOTADO
        End Get
        Set(ByVal Value As Integer)
            RH53_ID_FUNCAO_LOTADO = Value
        End Set
    End Property
    Public Property IdLotacaoServidor() As Integer
        Get
            Return RH14_ID_LOTACAO_SERVIDOR
        End Get
        Set(ByVal Value As Integer)
            RH14_ID_LOTACAO_SERVIDOR = Value
        End Set
    End Property
    Public Property IdFuncaoMapeamento() As Integer
        Get
            Return RH52_ID_FUNCAO_MAPEAMENTO
        End Get
        Set(ByVal Value As Integer)
            RH52_ID_FUNCAO_MAPEAMENTO = Value
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
    Public Property DataLotacao() As String
        Get
            Return RH53_DT_LOTACAO
        End Get
        Set(ByVal Value As String)
            RH53_DT_LOTACAO = Value
        End Set
    End Property
    Public Property DataHoraCadastro() As String
        Get
            Return RH53_DH_CADASTRO
        End Get
        Set(ByVal Value As String)
            RH53_DH_CADASTRO = Value
        End Set
    End Property
    Public Property DataHoraDesativacao() As String
        Get
            Return RH53_DH_DESATIVACAO
        End Get
        Set(ByVal Value As String)
            RH53_DH_DESATIVACAO = Value
        End Set
    End Property
    Public Property IdUsuarioAlteracao() As Integer
        Get
            Return CA04_ID_USUARIO_ALT
        End Get
        Set(ByVal Value As Integer)
            CA04_ID_USUARIO_ALT = Value
        End Set
    End Property
    Public Property Observavao() As String
        Get
            Return RH53_DS_OBSERVACAO
        End Get
        Set(ByVal Value As String)
            RH53_DS_OBSERVACAO = Value
        End Set
    End Property
    Public Property IdTurno() As Integer
        Get
            Return TG06_ID_TURNO
        End Get
        Set(value As Integer)
            TG06_ID_TURNO = value
        End Set
    End Property

    Public Sub New(Optional ByVal IdFuncaoLotado As Integer = 0, Optional idLotacaoServidor As Integer = 0)
        If IdFuncaoLotado > 0 Or idLotacaoServidor > 0 Then
            Obter(IdFuncaoLotado, idLotacaoServidor)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH53_FUNCAO_LOTADO")
        strSQL.Append(" where RH53_ID_FUNCAO_LOTADO = " & IdFuncaoLotado)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH14_ID_LOTACAO_SERVIDOR") = ProBanco(RH14_ID_LOTACAO_SERVIDOR, eTipoValor.CHAVE)
        dr("RH52_ID_FUNCAO_MAPEAMENTO") = ProBanco(RH52_ID_FUNCAO_MAPEAMENTO, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("RH53_DT_LOTACAO") = ProBanco(RH53_DT_LOTACAO, eTipoValor.DATA)
        dr("RH53_DH_CADASTRO") = ProBanco(RH53_DH_CADASTRO, eTipoValor.DATA)
        dr("RH53_DH_DESATIVACAO") = ProBanco(RH53_DH_DESATIVACAO, eTipoValor.DATA)
        dr("CA04_ID_USUARIO_ALT") = ProBanco(CA04_ID_USUARIO_ALT, eTipoValor.CHAVE)
        dr("RH53_DS_OBSERVACAO") = ProBanco(RH53_DS_OBSERVACAO, eTipoValor.TEXTO)
        dr("TG06_ID_TURNO") = probanco(TG06_ID_TURNO, eTipoValor.CHAVE)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(Optional ByVal IdFuncaoLotado As Integer = 0, Optional IdLotacaoServidor As Integer = 0)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH53_FUNCAO_LOTADO")
        strSQL.Append(" where RH53_ID_FUNCAO_LOTADO is not null ")

        If IdFuncaoLotado > 0 Then
            strSQL.Append(" and RH53_ID_FUNCAO_LOTADO = " & IdFuncaoLotado)
        End If

        If IdLotacaoServidor > 0 Then
            strSQL.Append(" and RH14_ID_LOTACAO_SERVIDOR = " & IdLotacaoServidor)
        End If


        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH53_ID_FUNCAO_LOTADO = DoBanco(dr("RH53_ID_FUNCAO_LOTADO"), eTipoValor.CHAVE)
            RH14_ID_LOTACAO_SERVIDOR = DoBanco(dr("RH14_ID_LOTACAO_SERVIDOR"), eTipoValor.CHAVE)
            RH52_ID_FUNCAO_MAPEAMENTO = DoBanco(dr("RH52_ID_FUNCAO_MAPEAMENTO"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            RH53_DT_LOTACAO = DoBanco(dr("RH53_DT_LOTACAO"), eTipoValor.DATA)
            RH53_DH_CADASTRO = DoBanco(dr("RH53_DH_CADASTRO"), eTipoValor.DATA)
            RH53_DH_DESATIVACAO = DoBanco(dr("RH53_DH_DESATIVACAO"), eTipoValor.DATA)
            CA04_ID_USUARIO_ALT = DoBanco(dr("CA04_ID_USUARIO_ALT"), eTipoValor.CHAVE)
            RH53_DS_OBSERVACAO = DoBanco(dr("RH53_DS_OBSERVACAO"), eTipoValor.TEXTO)
            TG06_ID_TURNO = DoBanco(dr("TG06_ID_TURNO"), eTipoValor.CHAVE)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional IdFuncaoLotado As Integer = 0, Optional IdLotacaoServidor As String = "", Optional IdFuncaoMapeamento As Integer = 0, Optional IdUsuario As Integer = 0, Optional DataLotacao As String = "", Optional DataHoraCadastro As String = "", Optional DataHoraDesativacao As String = "", Optional IdUsuarioAlteracao As Integer = 0, Optional Observavao As String = "", Optional ByVal Turno As Integer = 0, Optional ByVal Lotacao As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH53_FUNCAO_LOTADO RH53")
        strSQL.Append(" left join RH14_LOTACAO_SERVIDOR RH14 on RH14.RH14_ID_LOTACAO_SERVIDOR = RH53.RH14_ID_LOTACAO_SERVIDOR ")
        strSQL.Append(" left join rh88_PERIODO RH88 on RH14.RH88_ID_PERIODO = RH88.RH88_ID_PERIODO  ")
        strSQL.Append(" where RH53_ID_FUNCAO_LOTADO is not null AND RH14_DT_DESLIGAMENTO is null  AND RH88_NM_PERIODO = YEAR(GETDATE()) ")

        If IdFuncaoLotado > 0 Then
            strSQL.Append(" and RH53_ID_FUNCAO_LOTADO = " & IdFuncaoLotado)
        End If

        If IsNumeric(IdLotacaoServidor.Replace(".", "")) Then
            strSQL.Append(" and RH14.RH14_ID_LOTACAO_SERVIDOR = " & IdLotacaoServidor.Replace(".", "").Replace(",", "."))
        End If


        If IdFuncaoMapeamento > 0 Then
            strSQL.Append(" and RH52_ID_FUNCAO_MAPEAMENTO = " & IdFuncaoMapeamento)
        End If

        If IdUsuario > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & IdUsuario)
        End If

        If Turno > 0 Then
            strSQL.Append(" and TG06_ID_TURNO = " & Turno)
        End If

        If IsDate(DataLotacao) Then
            strSQL.Append(" and RH53_DT_LOTACAO = Convert(DateTime, '" & DataLotacao & "', 103)")
        End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" and RH53_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
        End If

        If IsDate(DataHoraDesativacao) Then
            strSQL.Append(" and RH53_DH_DESATIVACAO = Convert(DateTime, '" & DataHoraDesativacao & "', 103)")
        End If

        If IdUsuarioAlteracao > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO_ALT = " & IdUsuarioAlteracao)
        End If

        If Observavao <> "" Then
            strSQL.Append(" and upper(RH53_DS_OBSERVACAO) like '%" & Observavao.ToUpper & "%'")
        End If

        If Lotacao > 0 Then
            strSQL.Append(" and RH14.RH36_ID_LOTACAO = " & Lotacao)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH53_ID_FUNCAO_LOTADO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select RH53_ID_FUNCAO_LOTADO as CODIGO, RH14_ID_LOTACAO_SERVIDOR as DESCRICAO")
        strSQL.Append(" from RH53_FUNCAO_LOTADO")
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

        strSQL.Append(" select max(RH53_ID_FUNCAO_LOTADO) from RH53_FUNCAO_LOTADO")

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
    Public Function Excluir(ByVal IdFuncaoLotado As String) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from RH53_FUNCAO_LOTADO")
        strSQL.Append(" where RH53_ID_FUNCAO_LOTADO = " & IdFuncaoLotado)

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
'*                                 06/05/2019                                 *
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

