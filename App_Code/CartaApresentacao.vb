Imports Microsoft.VisualBasic
Imports System.Data

Public Class CartaApresentacao
    Implements IDisposable

    Private RH54_ID_CARTA_APRESENTACAO As Integer
    Private RH36_ID_LOTACAO As Integer
    Private RH02_ID_SERVIDOR As Integer
    Private RH55_ID_SITUACAO_CARTA As Integer
    Private TG62_ID_POLO As Integer
    Private CA04_ID_USUARIO As Integer
    Private CA04_ID_USUARIO_RECEBIMENTO As Integer
    Private RH54_DT_APRESENTACAO As String
    Private RH54_IN_CET As Integer
    Private RH54_DH_CADASTRO As String
    Private RH54_DS_MOTIVO As String
    Private RH55_DH_SITUACAO_CARTA As String
    Private RH06_ID_FUNCAO As Integer

    Public Property Id() As Integer
        Get
            Return RH54_ID_CARTA_APRESENTACAO
        End Get
        Set(ByVal Value As Integer)
            RH54_ID_CARTA_APRESENTACAO = Value
        End Set
    End Property
    Public Property Lotacao() As Integer
        Get
            Return RH36_ID_LOTACAO
        End Get
        Set(ByVal Value As Integer)
            RH36_ID_LOTACAO = Value
        End Set
    End Property
    Public Property Servidor() As Integer
        Get
            Return RH02_ID_SERVIDOR
        End Get
        Set(ByVal Value As Integer)
            RH02_ID_SERVIDOR = Value
        End Set
    End Property
    Public Property SituacaoCarta() As Integer
        Get
            Return RH55_ID_SITUACAO_CARTA
        End Get
        Set(ByVal Value As Integer)
            RH55_ID_SITUACAO_CARTA = Value
        End Set
    End Property
    Public Property Polo() As Integer
        Get
            Return TG62_ID_POLO
        End Get
        Set(ByVal Value As Integer)
            TG62_ID_POLO = Value
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
    Public Property UsuarioRecebimentoId() As Integer
        Get
            Return CA04_ID_USUARIO_RECEBIMENTO
        End Get
        Set(ByVal Value As Integer)
            CA04_ID_USUARIO_RECEBIMENTO = Value
        End Set
    End Property
    Public Property DataApresentacao() As String
        Get
            Return RH54_DT_APRESENTACAO
        End Get
        Set(ByVal Value As String)
            RH54_DT_APRESENTACAO = Value
        End Set
    End Property
    Public Property Cet() As Integer
        Get
            Return RH54_IN_CET
        End Get
        Set(ByVal Value As Integer)
            RH54_IN_CET = Value
        End Set
    End Property
    Public Property DataHoraCadastro() As String
        Get
            Return RH54_DH_CADASTRO
        End Get
        Set(ByVal Value As String)
            RH54_DH_CADASTRO = Value
        End Set
    End Property
    Public Property DescricaoMotivo() As String
        Get
            Return RH54_DS_MOTIVO
        End Get
        Set(ByVal Value As String)
            RH54_DS_MOTIVO = Value
        End Set
    End Property

    Public Property DataHoraSituacaoCarta() As String
        Get
            Return RH55_DH_SITUACAO_CARTA
        End Get
        Set(value As String)
            RH55_DH_SITUACAO_CARTA = value
        End Set
    End Property
    Public Property  FuncaoId() As Integer  
    get
        Return RH06_ID_FUNCAO 
    End Get
        Set(value As Integer)
           RH06_ID_FUNCAO = value
        End Set
    End Property

    Public Sub New(Optional ByVal Id As Integer = 0)
        If Id > 0 Then
            Obter(Id)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select *,convert(varchar,DATEADD(DAY, 1,RH55_DH_SITUACAO_CARTA),103) as DATA_LIMITE ")
        strSQL.Append(" from RH54_CARTA_APRESENTACAO")
        strSQL.Append(" where RH54_ID_CARTA_APRESENTACAO = " & Id)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH36_ID_LOTACAO") = ProBanco(RH36_ID_LOTACAO, eTipoValor.CHAVE)
        dr("RH02_ID_SERVIDOR") = ProBanco(RH02_ID_SERVIDOR, eTipoValor.CHAVE)
        dr("RH55_ID_SITUACAO_CARTA") = ProBanco(RH55_ID_SITUACAO_CARTA, eTipoValor.CHAVE)
        dr("TG62_ID_POLO") = ProBanco(TG62_ID_POLO, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO_RECEBIMENTO") = ProBanco(CA04_ID_USUARIO_RECEBIMENTO, eTipoValor.CHAVE)
        dr("RH54_DT_APRESENTACAO") = ProBanco(RH54_DT_APRESENTACAO, eTipoValor.DATA)
        dr("RH54_IN_CET") = ProBanco(RH54_IN_CET, eTipoValor.CHAVE)
        dr("RH54_DH_CADASTRO") = ProBanco(RH54_DH_CADASTRO, eTipoValor.DATA_COMPLETA)
        dr("RH54_DS_MOTIVO") = ProBanco(RH54_DS_MOTIVO, eTipoValor.TEXTO)
        dr("RH55_DH_SITUACAO_CARTA") = ProBanco(RH55_DH_SITUACAO_CARTA, eTipoValor.DATA_COMPLETA)
        dr("RH06_ID_FUNCAO") = probanco(RH06_ID_FUNCAO, eTipoValor.CHAVE)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal Id As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH54_CARTA_APRESENTACAO")
        strSQL.Append(" where RH54_ID_CARTA_APRESENTACAO = " & Id)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH54_ID_CARTA_APRESENTACAO = DoBanco(dr("RH54_ID_CARTA_APRESENTACAO"), eTipoValor.CHAVE)
            RH36_ID_LOTACAO = DoBanco(dr("RH36_ID_LOTACAO"), eTipoValor.CHAVE)
            RH02_ID_SERVIDOR = DoBanco(dr("RH02_ID_SERVIDOR"), eTipoValor.CHAVE)
            RH55_ID_SITUACAO_CARTA = DoBanco(dr("RH55_ID_SITUACAO_CARTA"), eTipoValor.CHAVE)
            TG62_ID_POLO = DoBanco(dr("TG62_ID_POLO"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            CA04_ID_USUARIO_RECEBIMENTO = DoBanco(dr("CA04_ID_USUARIO_RECEBIMENTO"), eTipoValor.CHAVE)
            RH54_DT_APRESENTACAO = DoBanco(dr("RH54_DT_APRESENTACAO"), eTipoValor.DATA)
            RH54_IN_CET = DoBanco(dr("RH54_IN_CET"), eTipoValor.CHAVE)
            RH54_DH_CADASTRO = DoBanco(dr("RH54_DH_CADASTRO"), eTipoValor.DATA_COMPLETA)
            RH54_DS_MOTIVO = DoBanco(dr("RH54_DS_MOTIVO"), eTipoValor.TEXTO)
            RH55_DH_SITUACAO_CARTA = DoBanco(dr("RH55_DH_SITUACAO_CARTA"), eTipoValor.DATA_COMPLETA)
            RH06_ID_FUNCAO = DoBanco(dr("RH06_ID_FUNCAO"),eTipoValor.CHAVE)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Id As Integer = 0, Optional Lotacao As Integer = 0, Optional Servidor As Integer = 0, Optional SituacaoCarta As Integer = 0, Optional Polo As Integer = 0, Optional UsuarioId As Integer = 0, Optional UsuarioRecebimentoId As Integer = 0, Optional DataApresentacao As String = "", Optional Cet As Integer = 0, Optional DataHoraCadastro As Integer = 0, Optional DescricaoMotivo As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH54_CARTA_APRESENTACAO")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where RH54_ID_CARTA_APRESENTACAO is not null")

        If Id > 0 Then
            strSQL.Append(" and RH54_ID_CARTA_APRESENTACAO = " & Id)
        End If

        If Lotacao > 0 Then
            strSQL.Append(" and RH36_ID_LOTACAO = " & Lotacao)
        End If

        If Servidor > 0 Then
            strSQL.Append(" and RH02_ID_SERVIDOR = " & Servidor)
        End If

        If SituacaoCarta > 0 Then
            strSQL.Append(" and RH55_ID_SITUACAO_CARTA = " & SituacaoCarta)
        End If

        If Polo > 0 Then
            strSQL.Append(" and TG62_ID_POLO = " & Polo)
        End If

        If UsuarioId > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & UsuarioId)
        End If

        If UsuarioRecebimentoId > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO_RECEBIMENTO = " & UsuarioRecebimentoId)
        End If

        If DataApresentacao <> "" Then
            strSQL.Append(" and upper(RH54_DT_APRESENTACAO) like '%" & DataApresentacao.ToUpper & "%'")
        End If

        If Cet > 0 Then
            strSQL.Append(" and RH54_IN_CET = " & Cet)
        End If

        If DataHoraCadastro > 0 Then
            strSQL.Append(" and RH54_DH_CADASTRO = " & DataHoraCadastro)
        End If

        If DescricaoMotivo <> "" Then
            strSQL.Append(" and upper(RH54_DS_MOTIVO) like '%" & DescricaoMotivo.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH54_ID_CARTA_APRESENTACAO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select RH54_ID_CARTA_APRESENTACAO as CODIGO, RH36_ID_LOTACAO as DESCRICAO")
        strSQL.Append(" from RH54_CARTA_APRESENTACAO")
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

        strSQL.Append(" select max(RH54_ID_CARTA_APRESENTACAO) from RH54_CARTA_APRESENTACAO")

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
    Public Function Excluir(ByVal Id As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from RH54_CARTA_APRESENTACAO")
        strSQL.Append(" where RH54_ID_CARTA_APRESENTACAO = " & Id)

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
'*                                 14/01/2019                                 *
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

