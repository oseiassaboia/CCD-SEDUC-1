Imports Microsoft.VisualBasic
Imports System.Data


Public Class HabilidadeServidor
    Implements IDisposable

    Private RH72_ID_HABILIDADE As Integer
    Private RH02_ID_SERVIDOR As Integer
    Private DE07_ID_ETAPA As Integer
    Private DE09_ID_DISCIPLINA As Integer
    Private RH72_TP_HABILIDADE As Integer
    Private CA04_ID_USUARIO As Integer
    Private CA04_ID_USUARIO_ALT As Integer
    Private RH72_DH_CADASTRO As String
    Private RH72_DH_DESATIVACAO As String

    Public Property IdHabilidade() As Integer
        Get
            Return RH72_ID_HABILIDADE
        End Get
        Set(ByVal Value As Integer)
            RH72_ID_HABILIDADE = Value
        End Set
    End Property
    Public Property IdServidor() As Integer
        Get
            Return RH02_ID_SERVIDOR
        End Get
        Set(ByVal Value As Integer)
            RH02_ID_SERVIDOR = Value
        End Set
    End Property
    Public Property IdEtapa() As Integer
        Get
            Return DE07_ID_ETAPA
        End Get
        Set(ByVal Value As Integer)
            DE07_ID_ETAPA = Value
        End Set
    End Property
    Public Property IdDisciplina() As Integer
        Get
            Return DE09_ID_DISCIPLINA
        End Get
        Set(ByVal Value As Integer)
            DE09_ID_DISCIPLINA = Value
        End Set
    End Property
    Public Property TpHabilidade() As Integer
        Get
            Return RH72_TP_HABILIDADE
        End Get
        Set(ByVal Value As Integer)
            RH72_TP_HABILIDADE = Value
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
    Public Property IdUsuarioAlteracao() As Integer
        Get
            Return CA04_ID_USUARIO_ALT
        End Get
        Set(ByVal Value As Integer)
            CA04_ID_USUARIO_ALT = Value
        End Set
    End Property
    Public Property DataHoraCadastro() As String
        Get
            Return RH72_DH_CADASTRO
        End Get
        Set(ByVal Value As String)
            RH72_DH_CADASTRO = Value
        End Set
    End Property
    Public Property DataHoraDesativacao() As String
        Get
            Return RH72_DH_DESATIVACAO
        End Get
        Set(ByVal Value As String)
            RH72_DH_DESATIVACAO = Value
        End Set
    End Property

    Public Sub New(Optional ByVal IdHabilidade As Integer = 0)
        If IdHabilidade > 0 Then
            Obter(IdHabilidade)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH72_HABILIDADE")
        strSQL.Append(" where RH72_ID_HABILIDADE = " & IdHabilidade)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH02_ID_SERVIDOR") = ProBanco(RH02_ID_SERVIDOR, eTipoValor.NUMERO_DECIMAL)

        dr("DE09_ID_DISCIPLINA") = ProBanco(DE09_ID_DISCIPLINA, eTipoValor.NUMERO_DECIMAL)
        dr("RH72_TP_HABILIDADE") = ProBanco(RH72_TP_HABILIDADE, eTipoValor.NUMERO_DECIMAL)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO_ALT") = ProBanco(CA04_ID_USUARIO_ALT, eTipoValor.CHAVE)
        dr("RH72_DH_CADASTRO") = ProBanco(RH72_DH_CADASTRO, eTipoValor.DATA)
        dr("RH72_DH_DESATIVACAO") = ProBanco(RH72_DH_DESATIVACAO, eTipoValor.DATA)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal IdHabilidade As String)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH72_HABILIDADE")
        strSQL.Append(" where RH72_ID_HABILIDADE = " & IdHabilidade)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH72_ID_HABILIDADE = DoBanco(dr("RH72_ID_HABILIDADE"), eTipoValor.CHAVE)
            RH02_ID_SERVIDOR = DoBanco(dr("RH02_ID_SERVIDOR"), eTipoValor.NUMERO_DECIMAL)
           
            DE09_ID_DISCIPLINA = DoBanco(dr("DE09_ID_DISCIPLINA"), eTipoValor.NUMERO_DECIMAL)
            RH72_TP_HABILIDADE = DoBanco(dr("RH72_TP_HABILIDADE"), eTipoValor.NUMERO_DECIMAL)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            CA04_ID_USUARIO_ALT = DoBanco(dr("CA04_ID_USUARIO_ALT"), eTipoValor.CHAVE)
            RH72_DH_CADASTRO = DoBanco(dr("RH72_DH_CADASTRO"), eTipoValor.DATA)
            RH72_DH_DESATIVACAO = DoBanco(dr("RH72_DH_DESATIVACAO"), eTipoValor.DATA)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional IdHabilidade As Integer = 0, Optional IdServidor As Integer = 0, Optional IdEtapa As Integer = 0, Optional IdDisciplina As Integer = 0, Optional TpHabilidade As Integer = 0, Optional IdUsuario As Integer = 0, Optional IdUsuarioAlteracao As Integer = 0, Optional DataHoraCadastro As String = "", Optional DataHoraDesativacao As String = "", Optional ByVal CodigoPessoa As Integer = 0, Optional ByVal RegistroAtivo As Boolean = True) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select *,   CASE WHEN RH72_TP_HABILIDADE = 1  THEN 'COMUM' WHEN RH72_TP_HABILIDADE = 2 THEN 'DESVIO' END AS TIPO_HABILIDADE ")
        strSQL.Append(" ,'DISCIPLINA: ' + DE09_NM_DISCIPLINA  as DESCRICAO , case when isnull(habilidade.RH72_DH_DESATIVACAO,0)   = ''  then 1 else 0 end as Desativado ")
        strSQL.Append(" from RH72_HABILIDADE habilidade")
        strSQL.Append(" left join RH02_SERVIDOR AS SERVIDOR on SERVIDOR.RH02_ID_SERVIDOR = habilidade.RH02_ID_SERVIDOR ")
        strSQL.Append(" LEFT JOIN DBDIARIO_TESTE..DE09_DISCIPLINA AS DISCIPLINA ON DISCIPLINA.DE09_ID_DISCIPLINA = HABILIDADE.DE09_ID_DISCIPLINA ")
        strSQL.Append(" where RH72_ID_HABILIDADE is not null")

        If RegistroAtivo Then
            strSQL.Append(" and RH72_DH_DESATIVACAO is null")
        End If

        If CodigoPessoa > 0 Then
            strSQL.Append(" and SERVIDOR.RH01_ID_PESSOA = " & CodigoPessoa)
        End If

        If IdHabilidade > 0 Then
            strSQL.Append(" and RH72_ID_HABILIDADE = " & IdHabilidade)
        End If

        If IdServidor > 0 Then
            strSQL.Append(" and SERVIDOR.RH02_ID_SERVIDOR = " & IdServidor)
        End If

        If IdEtapa > 0 Then
            strSQL.Append(" and ETAPA.DE07_ID_ETAPA = " & IdEtapa)
        End If

        If IdDisciplina > 0 Then
            strSQL.Append(" and DISCIPLINA.DE09_ID_DISCIPLINA = " & IdDisciplina)
        End If

        If TpHabilidade > 0 Then
            strSQL.Append(" and RH72_TP_HABILIDADE = " & TpHabilidade)
        End If

        If IdUsuario > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & IdUsuario)
        End If

        If IdUsuarioAlteracao > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO_ALT = " & IdUsuarioAlteracao)
        End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" and RH72_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
        End If

        If IsDate(DataHoraDesativacao) Then
            strSQL.Append(" and RH72_DH_DESATIVACAO = Convert(DateTime, '" & DataHoraDesativacao & "', 103)")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH72_ID_HABILIDADE", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select RH72_ID_HABILIDADE as CODIGO, RH02_ID_SERVIDOR as DESCRICAO")
        strSQL.Append(" from RH72_HABILIDADE")
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

        strSQL.Append(" select max(RH72_ID_HABILIDADE) from RH72_HABILIDADE")

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
    'Public Function Excluir(ByVal IdHabilidade As String) As Integer
    '    Dim cnn As New Conexao
    '    Dim strSQL As New StringBuilder
    '    Dim LinhasAfetadas As Integer

    '    strSQL.Append(" delete ")
    '    strSQL.Append(" from RH72_HABILIDADE")
    '    strSQL.Append(" where RH72_ID_HABILIDADE = " & IdHabilidade)

    '    LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

    '    cnn.FecharBanco()
    '    cnn = Nothing

    '    Return LinhasAfetadas
    'End Function

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
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

'******************************************************************************
'*                                 07/03/2019                                 *
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

