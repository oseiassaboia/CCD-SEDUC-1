Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Text

Public Class Session
    Private CA10_COD_SESSION As Integer
    Private FKCA10CA04_COD_USUARIO As Integer
    Private CA10_SESSION_ID As String
    Private CA10_DATA As String
    Private CA10_ATIVO As Boolean
    Private CA10_DATA_ATUAL As String
    Private CA10_IP As String
    Private CA10_NOME_COMPUTADOR As String

    Public Property Codigo() As Integer
        Get
            Return CA10_COD_SESSION
        End Get
        Set(ByVal Value As Integer)
            CA10_COD_SESSION = Value
        End Set
    End Property
    Public Property Usuario() As Integer
        Get
            Return FKCA10CA04_COD_USUARIO
        End Get
        Set(ByVal Value As Integer)
            FKCA10CA04_COD_USUARIO = Value
        End Set
    End Property
    Public Property SessionID() As String
        Get
            Return CA10_SESSION_ID
        End Get
        Set(ByVal Value As String)
            CA10_SESSION_ID = Value
        End Set
    End Property
    Public Property DataEntrada() As String
        Get
            Return CA10_DATA
        End Get
        Set(ByVal Value As String)
            CA10_DATA = Value
        End Set
    End Property
    Public Property Ativo() As Boolean
        Get
            Return CA10_ATIVO
        End Get
        Set(ByVal Value As Boolean)
            CA10_ATIVO = Value
        End Set
    End Property
    Public Property DataAtual() As String
        Get
            Return CA10_DATA_ATUAL
        End Get
        Set(ByVal Value As String)
            CA10_DATA_ATUAL = Value
        End Set
    End Property
    Public Property IP() As String
        Get
            Return CA10_IP
        End Get
        Set(ByVal Value As String)
            CA10_IP = Value
        End Set
    End Property
    Public Property NomeComputador() As String
        Get
            Return CA10_NOME_COMPUTADOR
        End Get
        Set(ByVal Value As String)
            CA10_NOME_COMPUTADOR = Value
        End Set
    End Property

    Public Sub New(Optional ByVal Codigo As Integer = 0)
        If Codigo > 0 Then
            Obter(Codigo)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New ConexaoAcesso
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from CA10_SESSION")
        strSQL.Append(" where CA10_COD_SESSION = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("FKCA10CA04_COD_USUARIO") = ProBanco(FKCA10CA04_COD_USUARIO, eTipoValor.CHAVE)
        dr("CA10_SESSION_ID") = ProBanco(CA10_SESSION_ID, eTipoValor.TEXTO_LIVRE)
        dr("CA10_DATA") = ProBanco(CA10_DATA, eTipoValor.DATA_COMPLETA)
        dr("CA10_ATIVO") = ProBanco(CA10_ATIVO, eTipoValor.BOOLEANO)
        dr("CA10_DATA_ATUAL") = ProBanco(CA10_DATA_ATUAL, eTipoValor.DATA_COMPLETA)
        dr("CA10_IP") = ProBanco(CA10_IP, eTipoValor.TEXTO_LIVRE)
        dr("CA10_NOME_COMPUTADOR") = ProBanco(CA10_NOME_COMPUTADOR, eTipoValor.TEXTO_LIVRE)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal Codigo As Integer)
        Dim cnn As New ConexaoAcesso
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from CA10_SESSION")
        strSQL.Append(" where CA10_COD_SESSION = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            CA10_COD_SESSION = DoBanco(dr("CA10_COD_SESSION"), eTipoValor.CHAVE)
            FKCA10CA04_COD_USUARIO = DoBanco(dr("FKCA10CA04_COD_USUARIO"), eTipoValor.CHAVE)
            CA10_SESSION_ID = DoBanco(dr("CA10_SESSION_ID"), eTipoValor.TEXTO_LIVRE)
            CA10_DATA = DoBanco(dr("CA10_DATA"), eTipoValor.DATA_COMPLETA)
            CA10_ATIVO = DoBanco(dr("CA10_ATIVO"), eTipoValor.BOOLEANO)
            CA10_DATA_ATUAL = DoBanco(dr("CA10_DATA_ATUAL"), eTipoValor.DATA_COMPLETA)
            CA10_IP = DoBanco(dr("CA10_IP"), eTipoValor.TEXTO)
            CA10_NOME_COMPUTADOR = DoBanco(dr("CA10_NOME_COMPUTADOR"), eTipoValor.TEXTO_LIVRE)
        End If

        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional ByVal Codigo As Integer = 0, Optional ByVal Usuario As Integer = 0, Optional ByVal SessionID As String = "", Optional ByVal DataEntrada As String = "", Optional ByVal DataAtual As String = "", Optional ByVal ApenasAtivos As Boolean = False, Optional ByVal IP As String = "", Optional ByVal NomeComputador As String = "") As DataTable
        Dim cnn As New ConexaoAcesso
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from CA10_SESSION")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where CA10_COD_SESSION is not null")

        If Codigo > 0 Then
            strSQL.Append(" and CA10_COD_SESSION = " & Codigo)
        End If

        If Usuario > 0 Then
            strSQL.Append(" and FKCA10CA04_COD_USUARIO = " & Usuario)
        End If

        If SessionID <> "" Then
            strSQL.Append(" and upper(CA10_SESSION_ID) like '%" & SessionID.ToUpper & "%'")
        End If

        If IsDate(DataEntrada) Then
            strSQL.Append(" and CA10_DATA = Convert(DateTime, '" & DataEntrada & "', 103)")
        End If

        If ApenasAtivos Then
            strSQL.Append(" and CA10_ATIVO = 1")
        End If

        If IsDate(DataAtual) Then
            strSQL.Append(" and CA10_DATA_ATUAL = Convert(DateTime, '" & DataAtual & "', 103)")
        End If

        If IP <> "" Then
            strSQL.Append(" and CA10_IP like '" & IP.ToUpper & "'")
        End If

        If NomeComputador <> "" Then
            strSQL.Append(" and CA10_NOME_COMPUTADOR like '" & NomeComputador.ToUpper & "'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "CA10_COD_SESSION desc", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New ConexaoAcesso
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select CA10_COD_SESSION as CODIGO, CA10_SESSION_ID as DESCRICAO")
        strSQL.Append(" from CA10_SESSION")
        strSQL.Append(" order by 2 ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn = Nothing

        Return dt
    End Function

    Public Function VerificarSessionAtivas(ByVal Usuario As Integer) As DataTable
        Dim cnn As New ConexaoAcesso
        Dim dt As DataTable
        Dim strSQL As New StringBuilder
        Dim aux As Integer = 0

        strSQL.Append(" select CA10_COD_SESSION")
        strSQL.Append(" from CA10_SESSION")
        strSQL.Append(" where CA10_COD_SESSION is not null")
        strSQL.Append(" and FKCA10CA04_COD_USUARIO = " & Usuario)
        strSQL.Append(" and CA10_ATIVO = 1")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            While aux < dt.Rows.Count
                VerificarSession(dt.Rows(aux)("CA10_COD_SESSION"))
                aux = aux + 1
            End While
        End If

        cnn = Nothing

        Return dt
    End Function

    Public Function VerificarSession(ByVal session As String) As DataTable
        Dim cnn As New ConexaoAcesso
        Dim dt As DataTable
        Dim strSQL As New StringBuilder
        Dim aux As Integer = 0

        strSQL.Append(" select CA10_COD_SESSION")
        strSQL.Append(" from CA10_SESSION")
        strSQL.Append(" where CA10_COD_SESSION is not null")
        strSQL.Append(" and upper(CA10_SESSION_ID) like '%" & session.ToUpper & "%'")
        strSQL.Append(" and CA10_ATIVO = 1")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            While aux < dt.Rows.Count
                VerificarSession(dt.Rows(aux)("CA10_COD_SESSION"))
                aux = aux + 1
            End While
        End If

        cnn = Nothing

        Return dt
    End Function

    Public Sub VerificarSession(ByVal Session As Integer)
        Dim objSession As New Session(Session)
        Dim runLength As Global.System.TimeSpan = Now.Subtract(objSession.DataAtual)
        Dim Minutos As Double = runLength.TotalMinutes

        If Minutos <= 20 Then
            objSession.DataAtual = Now
        Else
            objSession.Ativo = False
        End If

        objSession.Salvar()

        objSession = Nothing
    End Sub

    Public Function ObterSessionAtiva(ByVal Usuario As Integer) As Integer
        Dim cnn As New ConexaoAcesso
        Dim strSQL As New StringBuilder
        Dim CodigoSession As Integer

        strSQL.Append(" select top 1 CA10_COD_SESSION ")
        strSQL.Append(" from CA10_SESSION")
        strSQL.Append(" where CA10_COD_SESSION is not null")
        strSQL.Append(" and CA10_ATIVO = 1")
        strSQL.Append(" and FKCA10CA04_COD_USUARIO = " & Usuario)

        With cnn.AbrirDataTable(strSQL.ToString)
            If Not IsDBNull(.Rows(0)(0)) Then
                CodigoSession = .Rows(0)(0)
            Else
                CodigoSession = 0
            End If
        End With

        '
        cnn = Nothing

        Return CodigoSession

    End Function

    Public Function ObterUltimo() As Integer
        Dim cnn As New ConexaoAcesso
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(CA10_COD_SESSION) from CA10_SESSION")

        With cnn.AbrirDataTable(strSQL.ToString)
            If Not IsDBNull(.Rows(0)(0)) Then
                CodigoUltimo = .Rows(0)(0)
            Else
                CodigoUltimo = 0
            End If
        End With

        '
        cnn = Nothing

        Return CodigoUltimo

    End Function

    Public Function Excluir(ByVal Codigo As Integer) As Integer
        Dim cnn As New ConexaoAcesso
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from CA10_SESSION")
        strSQL.Append(" where CA10_COD_SESSION = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        '
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class

'******************************************************************************
'*                                 19/08/2011                                 *
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

