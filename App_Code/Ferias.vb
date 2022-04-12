Imports Microsoft.VisualBasic
Imports System.Data

Public Class Ferias

    Implements IDisposable

    Private RH28_ID_FERIAS As Integer
    Private RH02_ID_SERVIDOR As String
    Private RH87_ID_PERIODO_FERIAS As String
    Private CA04_ID_USUARIO As Integer
    Private CA04_ID_USUARIO_ALT As Integer
    Private RH28_DT_INICIO_AQUISICAO As String
    Private RH28_DT_TERMINO_AQUISICAO As String
    Private RH28_DT_INICIO_GOZO As String
    Private RH28_DT_TERMINO_GOZO As String
    Private RH28_DH_CADASTRO As String
    Private RH28_DH_ALTERACAO As String

    Public Property IdFerias() As Integer
        Get
            Return RH28_ID_FERIAS
        End Get
        Set(ByVal Value As Integer)
            RH28_ID_FERIAS = Value
        End Set
    End Property
    Public Property IdServidor() As String
        Get
            Return RH02_ID_SERVIDOR
        End Get
        Set(ByVal Value As String)
            RH02_ID_SERVIDOR = Value
        End Set
    End Property
    Public Property IdPeriodoFerias() As String
        Get
            Return RH87_ID_PERIODO_FERIAS
        End Get
        Set(ByVal Value As String)
            RH87_ID_PERIODO_FERIAS = Value
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
    Public Property DataInicioAquisicao() As String
        Get
            Return RH28_DT_INICIO_AQUISICAO
        End Get
        Set(ByVal Value As String)
            RH28_DT_INICIO_AQUISICAO = Value
        End Set
    End Property
    Public Property DataTerminoAquisicao() As String
        Get
            Return RH28_DT_TERMINO_AQUISICAO
        End Get
        Set(ByVal Value As String)
            RH28_DT_TERMINO_AQUISICAO = Value
        End Set
    End Property
    Public Property DataInicioGozo() As String
        Get
            Return RH28_DT_INICIO_GOZO
        End Get
        Set(ByVal Value As String)
            RH28_DT_INICIO_GOZO = Value
        End Set
    End Property
    Public Property DataTerminoGozo() As String
        Get
            Return RH28_DT_TERMINO_GOZO
        End Get
        Set(ByVal Value As String)
            RH28_DT_TERMINO_GOZO = Value
        End Set
    End Property
    Public Property DataHoraCadastro() As String
        Get
            Return RH28_DH_CADASTRO
        End Get
        Set(ByVal Value As String)
            RH28_DH_CADASTRO = Value
        End Set
    End Property
    Public Property DataHoraAlteracao() As String
        Get
            Return RH28_DH_ALTERACAO
        End Get
        Set(ByVal Value As String)
            RH28_DH_ALTERACAO = Value
        End Set
    End Property

    Public Sub New(Optional ByVal IdFerias As Integer = 0)
        If IdFerias > 0 Then
            Obter(IdFerias)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH28_FERIAS")
        strSQL.Append(" where RH28_ID_FERIAS = " & IdFerias)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH02_ID_SERVIDOR") = ProBanco(RH02_ID_SERVIDOR, eTipoValor.NUMERO_DECIMAL)
        dr("RH87_ID_PERIODO_FERIAS") = ProBanco(RH87_ID_PERIODO_FERIAS, eTipoValor.NUMERO_DECIMAL)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO_ALT") = ProBanco(CA04_ID_USUARIO_ALT, eTipoValor.CHAVE)
        dr("RH28_DT_INICIO_AQUISICAO") = ProBanco(RH28_DT_INICIO_AQUISICAO, eTipoValor.DATA)
        dr("RH28_DT_TERMINO_AQUISICAO") = ProBanco(RH28_DT_TERMINO_AQUISICAO, eTipoValor.DATA)
        dr("RH28_DT_INICIO_GOZO") = ProBanco(RH28_DT_INICIO_GOZO, eTipoValor.DATA)
        dr("RH28_DT_TERMINO_GOZO") = ProBanco(RH28_DT_TERMINO_GOZO, eTipoValor.DATA)
        dr("RH28_DH_CADASTRO") = ProBanco(RH28_DH_CADASTRO, eTipoValor.DATA)
        dr("RH28_DH_ALTERACAO") = ProBanco(RH28_DH_ALTERACAO, eTipoValor.DATA)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal IdFerias As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH28_FERIAS")
        strSQL.Append(" where RH28_ID_FERIAS = " & IdFerias)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH28_ID_FERIAS = DoBanco(dr("RH28_ID_FERIAS"), eTipoValor.CHAVE)
            RH02_ID_SERVIDOR = DoBanco(dr("RH02_ID_SERVIDOR"), eTipoValor.NUMERO_DECIMAL)
            RH87_ID_PERIODO_FERIAS = DoBanco(dr("RH87_ID_PERIODO_FERIAS"), eTipoValor.NUMERO_DECIMAL)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            CA04_ID_USUARIO_ALT = DoBanco(dr("CA04_ID_USUARIO_ALT"), eTipoValor.CHAVE)
            RH28_DT_INICIO_AQUISICAO = DoBanco(dr("RH28_DT_INICIO_AQUISICAO"), eTipoValor.DATA)
            RH28_DT_TERMINO_AQUISICAO = DoBanco(dr("RH28_DT_TERMINO_AQUISICAO"), eTipoValor.DATA)
            RH28_DT_INICIO_GOZO = DoBanco(dr("RH28_DT_INICIO_GOZO"), eTipoValor.DATA)
            RH28_DT_TERMINO_GOZO = DoBanco(dr("RH28_DT_TERMINO_GOZO"), eTipoValor.DATA)
            RH28_DH_CADASTRO = DoBanco(dr("RH28_DH_CADASTRO"), eTipoValor.DATA)
            RH28_DH_ALTERACAO = DoBanco(dr("RH28_DH_ALTERACAO"), eTipoValor.DATA)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional IdFerias As Integer = 0, Optional IdServidor As String = "", Optional IdPeriodoFerias As String = "", Optional IdUsuario As Integer = 0, Optional IdUsuarioAlteracao As Integer = 0, Optional DataInicioAquisicao As String = "", Optional DataTerminoAquisicao As String = "", Optional DataInicioGozo As String = "", Optional DataTerminoGozo As String = "", Optional DataHoraCadastro As String = "", Optional DataHoraAlteracao As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH28_FERIAS")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where RH28_ID_FERIAS is not null")

        If IdFerias > 0 Then
            strSQL.Append(" and RH28_ID_FERIAS = " & IdFerias)
        End If

        If IsNumeric(IdServidor.Replace(".", "")) Then
            strSQL.Append(" and RH02_ID_SERVIDOR = " & IdServidor.Replace(".", "").Replace(",", "."))
        End If

        If IsNumeric(IdPeriodoFerias.Replace(".", "")) Then
            strSQL.Append(" and RH87_ID_PERIODO_FERIAS = " & IdPeriodoFerias.Replace(".", "").Replace(",", "."))
        End If

        If IdUsuario > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & IdUsuario)
        End If

        If IdUsuarioAlteracao > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO_ALT = " & IdUsuarioAlteracao)
        End If

        If IsDate(DataInicioAquisicao) Then
            strSQL.Append(" and RH28_DT_INICIO_AQUISICAO = Convert(DateTime, '" & DataInicioAquisicao & "', 103)")
        End If

        If IsDate(DataTerminoAquisicao) Then
            strSQL.Append(" and RH28_DT_TERMINO_AQUISICAO = Convert(DateTime, '" & DataTerminoAquisicao & "', 103)")
        End If

        If IsDate(DataInicioGozo) Then
            strSQL.Append(" and RH28_DT_INICIO_GOZO = Convert(DateTime, '" & DataInicioGozo & "', 103)")
        End If

        If IsDate(DataTerminoGozo) Then
            strSQL.Append(" and RH28_DT_TERMINO_GOZO = Convert(DateTime, '" & DataTerminoGozo & "', 103)")
        End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" and RH28_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
        End If

        If IsDate(DataHoraAlteracao) Then
            strSQL.Append(" and RH28_DH_ALTERACAO = Convert(DateTime, '" & DataHoraAlteracao & "', 103)")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH28_ID_FERIAS", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function
    Public Function PesquisarServidoresFerias(Optional ByVal Sort As String = "", Optional PeriodoFerias As String = "", Optional ByVal MesGozo As Integer = 0 _
                                              , Optional ByVal Regional As Integer = 0, Optional ByVal Cidade As Integer = 0, Optional ByVal Lotacao As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append("  select RH28_ID_FERIAS,TG05_NM_REGIONAL, TG03_NM_MUNICIPIO, RH36_NM_LOTACAO, RH01_NM_PESSOA, RH02_CD_MATRICULA, RH28.RH28_DT_INICIO_GOZO,RH28.RH28_DT_TERMINO_GOZO  ")
        strSQL.Append("  	from RH28_FERIAS RH28  ")
        strSQL.Append("  	LEFT JOIN	RH02_SERVIDOR			RH02	ON	RH28.RH02_ID_SERVIDOR	=	RH02.RH02_ID_SERVIDOR and RH87_ID_PERIODO_FERIAS = 3  ")
        strSQL.Append("  	LEFT JOIN	RH01_PESSOA				RH01	ON	RH02.RH01_ID_PESSOA		=	RH01.RH01_ID_PESSOA  ")
        strSQL.Append("  	LEFT JOIN	RH14_LOTACAO_SERVIDOR	RH14	ON	RH02.RH02_ID_SERVIDOR	=	RH14.RH02_ID_SERVIDOR and rh14.rh88_id_periodo = 39  ")
        strSQL.Append("  	LEFT JOIN	RH36_LOTACAO			RH36	ON	RH14.RH36_ID_LOTACAO	=	RH36.RH36_ID_LOTACAO  ")
        strSQL.Append("  	LEFT JOIN	DBGERAL..TG03_MUNICIPIO	TG03	ON	RH36.TG03_ID_MUNICIPIO	=	TG03.TG03_ID_MUNICIPIO  ")
        strSQL.Append("  	LEFT JOIN	DBGERAL..TG05_REGIONAL	TG05	ON	TG03.TG05_ID_REGIONAL	=	TG05.TG05_ID_REGIONAL  ")
        strSQL.Append("  	WHERE RH14.RH14_TP_LOTACAO_SERVIDOR='P' AND RH14.RH14_DT_DESLIGAMENTO IS NULL AND RH02.RH07_ID_SITUACAO_SERVIDOR IN (1,10,11)  ")

        If Regional > 0 Then
            strSQL.Append("  and TG05.TG05_ID_REGIONAL =  " & Regional)
        End If

        If Cidade > 0 Then
            strSQL.Append("  and tg03.TG03_ID_MUNICIPIO =  " & Cidade)

        End If

        If Lotacao > 0 Then
            strSQL.Append(" and RH36.RH36_ID_LOTACAO =  " & Lotacao)
        End If

        If MesGozo > 0 Then
            strSQL.Append(" and RH28_DT_INICIO_GOZO =  '" & Date.Now.Year + 1 & "-" & MesGozo & "-01'")
        End If


        strSQL.Append(" Order By " & IIf(Sort = "", "RH28_DT_INICIO_GOZO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select RH28_ID_FERIAS as CODIGO, RH02_ID_SERVIDOR as DESCRICAO")
        strSQL.Append(" from RH28_FERIAS")
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

        strSQL.Append(" select max(RH28_ID_FERIAS) from RH28_FERIAS")

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
    Public Function Excluir(ByVal IdFerias As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from RH28_FERIAS")
        strSQL.Append(" where RH28_ID_FERIAS = " & IdFerias)

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
'*                                 25/10/2019                                 *
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

