Imports Microsoft.VisualBasic
Imports System.Data

Public Class Etapa
	Private DE07_ID_ETAPA as Integer
    Private DE06_ID_NIVEL As Integer
    Private DE15_ID_CURSO As Integer
    Private DE07_NM_ETAPA as String
    Private DE07_CD_ETAPA_CENSO As String
    Private DE07_NR_IDADE_MINIMA as String
    Private DE07_NR_IDADE_MAXIMA As String
    Private DE07_PC_REPROVACAO_FALTA As String
    Private DE07_VL_MEDIA_APROVACAO As String
    Private DE07_QT_MAX_PENDENCIA_DISC As String
    Private DE07_CD_PREFIXO_NU_TURMA As String
    Private DE07_DH_CADASTRO as String
	Private DE07_DH_DESATIVACAO as String

	Public Property Codigo() as Integer
		Get
			Return DE07_ID_ETAPA
		End Get
		Set(ByVal Value As Integer)
			DE07_ID_ETAPA = Value
		End Set
	End Property
    Public Property Nivel() As Integer
        Get
            Return DE06_ID_NIVEL
        End Get
        Set(ByVal Value As Integer)
            DE06_ID_NIVEL = Value
        End Set
    End Property
    Public Property Curso() As Integer
        Get
            Return DE15_ID_CURSO
        End Get
        Set(ByVal Value As Integer)
            DE15_ID_CURSO = Value
        End Set
    End Property
    Public Property Nome() As String
        Get
            Return DE07_NM_ETAPA
        End Get
        Set(ByVal Value As String)
            DE07_NM_ETAPA = Value
        End Set
    End Property
    Public Property CodigoCenso() as String
		Get
			Return DE07_CD_ETAPA_CENSO
		End Get
		Set(ByVal Value As String)
			DE07_CD_ETAPA_CENSO = Value
		End Set
	End Property
	Public Property IdadeMinima() as String
		Get
			Return DE07_NR_IDADE_MINIMA
		End Get
		Set(ByVal Value As String)
			DE07_NR_IDADE_MINIMA = Value
		End Set
	End Property
    Public Property IdadeMaxima() As String
        Get
            Return DE07_NR_IDADE_MAXIMA
        End Get
        Set(ByVal Value As String)
            DE07_NR_IDADE_MAXIMA = Value
        End Set
    End Property

    Public Property PercentualReprovacaoFalta() As String
        Get
            Return DE07_PC_REPROVACAO_FALTA
        End Get
        Set(ByVal Value As String)
            DE07_PC_REPROVACAO_FALTA = Value
        End Set
    End Property

    Public Property MediaAprovacao() As String
        Get
            Return DE07_VL_MEDIA_APROVACAO
        End Get
        Set(ByVal Value As String)
            DE07_VL_MEDIA_APROVACAO = Value
        End Set
    End Property

    Public Property QtdMaximaPendenciaDisciplina() As String
        Get
            Return DE07_QT_MAX_PENDENCIA_DISC
        End Get
        Set(ByVal Value As String)
            DE07_QT_MAX_PENDENCIA_DISC = Value
        End Set
    End Property

    Public Property PrefixoNumeroturma() As String
        Get
            Return DE07_CD_PREFIXO_NU_TURMA
        End Get
        Set(ByVal Value As String)
            DE07_CD_PREFIXO_NU_TURMA = Value
        End Set
    End Property
    Public Property DataHoraCadastro() as String
		Get
			Return DE07_DH_CADASTRO
		End Get
		Set(ByVal Value As String)
			DE07_DH_CADASTRO = Value
		End Set
	End Property
	Public Property DataHoraDesativacao() as String
		Get
			Return DE07_DH_DESATIVACAO
		End Get
		Set(ByVal Value As String)
			DE07_DH_DESATIVACAO = Value
		End Set
	End Property

    Public Sub New(Optional ByVal Codigo As Integer = 0)
        If Codigo > 0 Then
            Obter(Codigo)
        End If
    End Sub

    Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
        strSQL.Append(" from DBDIARIO..DE07_ETAPA")
        strSQL.Append(" where DE07_ID_ETAPA = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
            dr = dt.NewRow

        Else
			dr = dt.Rows(0)
		End If

        dr("DE06_ID_NIVEL") = ProBanco(DE06_ID_NIVEL, eTipoValor.CHAVE)
        dr("DE15_ID_CURSO") = ProBanco(DE15_ID_CURSO, eTipoValor.CHAVE)
        dr("DE07_NM_ETAPA") = ProBanco(DE07_NM_ETAPA, eTipoValor.TEXTO)
        dr("DE07_CD_ETAPA_CENSO") = ProBanco(DE07_CD_ETAPA_CENSO, eTipoValor.TEXTO)
        dr("DE07_NR_IDADE_MINIMA") = ProBanco(DE07_NR_IDADE_MINIMA, eTipoValor.NUMERO_INTEIRO)
        dr("DE07_NR_IDADE_MAXIMA") = ProBanco(DE07_NR_IDADE_MAXIMA, eTipoValor.NUMERO_INTEIRO)
        dr("DE07_PC_REPROVACAO_FALTA") = ProBanco(DE07_PC_REPROVACAO_FALTA, eTipoValor.MONETARIO)
        dr("DE07_VL_MEDIA_APROVACAO") = ProBanco(DE07_VL_MEDIA_APROVACAO, eTipoValor.MONETARIO)
        dr("DE07_QT_MAX_PENDENCIA_DISC") = ProBanco(DE07_QT_MAX_PENDENCIA_DISC, eTipoValor.NUMERO_INTEIRO)
        dr("DE07_CD_PREFIXO_NU_TURMA") = ProBanco(DE07_CD_PREFIXO_NU_TURMA, eTipoValor.TEXTO)
        dr("DE07_DH_DESATIVACAO") = ProBanco(DE07_DH_DESATIVACAO, eTipoValor.DATA_COMPLETA)
        dr("DE07_DH_CADASTRO") = ProBanco(DE07_DH_CADASTRO, eTipoValor.DATA_COMPLETA)

        cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal Codigo as Integer)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
        strSQL.Append(" from DBDIARIO..DE07_ETAPA")
        strSQL.Append(" where DE07_ID_ETAPA = " & Codigo)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			DE07_ID_ETAPA = DoBanco(dr("DE07_ID_ETAPA"), eTipoValor.CHAVE)
            DE06_ID_NIVEL = DoBanco(dr("DE06_ID_NIVEL"), eTipoValor.CHAVE)
            DE15_ID_CURSO = DoBanco(dr("DE15_ID_CURSO"), eTipoValor.CHAVE)
            DE07_NM_ETAPA = DoBanco(dr("DE07_NM_ETAPA"), eTipoValor.TEXTO)
            DE07_CD_ETAPA_CENSO = DoBanco(dr("DE07_CD_ETAPA_CENSO"), eTipoValor.TEXTO)
            DE07_NR_IDADE_MINIMA = DoBanco(dr("DE07_NR_IDADE_MINIMA"), eTipoValor.NUMERO_INTEIRO)
            DE07_NR_IDADE_MAXIMA = DoBanco(dr("DE07_NR_IDADE_MAXIMA"), eTipoValor.NUMERO_INTEIRO)
            DE07_PC_REPROVACAO_FALTA = DoBanco(dr("DE07_PC_REPROVACAO_FALTA"), eTipoValor.MONETARIO)
            DE07_VL_MEDIA_APROVACAO = DoBanco(dr("DE07_VL_MEDIA_APROVACAO"), eTipoValor.MONETARIO)
            DE07_QT_MAX_PENDENCIA_DISC = DoBanco(dr("DE07_QT_MAX_PENDENCIA_DISC"), eTipoValor.NUMERO_INTEIRO)
            DE07_CD_PREFIXO_NU_TURMA = DoBanco(dr("DE07_CD_PREFIXO_NU_TURMA"), eTipoValor.TEXTO)
            DE07_DH_CADASTRO = DoBanco(dr("DE07_DH_CADASTRO"), eTipoValor.DATA_COMPLETA)
            DE07_DH_DESATIVACAO = DoBanco(dr("DE07_DH_DESATIVACAO"), eTipoValor.DATA_COMPLETA)
        End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "",
                              Optional Codigo As Integer = 0,
                              Optional Nivel As Integer = 0,
                              Optional Nome As String = "",
                              Optional CodigoCenso As String = "",
                              Optional IdadeMinima As String = "",
                              Optional IdadeMaxima As String = "",
                              Optional DataHoraCadastro As String = "",
                              Optional DataHoraDesativacao As String = "",
                              Optional Desativado As Boolean = False,
                              Optional Modalidade As Integer = 0,
                              Optional Curso As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select top 50 DE07_ID_ETAPA, DE07.DE06_ID_NIVEL, DE07_NM_ETAPA, DE07_PC_REPROVACAO_FALTA, DE07_QT_MAX_PENDENCIA_DISC ")
        strSQL.Append(" , DE07_CD_ETAPA_CENSO, DE07_NR_IDADE_MINIMA, DE07_NR_IDADE_MAXIMA, DE07_DH_DESATIVACAO")
        strSQL.Append(" , DE06.DE06_NM_NIVEL, DE05.DE05_NM_MODALIDADE + ' - ' + DE06.DE06_NM_NIVEL + ' - ' + DE07_NM_ETAPA + isnull(' - ' + DE12.DE12_NM_AREA + ' - ' + DE15.DE15_NM_CURSO , '') as DESCRICAO")
        strSQL.Append(" , CASE WHEN DE07_DH_DESATIVACAO IS NULL THEN 'ATIVADO' ELSE 'DESATIVADO' END as SITUACAO")
        strSQL.Append(" , DE07.DE15_ID_CURSO, DE07.DE07_VL_MEDIA_APROVACAO, DE07.DE07_CD_PREFIXO_NU_TURMA")
        strSQL.Append(" from DBDIARIO..DE07_ETAPA as DE07")
        strSQL.Append(" left join DBDIARIO..DE06_NIVEL as DE06 on DE06.DE06_ID_NIVEL = DE07.DE06_ID_NIVEL ")
        strSQL.Append(" left join DBDIARIO..DE05_MODALIDADE as DE05 on DE05.DE05_ID_MODALIDADE = DE06.DE05_ID_MODALIDADE ")
        strSQL.Append(" left join DBDIARIO..DE15_CURSO as DE15 on DE15.DE15_ID_CURSO = DE07.DE15_ID_CURSO ")
        strSQL.Append("  left join DBDIARIO..DE12_AREA as DE12 on DE12.DE12_ID_AREA = DE15.DE12_ID_AREA  ")
        strSQL.Append(" where DE07_ID_ETAPA is not null")

        If Codigo > 0 Then
            strSQL.Append(" and DE07_ID_ETAPA = " & Codigo)
        End If

        If Nivel > 0 Then
            strSQL.Append(" and DE07.DE06_ID_NIVEL = " & Nivel)
        End If

        If Curso > 0 Then
            strSQL.Append(" and DE07.DE15_ID_CURSO = " & Curso)
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(DE05.DE05_NM_MODALIDADE + ' - ' + DE06.DE06_NM_NIVEL + ' - ' + DE07_NM_ETAPA + isnull(' - ' + DE12.DE12_NM_AREA + ' - ' + DE15.DE15_NM_CURSO , '')) like '%" & Nome.ToUpper & "%'")
        End If

        If CodigoCenso <> "" Then
            strSQL.Append(" and upper(DE07_CD_ETAPA_CENSO) like '%" & CodigoCenso.ToUpper & "%'")
        End If

        If IdadeMinima <> "" Then
            strSQL.Append(" and upper(DE07_NR_IDADE_MINIMA) like '%" & IdadeMinima.ToUpper & "%'")
        End If

        If IdadeMaxima <> "" Then
            strSQL.Append(" and upper(DE07_NR_IDADE_MAXIMA) like '%" & IdadeMaxima.ToUpper & "%'")
        End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" and DE07_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
        End If

        If IsDate(DataHoraDesativacao) Then
            strSQL.Append(" and DE07_DH_DESATIVACAO = Convert(DateTime, '" & DataHoraDesativacao & "', 103)")
        End If

        If Desativado Then
            strSQL.Append(" and DE07_DH_DESATIVACAO is not null")
        Else
            strSQL.Append(" and DE07_DH_DESATIVACAO is null")
        End If

        If Modalidade > 0 Then
            strSQL.Append(" and DE05.DE05_ID_MODALIDADE = " & Modalidade)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "DE05.DE05_NM_MODALIDADE + ' - ' + DE06.DE06_NM_NIVEL + ' - ' + DE07_NM_ETAPA + isnull(' - ' + DE12.DE12_NM_AREA + ' - ' + DE15.DE15_NM_CURSO , '')", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarNomeEtapa(ByVal Nivel As Integer, ByVal NomeEtapa As String, Optional Curso As Integer = 0) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhaEncontrada As Integer

        strSQL.Append(" select 1 as QTD where exists ( ")
        strSQL.Append("     select DE07_NM_ETAPA ")
        strSQL.Append("     from DBDIARIO..DE07_ETAPA ")
        strSQL.Append("     where DE07_ID_ETAPA is not null")
        strSQL.Append("     and DE06_ID_NIVEL = " & Nivel)
        strSQL.Append("     and DE07_NM_ETAPA = '" & NomeEtapa.ToUpper & "'")

        If Curso > 0 Then
            strSQL.Append(" and DE15_ID_CURSO = " & Curso)
        End If

        strSQL.Append(" )")

        LinhaEncontrada = cnn.AbrirDataTable(strSQL.ToString).Rows.Count

        Return LinhaEncontrada
    End Function

    Public Function ObterTabela(Optional Nivel As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select DE07.DE07_ID_ETAPA As CODIGO, DE05.DE05_NM_MODALIDADE + ' - ' + DE06.DE06_NM_NIVEL + ' - ' + DE07.DE07_NM_ETAPA + ")
        strSQL.Append(" Case when DE07.DE15_ID_CURSO Is null then '' else ' - ' + DE12_NM_AREA + ' - ' + DE15_NM_CURSO end ")
        strSQL.Append(" as DESCRICAO ")
        strSQL.Append(" From DBDIARIO..DE07_ETAPA As DE07  ")
        strSQL.Append(" Left Join DBDIARIO..DE06_NIVEL as DE06 on DE06.DE06_ID_NIVEL = DE07.DE06_ID_NIVEL  ")
        strSQL.Append(" Left Join DBDIARIO..DE05_MODALIDADE as DE05 on DE05.DE05_ID_MODALIDADE = DE06.DE05_ID_MODALIDADE ")
        strSQL.Append(" Left Join DBDIARIO..DE15_CURSO as DE15 on DE15.DE15_ID_CURSO = DE07.DE15_ID_CURSO ")
        strSQL.Append(" Left Join DBDIARIO..DE12_AREA as DE12 on DE12.DE12_ID_AREA = DE15.DE12_ID_AREA ")
        strSQL.Append(" where DE07.DE07_ID_ETAPA Is Not null  ")
        strSQL.Append(" And DE07.DE07_DH_DESATIVACAO Is null  ")

        If Nivel > 0 Then
            strSQL.Append(" and DE06.DE06_ID_NIVEL = " & Nivel)
        End If

        strSQL.Append(" order by 2 ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return dt
    End Function

    Public Function ObterUltimo() as Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim CodigoUltimo As Integer

        strSQL.Append(" Select max(DE07_ID_ETAPA) from DBDIARIO..DE07_ETAPA")

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
	Public Function Excluir(ByVal Codigo as Integer) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
        strSQL.Append(" from DBDIARIO..DE07_ETAPA")
        strSQL.Append(" where DE07_ID_ETAPA = " & Codigo)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class


