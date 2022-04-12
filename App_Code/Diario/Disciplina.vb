Imports Microsoft.VisualBasic
Imports System.Data

Public Class Disciplina
	Private DE09_ID_DISCIPLINA as Integer
	Private DE26_ID_ATIVIDADE as Integer
    Private DE45_ID_MATERIA As Integer
    Private DE70_ID_AREA_CONHECIMENTO As Integer
    Private DE72_ID_CATEGORIA_DISCIPLINA As Integer
    Private DE09_CD_DISCIPLINA as String
	Private DE09_SG_DISCIPLINA as String
	Private DE09_NM_DISCIPLINA as String
    Private DE09_IN_CONDICAO_ESPECIAL As Boolean
    Private DE09_DS_EMENTA As String
    Private DE09_DH_DESATIVACAO As String

    Public Property Codigo() as Integer
		Get
			Return DE09_ID_DISCIPLINA
		End Get
		Set(ByVal Value As Integer)
			DE09_ID_DISCIPLINA = Value
		End Set
	End Property
	Public Property Atividade() as Integer
		Get
			Return DE26_ID_ATIVIDADE
		End Get
		Set(ByVal Value As Integer)
			DE26_ID_ATIVIDADE = Value
		End Set
	End Property
    Public Property Materia() As Integer
        Get
            Return DE45_ID_MATERIA
        End Get
        Set(ByVal Value As Integer)
            DE45_ID_MATERIA = Value
        End Set
    End Property
    Public Property AreaConhecimento() As Integer
        Get
            Return DE70_ID_AREA_CONHECIMENTO
        End Get
        Set(ByVal Value As Integer)
            DE70_ID_AREA_CONHECIMENTO = Value
        End Set
    End Property
    Public Property TipoComponenteCurricular() As Integer
        Get
            Return DE72_ID_CATEGORIA_DISCIPLINA
        End Get
        Set(ByVal Value As Integer)
            DE72_ID_CATEGORIA_DISCIPLINA = Value
        End Set
    End Property
    Public Property CodigoCenso() as String
		Get
			Return DE09_CD_DISCIPLINA
		End Get
		Set(ByVal Value As String)
			DE09_CD_DISCIPLINA = Value
		End Set
	End Property
	Public Property Sigla() as String
		Get
			Return DE09_SG_DISCIPLINA
		End Get
		Set(ByVal Value As String)
			DE09_SG_DISCIPLINA = Value
		End Set
	End Property
	Public Property Nome() as String
		Get
			Return DE09_NM_DISCIPLINA
		End Get
		Set(ByVal Value As String)
			DE09_NM_DISCIPLINA = Value
		End Set
	End Property
    Public Property CondicaoEspecial() As Boolean
        Get
            Return DE09_IN_CONDICAO_ESPECIAL
        End Get
        Set(ByVal Value As Boolean)
            DE09_IN_CONDICAO_ESPECIAL = Value
        End Set
    End Property

    Public Property Ementa() As String
        Get
            Return DE09_DS_EMENTA
        End Get
        Set(ByVal Value As String)
            DE09_DS_EMENTA = Value
        End Set
    End Property
    Public Property DataHoraDesativacao() As String
        Get
            Return DE09_DH_DESATIVACAO
        End Get
        Set(ByVal Value As String)
            DE09_DH_DESATIVACAO = Value
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
		
		strSQL.Append(" Select * ")
        strSQL.Append(" from DBDIARIO..DE09_DISCIPLINA")
        strSQL.Append(" where DE09_ID_DISCIPLINA = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("DE26_ID_ATIVIDADE") = ProBanco(DE26_ID_ATIVIDADE, eTipoValor.CHAVE)
        dr("DE45_ID_MATERIA") = ProBanco(DE45_ID_MATERIA, eTipoValor.CHAVE)
        dr("DE70_ID_AREA_CONHECIMENTO") = ProBanco(DE70_ID_AREA_CONHECIMENTO, eTipoValor.CHAVE)
        dr("DE72_ID_CATEGORIA_DISCIPLINA") = ProBanco(DE72_ID_CATEGORIA_DISCIPLINA, eTipoValor.CHAVE)
        dr("DE09_CD_DISCIPLINA") = ProBanco(DE09_CD_DISCIPLINA, eTipoValor.TEXTO)
		dr("DE09_SG_DISCIPLINA") = ProBanco(DE09_SG_DISCIPLINA, eTipoValor.TEXTO)
		dr("DE09_NM_DISCIPLINA") = ProBanco(DE09_NM_DISCIPLINA, eTipoValor.TEXTO)
        dr("DE09_IN_CONDICAO_ESPECIAL") = ProBanco(DE09_IN_CONDICAO_ESPECIAL, eTipoValor.BOOLEANO)
        dr("DE09_DS_EMENTA") = ProBanco(DE09_DS_EMENTA, eTipoValor.TEXTO_LIVRE)
        dr("DE09_DH_DESATIVACAO") = ProBanco(DE09_DH_DESATIVACAO, eTipoValor.DATA_COMPLETA)


        cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

    Public Sub Obter(ByVal Codigo As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select * ")
        strSQL.Append(" from DBDIARIO..DE09_DISCIPLINA")
        strSQL.Append(" where DE09_ID_DISCIPLINA = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            DE09_ID_DISCIPLINA = DoBanco(dr("DE09_ID_DISCIPLINA"), eTipoValor.CHAVE)
            DE26_ID_ATIVIDADE = DoBanco(dr("DE26_ID_ATIVIDADE"), eTipoValor.CHAVE)
            DE45_ID_MATERIA = DoBanco(dr("DE45_ID_MATERIA"), eTipoValor.CHAVE)
            DE70_ID_AREA_CONHECIMENTO = DoBanco(dr("DE70_ID_AREA_CONHECIMENTO"), eTipoValor.CHAVE)
            DE72_ID_CATEGORIA_DISCIPLINA = DoBanco(dr("DE72_ID_CATEGORIA_DISCIPLINA"), eTipoValor.CHAVE)
            DE09_CD_DISCIPLINA = DoBanco(dr("DE09_CD_DISCIPLINA"), eTipoValor.TEXTO)
            DE09_SG_DISCIPLINA = DoBanco(dr("DE09_SG_DISCIPLINA"), eTipoValor.TEXTO)
            DE09_NM_DISCIPLINA = DoBanco(dr("DE09_NM_DISCIPLINA"), eTipoValor.TEXTO)
            DE09_IN_CONDICAO_ESPECIAL = DoBanco(dr("DE09_IN_CONDICAO_ESPECIAL"), eTipoValor.BOOLEANO)
            DE09_DS_EMENTA = DoBanco(dr("DE09_DS_EMENTA"), eTipoValor.TEXTO_LIVRE)
            DE09_DH_DESATIVACAO = DoBanco(dr("DE09_DH_DESATIVACAO"), eTipoValor.DATA_COMPLETA)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "",
                              Optional Codigo As Integer = 0,
                              Optional Atividade As Integer = 0,
                              Optional Materia As Integer = 0,
                              Optional CodigoCenso As String = "",
                              Optional Sigla As String = "",
                              Optional Nome As String = "",
                              Optional EletivaComum As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select top 50  DE09_ID_DISCIPLINA, DE09.DE26_ID_ATIVIDADE, DE09.DE45_ID_MATERIA ")
        strSQL.Append(" , DE09_SG_DISCIPLINA, DE09_NM_DISCIPLINA, DE09_IN_CONDICAO_ESPECIAL")
        strSQL.Append(" , DE09_CD_DISCIPLINA, iif(DE09_IN_CONDICAO_ESPECIAL = 1, 'SIM', 'NAO') as CONDICAO_ESPECIAL")
        strSQL.Append(" , DE26_NM_ATIVIDADE, DE45_NM_MATERIA")
        strSQL.Append(" from DBDIARIO..DE09_DISCIPLINA as DE09")
        strSQL.Append(" left join DBDIARIO..DE26_ATIVIDADE as DE26 on DE26.DE26_ID_ATIVIDADE = DE09.DE26_ID_ATIVIDADE ")
        strSQL.Append(" left join DBDIARIO..DE45_MATERIA as DE45 on DE45.DE45_ID_MATERIA = DE09.DE45_ID_MATERIA ")
        strSQL.Append(" where DE09_ID_DISCIPLINA Is Not null")

        If Codigo > 0 Then
            strSQL.Append(" And DE09_ID_DISCIPLINA = " & Codigo)
        End If

        If Atividade > 0 Then
            strSQL.Append(" And DE09.DE26_ID_ATIVIDADE = " & Atividade)
        End If

        If Materia > 0 Then
            strSQL.Append(" And vDE45_ID_MATERIA = " & Materia)
        End If

        If CodigoCenso <> "" Then
            strSQL.Append(" And upper(DE09_CD_DISCIPLINA) like '%" & CodigoCenso.ToUpper & "%'")
        End If

        If Sigla <> "" Then
            strSQL.Append(" and upper(DE09_SG_DISCIPLINA) like '%" & Sigla.ToUpper & "%'")
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(DE09_NM_DISCIPLINA) like '%" & Nome.ToUpper & "%'")
        End If

        If EletivaComum > 0 Then
            If EletivaComum = 1 Then
                strSQL.Append("  And DE09.DE26_ID_ATIVIDADE is null ")
            Else
                strSQL.Append(" And DE09.DE26_ID_ATIVIDADE is not null ")
            End If

        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "DE09_ID_DISCIPLINA", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarCargaHorariaDisciplina(Optional ByVal Sort As String = "",
                                      Optional Lotacao As Integer = 0,
                                      Optional Etapa As Integer = 0,
                                      Optional Nome As String = "",
                                      Optional EletivaComum As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select top 50 DE09.DE09_ID_DISCIPLINA, DE09_NM_DISCIPLINA, DE07.DE07_ID_ETAPA, DE26.DE26_ID_ATIVIDADE")
        strSQL.Append(" , DE09_SG_DISCIPLINA, DE09_NM_DISCIPLINA, DE09_IN_CONDICAO_ESPECIAL")
        strSQL.Append(" , DE09_CD_DISCIPLINA, iif(DE09_IN_CONDICAO_ESPECIAL = 1, 'SIM', 'NAO') as CONDICAO_ESPECIAL")
        strSQL.Append(" , DE26_NM_ATIVIDADE, DE45_NM_MATERIA")
        strSQL.Append(" from DBDIARIO..DE09_DISCIPLINA as DE09 ")
        strSQL.Append(" left join  DBDIARIO..DE26_ATIVIDADE as DE26 on DE26.DE26_ID_ATIVIDADE = DE09.DE26_ID_ATIVIDADE ")
        strSQL.Append(" left join DBDIARIO..DE45_MATERIA as DE45 on DE45.DE45_ID_MATERIA = DE09.DE45_ID_MATERIA ")
        strSQL.Append(" left join DBDIARIO..DE11_ETAPA_DISCIPLINA as DE11 on DE11.DE09_ID_DISCIPLINA = DE09.DE09_ID_DISCIPLINA And DE11.DE11_DH_DESATIVACAO Is null  ")
        strSQL.Append(" left join DBDIARIO..DE07_ETAPA as DE07 on DE07.DE07_ID_ETAPA = DE11.DE07_ID_ETAPA  And DE07.DE07_DH_DESATIVACAO Is null  ")
        strSQL.Append(" where DE09.DE09_ID_DISCIPLINA Is Not null ")


        If EletivaComum > 0 Then
            If EletivaComum = 1 Then
                strSQL.Append("   and DE26.DE26_ID_ATIVIDADE is null ")
            Else
                strSQL.Append("   and DE26.DE26_ID_ATIVIDADE is not null ")
            End If

        End If

        If Etapa > 0 Then
            strSQL.Append(" and DE07.DE07_ID_ETAPA in (  ")
            strSQL.Append(" Select DE07_ID_ETAPA  From DBDIARIO..DE31_ESCOLA_ETAPA  Where RH36_ID_LOTACAO = " & Lotacao & ") ")
            strSQL.Append(" and DE07.DE07_ID_ETAPA = " & Etapa)
        End If


        If Nome <> "" Then
            strSQL.Append(" and upper(DE09_NM_DISCIPLINA) like '%" & Nome.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "DE09.DE09_ID_DISCIPLINA", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarDisciplinaEtapas(Optional ByVal Sort As String = "", Optional Disciplina As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select top 50 DE09.DE09_ID_DISCIPLINA, DE09.DE09_NM_DISCIPLINA, DE07.DE07_ID_ETAPA, DE07.DE07_DH_DESATIVACAO ")
        strSQL.Append(" ,  DE05.DE05_NM_MODALIDADE + ' - ' + DE06.DE06_NM_NIVEL + ' - ' + DE07_NM_ETAPA + isnull(' - ' + DE12.DE12_NM_AREA + ' - ' + DE15.DE15_NM_CURSO , '') as ETAPA ")
        strSQL.Append(" , IIF(DE07.DE07_DH_DESATIVACAO is null,'ATIVADO', 'DESATIVADO') AS ATIVOS")
        strSQL.Append(" from DBDIARIO..DE09_DISCIPLINA as DE09")
        strSQL.Append(" left join DBDIARIO..DE11_ETAPA_DISCIPLINA  as DE11 on DE11.DE09_ID_DISCIPLINA = DE09.DE09_ID_DISCIPLINA ")
        strSQL.Append(" left join DBDIARIO..DE07_ETAPA as DE07 on DE07.DE07_ID_ETAPA = DE11.DE07_ID_ETAPA ")
        strSQL.Append(" left join DBDIARIO..DE06_NIVEL as DE06 on DE06.DE06_ID_NIVEL = DE07.DE06_ID_NIVEL ")
        strSQL.Append(" left join DBDIARIO..DE05_MODALIDADE as DE05 on DE05.DE05_ID_MODALIDADE = DE06.DE05_ID_MODALIDADE ")
        strSQL.Append(" Left Join DBDIARIO..DE15_CURSO as DE15 on DE15.DE15_ID_CURSO = DE07.DE15_ID_CURSO ")
        strSQL.Append(" Left Join DBDIARIO..DE12_AREA as DE12 on DE12.DE12_ID_AREA = DE15.DE12_ID_AREA ")
        strSQL.Append(" where DE09.DE09_ID_DISCIPLINA Is Not null")

        If Disciplina > 0 Then
            strSQL.Append(" And DE09.DE09_ID_DISCIPLINA = " & Disciplina)
        End If


        strSQL.Append(" Order By " & IIf(Sort = "", "DE05.DE05_NM_MODALIDADE + ' - ' + DE06.DE06_NM_NIVEL + ' - ' + DE07_NM_ETAPA + isnull(' - ' + DE12.DE12_NM_AREA + ' - ' + DE15.DE15_NM_CURSO , '')", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela(Optional CategoriaDisciplina As Integer = 0, Optional Etapa As Integer = 0, Optional TodasAsDisciplinas As Boolean = False) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select DE09.DE09_ID_DISCIPLINA as CODIGO, DE09.DE09_NM_DISCIPLINA as DESCRICAO")
        strSQL.Append(" from DBDIARIO..DE09_DISCIPLINA as DE09 ")
        strSQL.Append(" where DE09.DE09_ID_DISCIPLINA is not null ")
        strSQL.Append(" And DE09_DH_DESATIVACAO Is null ")

        If Etapa > 0 Then
            If TodasAsDisciplinas Then
                strSQL.Append(" and DE09.DE09_ID_DISCIPLINA in ( ")
            Else
                strSQL.Append(" and DE09.DE09_ID_DISCIPLINA not in ( ")
            End If

            strSQL.Append("     select DE11.DE09_ID_DISCIPLINA ")
            strSQL.Append("     from DBDIARIO..DE11_ETAPA_DISCIPLINA as DE11 ")
            strSQL.Append("     where DE11.DE07_ID_ETAPA = " & Etapa)
            strSQL.Append("     And DE11_DH_DESATIVACAO Is null) ")
        End If

        If CategoriaDisciplina > 0 Then
            strSQL.Append(" and DE09.DE72_ID_CATEGORIA_DISCIPLINA = " & CategoriaDisciplina)
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

        strSQL.Append(" Select max(DE09_ID_DISCIPLINA) from DBDIARIO..DE09_DISCIPLINA")

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
    Public Function Excluir(ByVal Codigo As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from DBDIARIO..DE09_DISCIPLINA")
        strSQL.Append(" where DE09_ID_DISCIPLINA = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class


