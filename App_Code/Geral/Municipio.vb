Imports Microsoft.VisualBasic
Imports System.Data

Public Class Municipio
	Private TG03_ID_MUNICIPIO as Integer
	Private TG02_ID_UF as Integer
	Private TG05_ID_REGIONAL as Integer
	Private TG05_ID_REGIONAL_ANTES as Integer
	Private TG03_NM_MUNICIPIO as String
	Private TG03_CD_IBGE as String
	Private TG03_IN_SEDE_REGIONAL as String
	Private TG03_QT_POPULACAO as String
	Private TG03_NR_DISTANCIA_URE as String

    Public Property Codigo() As Integer
        Get
            Return TG03_ID_MUNICIPIO
        End Get
        Set(ByVal Value As Integer)
            TG03_ID_MUNICIPIO = Value
        End Set
    End Property
    Public Property UF() as Integer
		Get
			Return TG02_ID_UF
		End Get
		Set(ByVal Value As Integer)
			TG02_ID_UF = Value
		End Set
	End Property
	Public Property Regional() as Integer
		Get
			Return TG05_ID_REGIONAL
		End Get
		Set(ByVal Value As Integer)
			TG05_ID_REGIONAL = Value
		End Set
	End Property
	Public Property RegionalAntes() as Integer
		Get
			Return TG05_ID_REGIONAL_ANTES
		End Get
		Set(ByVal Value As Integer)
			TG05_ID_REGIONAL_ANTES = Value
		End Set
	End Property
	Public Property Nome() as String
		Get
			Return TG03_NM_MUNICIPIO
		End Get
		Set(ByVal Value As String)
			TG03_NM_MUNICIPIO = Value
		End Set
	End Property
	Public Property CodigoIBGE() as String
		Get
			Return TG03_CD_IBGE
		End Get
		Set(ByVal Value As String)
			TG03_CD_IBGE = Value
		End Set
	End Property
	Public Property SedeRegional() as String
		Get
			Return TG03_IN_SEDE_REGIONAL
		End Get
		Set(ByVal Value As String)
			TG03_IN_SEDE_REGIONAL = Value
		End Set
	End Property
	Public Property QuantidadePopulacao() as String
		Get
			Return TG03_QT_POPULACAO
		End Get
		Set(ByVal Value As String)
			TG03_QT_POPULACAO = Value
		End Set
	End Property
	Public Property DistanciaUre() as String
		Get
			Return TG03_NR_DISTANCIA_URE
		End Get
		Set(ByVal Value As String)
			TG03_NR_DISTANCIA_URE = Value
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
        strSQL.Append(" from DBGERAL.DBO.TG03_MUNICIPIO")
        strSQL.Append(" where TG03_ID_MUNICIPIO = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("TG02_ID_UF") = ProBanco(TG02_ID_UF, eTipoValor.CHAVE)
		dr("TG05_ID_REGIONAL") = ProBanco(TG05_ID_REGIONAL, eTipoValor.CHAVE)
		dr("TG05_ID_REGIONAL_ANTES") = ProBanco(TG05_ID_REGIONAL_ANTES, eTipoValor.CHAVE)
		dr("TG03_NM_MUNICIPIO") = ProBanco(TG03_NM_MUNICIPIO, eTipoValor.TEXTO)
		dr("TG03_CD_IBGE") = ProBanco(TG03_CD_IBGE, eTipoValor.TEXTO)
		dr("TG03_IN_SEDE_REGIONAL") = ProBanco(TG03_IN_SEDE_REGIONAL, eTipoValor.TEXTO)
		dr("TG03_QT_POPULACAO") = ProBanco(TG03_QT_POPULACAO, eTipoValor.NUMERO_INTEIRO)
		dr("TG03_NR_DISTANCIA_URE") = ProBanco(TG03_NR_DISTANCIA_URE, eTipoValor.NUMERO_INTEIRO)

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

        strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG03_MUNICIPIO")
        strSQL.Append(" where TG03_ID_MUNICIPIO = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG03_ID_MUNICIPIO = DoBanco(dr("TG03_ID_MUNICIPIO"), eTipoValor.CHAVE)
            TG02_ID_UF = DoBanco(dr("TG02_ID_UF"), eTipoValor.CHAVE)
            TG05_ID_REGIONAL = DoBanco(dr("TG05_ID_REGIONAL"), eTipoValor.CHAVE)
            TG05_ID_REGIONAL_ANTES = DoBanco(dr("TG05_ID_REGIONAL_ANTES"), eTipoValor.CHAVE)
            TG03_NM_MUNICIPIO = DoBanco(dr("TG03_NM_MUNICIPIO"), eTipoValor.TEXTO)
            TG03_CD_IBGE = DoBanco(dr("TG03_CD_IBGE"), eTipoValor.TEXTO)
            TG03_IN_SEDE_REGIONAL = DoBanco(dr("TG03_IN_SEDE_REGIONAL"), eTipoValor.TEXTO)
            TG03_QT_POPULACAO = DoBanco(dr("TG03_QT_POPULACAO"), eTipoValor.NUMERO_INTEIRO)
            TG03_NR_DISTANCIA_URE = DoBanco(dr("TG03_NR_DISTANCIA_URE"), eTipoValor.NUMERO_INTEIRO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional UF As Integer = 0, Optional Regional As Integer = 0, Optional RegionalAntes As Integer = 0, Optional Nome As String = "", Optional CodigoIBGE As String = "", Optional SedeRegional As String = "", Optional QuantidadePopulacao As String = "", Optional DistanciaUre As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG03_MUNICIPIO")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where TG03_ID_MUNICIPIO is not null")

        If Codigo > 0 Then
            strSQL.Append(" and TG03_ID_MUNICIPIO = " & Codigo)
        End If

        If UF > 0 Then
            strSQL.Append(" and TG02_ID_UF = " & UF)
        End If

        If Regional > 0 Then
            strSQL.Append(" and TG05_ID_REGIONAL = " & Regional)
        End If

        If RegionalAntes > 0 Then
            strSQL.Append(" and TG05_ID_REGIONAL_ANTES = " & RegionalAntes)
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(TG03_NM_MUNICIPIO) like '%" & Nome.toUpper & "%'")
        End If

        If CodigoIBGE <> "" Then
            strSQL.Append(" and upper(TG03_CD_IBGE) like '%" & CodigoIBGE.toUpper & "%'")
        End If

        If SedeRegional <> "" Then
            strSQL.Append(" and upper(TG03_IN_SEDE_REGIONAL) like '%" & SedeRegional.toUpper & "%'")
        End If

        If QuantidadePopulacao <> "" Then
            strSQL.Append(" and upper(TG03_QT_POPULACAO) like '%" & QuantidadePopulacao.toUpper & "%'")
        End If

        If DistanciaUre <> "" Then
            strSQL.Append(" and upper(TG03_NR_DISTANCIA_URE) like '%" & DistanciaUre.toUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "TG03_ID_MUNICIPIO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarParaCep(ByVal Nome As String) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select TG03_ID_MUNICIPIO, TG02.TG02_ID_UF, TG03_NM_MUNICIPIO, TG02_SG_UF  ")
        strSQL.Append(" from DBGERAL.DBO.TG03_MUNICIPIO as TG03")
        strSQL.Append(" left join DBGERAL.DBO.TG02_UF as TG02 on TG02.TG02_ID_UF = TG03.TG02_ID_UF")
        strSQL.Append(" where TG03_ID_MUNICIPIO is not null")
        strSQL.Append(" and TG03_NM_MUNICIPIO = '" & Nome.ToUpper & "'")

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela(Optional Regional As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select TG03_ID_MUNICIPIO as CODIGO, TG03_NM_MUNICIPIO as DESCRICAO ")
        strSQL.Append(" from DBGERAL.DBO.TG03_MUNICIPIO ")
        strSQL.Append(" where TG03_ID_MUNICIPIO is not null ")

        If Regional > 0 Then
            strSQL.Append(" and TG05_ID_REGIONAL = " & Regional)
        End If

        strSQL.Append(" order by 2 ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return dt
    End Function

    Public Function ObterTabelaEstado(ByVal Codigo as Integer,Optional SiglaEstado As String = "") As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select TG03_ID_MUNICIPIO as CODIGO, TG03_NM_MUNICIPIO as DESCRICAO ")
        strSQL.Append(" from DBGERAL.DBO.TG03_MUNICIPIO as TG03 ")
        strSQL.Append(" left join DBGERAL.DBO.TG02_UF as TG02 on TG02.TG02_ID_UF = TG03.TG02_ID_UF ")
        strSQL.Append(" where TG03_ID_MUNICIPIO is not null ")


        if Codigo > 0 Then
            strSQL.Append(" and TG02.TG02_ID_UF = " & Codigo )
        End If

        If SiglaEstado <> "" Then
            strSQL.Append(" and TG02.TG02_SG_UF = '" & SiglaEstado & "' ")
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

        strSQL.Append(" select max(TG03_ID_MUNICIPIO) from DBGERAL.DBO.TG03_MUNICIPIO")

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
        strSQL.Append(" from DBGERAL.DBO.TG03_MUNICIPIO")
        strSQL.Append(" where TG03_ID_MUNICIPIO = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class


