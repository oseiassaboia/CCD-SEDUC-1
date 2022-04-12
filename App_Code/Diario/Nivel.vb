Imports Microsoft.VisualBasic
Imports System.Data

Public Class Nivel
	Private DE06_ID_NIVEL as Integer
	Private DE05_ID_MODALIDADE as Integer
    Private DE06_NM_NIVEL As String
    Private DE06_IN_PROFISSIONALIZANTE As String

    Public Property Codigo() as Integer
		Get
			Return DE06_ID_NIVEL
		End Get
		Set(ByVal Value As Integer)
			DE06_ID_NIVEL = Value
		End Set
	End Property
	Public Property Modalidade() as Integer
		Get
			Return DE05_ID_MODALIDADE
		End Get
		Set(ByVal Value As Integer)
			DE05_ID_MODALIDADE = Value
		End Set
	End Property
    Public Property Nome() As String
        Get
            Return DE06_NM_NIVEL
        End Get
        Set(ByVal Value As String)
            DE06_NM_NIVEL = Value
        End Set
    End Property
    Public Property Profissionalizante() As String
        Get
            Return DE06_IN_PROFISSIONALIZANTE
        End Get
        Set(ByVal Value As String)
            DE06_IN_PROFISSIONALIZANTE = Value
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
        strSQL.Append(" from DBDIARIO..DE06_NIVEL")
        strSQL.Append(" where DE06_ID_NIVEL = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("DE05_ID_MODALIDADE") = ProBanco(DE05_ID_MODALIDADE, eTipoValor.CHAVE)
        dr("DE06_NM_NIVEL") = ProBanco(DE06_NM_NIVEL, eTipoValor.TEXTO)
        dr("DE06_IN_PROFISSIONALIZANTE") = ProBanco(DE06_IN_PROFISSIONALIZANTE, eTipoValor.BOOLEANO)

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
        strSQL.Append(" from DBDIARIO..DE06_NIVEL")
        strSQL.Append(" where DE06_ID_NIVEL = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            DE06_ID_NIVEL = DoBanco(dr("DE06_ID_NIVEL"), eTipoValor.CHAVE)
            DE05_ID_MODALIDADE = DoBanco(dr("DE05_ID_MODALIDADE"), eTipoValor.CHAVE)
            DE06_NM_NIVEL = DoBanco(dr("DE06_NM_NIVEL"), eTipoValor.TEXTO)
            DE06_IN_PROFISSIONALIZANTE = DoBanco(dr("DE06_IN_PROFISSIONALIZANTE"), eTipoValor.BOOLEANO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional Modalidade As Integer = 0, Optional Nome As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select DE06.DE06_ID_NIVEL, DE06.DE05_ID_MODALIDADE, DE06.DE06_NM_NIVEL  ")
        strSQL.Append(" , DE05.DE05_NM_MODALIDADE")
        strSQL.Append(" from DBDIARIO..DE06_NIVEL as DE06")
        strSQL.Append(" left join DE05_MODALIDADE as DE05 on DE05.DE05_ID_MODALIDADE = DE06.DE05_ID_MODALIDADE ")
        strSQL.Append(" where DE06_ID_NIVEL is not null")

        If Codigo > 0 Then
            strSQL.Append(" and DE06.DE06_ID_NIVEL = " & Codigo)
        End If

        If Modalidade > 0 Then
            strSQL.Append(" and DE06.DE05_ID_MODALIDADE = " & Modalidade)
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(DE06.DE06_NM_NIVEL) like '%" & Nome.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "DE06.DE06_ID_NIVEL", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarNomeNivel(ByVal Modalidade As Integer, ByVal NomeNivel As String) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhaEncontrada As Integer

        strSQL.Append(" select 1 as QTD where exists ( ")
        strSQL.Append("     select DE06_NM_NIVEL ")
        strSQL.Append("     from DBDIARIO..DE06_NIVEL ")
        strSQL.Append("     where DE06_ID_NIVEL is not null")
        strSQL.Append("     and DE05_ID_MODALIDADE = " & Modalidade)
        strSQL.Append("     and DE06_NM_NIVEL = '" & NomeNivel.ToUpper & "'")
        strSQL.Append(" )")

        LinhaEncontrada = cnn.AbrirDataTable(strSQL.ToString).Rows.Count

        Return LinhaEncontrada
    End Function

    Public Function ObterTabela(Optional Modalidade As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select DE06_ID_NIVEL as CODIGO, DE06_NM_NIVEL as DESCRICAO")
        strSQL.Append(" from DBDIARIO..DE06_NIVEL")
        strSQL.Append(" where DE06_ID_NIVEL is not null ")

        If Modalidade > 0 Then
            strSQL.Append(" and DE05_ID_MODALIDADE = " & Modalidade)
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

        strSQL.Append(" select max(DE06_ID_NIVEL) from DBDIARIO..DE06_NIVEL")

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
        strSQL.Append(" from DE06_NIVEL")
        strSQL.Append(" where DE06_ID_NIVEL = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class


