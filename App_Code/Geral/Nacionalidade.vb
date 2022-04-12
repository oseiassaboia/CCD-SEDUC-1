Imports Microsoft.VisualBasic
Imports System.Data

Public Class Nacionalidade
	Private TG13_ID_NACIONALIDADE as Integer
	Private TG13_NM_NACIONALIDADE as String
	Private TG13_IN_CENSO as String

	Public Property Codigo() as Integer
		Get
			Return TG13_ID_NACIONALIDADE
		End Get
		Set(ByVal Value As Integer)
			TG13_ID_NACIONALIDADE = Value
		End Set
	End Property
    Public Property Nome() As String
        Get
            Return TG13_NM_NACIONALIDADE
        End Get
        Set(ByVal Value As String)
            TG13_NM_NACIONALIDADE = Value
        End Set
    End Property
    Public Property Censo() as String
		Get
			Return TG13_IN_CENSO
		End Get
		Set(ByVal Value As String)
			TG13_IN_CENSO = Value
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
        strSQL.Append(" from DBGERAL.DBO.TG13_NACIONALIDADE")
        strSQL.Append(" where TG13_ID_NACIONALIDADE = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("TG13_NM_NACIONALIDADE") = ProBanco(TG13_NM_NACIONALIDADE, eTipoValor.TEXTO)
		dr("TG13_IN_CENSO") = ProBanco(TG13_IN_CENSO, eTipoValor.TEXTO)

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
        strSQL.Append(" from DBGERAL.DBO.TG13_NACIONALIDADE")
        strSQL.Append(" where TG13_ID_NACIONALIDADE = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG13_ID_NACIONALIDADE = DoBanco(dr("TG13_ID_NACIONALIDADE"), eTipoValor.CHAVE)
            TG13_NM_NACIONALIDADE = DoBanco(dr("TG13_NM_NACIONALIDADE"), eTipoValor.TEXTO)
            TG13_IN_CENSO = DoBanco(dr("TG13_IN_CENSO"), eTipoValor.TEXTO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort as String = "", Optional Codigo as Integer = 0, Optional Nome as String = "", Optional Censo as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG13_NACIONALIDADE")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where TG13_ID_NACIONALIDADE is not null")
		
		If Codigo > 0 then 
			strSQL.Append(" and TG13_ID_NACIONALIDADE = " & Codigo)
		End If
		
		If Nome <> "" then 
			strSQL.Append(" and upper(TG13_NM_NACIONALIDADE) like '%" & Nome.toUpper & "%'")
		End If
		
		If Censo <> "" then 
			strSQL.Append(" and upper(TG13_IN_CENSO) like '%" & Censo.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "TG13_ID_NACIONALIDADE", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select TG13_ID_NACIONALIDADE as CODIGO, TG13_NM_NACIONALIDADE as DESCRICAO")
        strSQL.Append(" from DBGERAL.DBO.TG13_NACIONALIDADE")
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

        strSQL.Append(" select max(TG13_ID_NACIONALIDADE) from DBGERAL.DBO.TG13_NACIONALIDADE")

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
        strSQL.Append(" from DBGERAL.DBO.TG13_NACIONALIDADE")
        strSQL.Append(" where TG13_ID_NACIONALIDADE = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class


