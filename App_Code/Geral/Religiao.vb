Imports Microsoft.VisualBasic
Imports System.Data

Public Class Religiao
	Private TG10_ID_RELIGIAO as Integer
	Private TG10_NM_RELIGIAO as String

	Public Property Codigo() as Integer
		Get
			Return TG10_ID_RELIGIAO
		End Get
		Set(ByVal Value As Integer)
			TG10_ID_RELIGIAO = Value
		End Set
	End Property
	Public Property Nome() as String
		Get
			Return TG10_NM_RELIGIAO
		End Get
		Set(ByVal Value As String)
			TG10_NM_RELIGIAO = Value
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
        strSQL.Append(" from DBGERAL.DBO.TG10_RELIGIAO")
        strSQL.Append(" where TG10_ID_RELIGIAO = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("TG10_NM_RELIGIAO") = ProBanco(TG10_NM_RELIGIAO, eTipoValor.TEXTO)

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
        strSQL.Append(" from DBGERAL.DBO.TG10_RELIGIAO")
        strSQL.Append(" where TG10_ID_RELIGIAO = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG10_ID_RELIGIAO = DoBanco(dr("TG10_ID_RELIGIAO"), eTipoValor.CHAVE)
            TG10_NM_RELIGIAO = DoBanco(dr("TG10_NM_RELIGIAO"), eTipoValor.TEXTO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort as String = "", Optional Codigo as Integer = 0, Optional Nome as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG10_RELIGIAO")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where TG10_ID_RELIGIAO is not null")
		
		If Codigo > 0 then 
			strSQL.Append(" and TG10_ID_RELIGIAO = " & Codigo)
		End If
		
		If Nome <> "" then 
			strSQL.Append(" and upper(TG10_NM_RELIGIAO) like '%" & Nome.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "TG10_ID_RELIGIAO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select TG10_ID_RELIGIAO as CODIGO, TG10_NM_RELIGIAO as DESCRICAO")
        strSQL.Append(" from DBGERAL.DBO.TG10_RELIGIAO")
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

        strSQL.Append(" select max(TG10_ID_RELIGIAO) from DBGERAL.DBO.TG10_RELIGIAO")

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
        strSQL.Append(" from DBGERAL.DBO.TG10_RELIGIAO")
        strSQL.Append(" where TG10_ID_RELIGIAO = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class


