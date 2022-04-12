Imports Microsoft.VisualBasic
Imports System.Data

Public Class UF
	Private TG02_ID_UF as Integer
	Private TG01_ID_PAIS as Integer
	Private TG02_ID_UF_CENSO as String
	Private TG02_NM_UF as String
	Private TG02_SG_UF as String

	Public Property Codigo() as Integer
		Get
			Return TG02_ID_UF
		End Get
		Set(ByVal Value As Integer)
			TG02_ID_UF = Value
		End Set
	End Property
	Public Property Pais() as Integer
		Get
			Return TG01_ID_PAIS
		End Get
		Set(ByVal Value As Integer)
			TG01_ID_PAIS = Value
		End Set
	End Property
	Public Property CodigoUfCenso() as String
		Get
            Return TG02_ID_UF_CENSO
        End Get
		Set(ByVal Value As String)
			TG02_ID_UF_CENSO = Value
		End Set
	End Property
	Public Property Nome() as String
		Get
			Return TG02_NM_UF
		End Get
		Set(ByVal Value As String)
			TG02_NM_UF = Value
		End Set
	End Property
	Public Property Sigla() as String
		Get
			Return TG02_SG_UF
		End Get
		Set(ByVal Value As String)
			TG02_SG_UF = Value
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
        strSQL.Append(" from DBGERAL.DBO.TG02_UF")
        strSQL.Append(" where TG02_ID_UF = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("TG01_ID_PAIS") = ProBanco(TG01_ID_PAIS, eTipoValor.CHAVE)
		dr("TG02_ID_UF_CENSO") = ProBanco(TG02_ID_UF_CENSO, eTipoValor.NUMERO_INTEIRO)
		dr("TG02_NM_UF") = ProBanco(TG02_NM_UF, eTipoValor.TEXTO)
		dr("TG02_SG_UF") = ProBanco(TG02_SG_UF, eTipoValor.TEXTO)

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
        strSQL.Append(" from DBGERAL.DBO.TG02_UF")
        strSQL.Append(" where TG02_ID_UF = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG02_ID_UF = DoBanco(dr("TG02_ID_UF"), eTipoValor.CHAVE)
            TG01_ID_PAIS = DoBanco(dr("TG01_ID_PAIS"), eTipoValor.CHAVE)
            TG02_ID_UF_CENSO = DoBanco(dr("TG02_ID_UF_CENSO"), eTipoValor.NUMERO_INTEIRO)
            TG02_NM_UF = DoBanco(dr("TG02_NM_UF"), eTipoValor.TEXTO)
            TG02_SG_UF = DoBanco(dr("TG02_SG_UF"), eTipoValor.TEXTO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort as String = "", Optional Codigo as Integer = 0, Optional Pais as Integer = 0, Optional CodigoUfCenso as String = "", Optional Nome as String = "", Optional Sigla as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG02_UF")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where TG02_ID_UF is not null")
		
		If Codigo > 0 then 
			strSQL.Append(" and TG02_ID_UF = " & Codigo)
		End If
		
		If Pais > 0 then 
			strSQL.Append(" and TG01_ID_PAIS = " & Pais)
		End If
		
		If CodigoUfCenso <> "" then 
			strSQL.Append(" and upper(TG02_ID_UF_CENSO) like '%" & CodigoUfCenso.toUpper & "%'")
		End If
		
		If Nome <> "" then 
			strSQL.Append(" and upper(TG02_NM_UF) like '%" & Nome.toUpper & "%'")
		End If
		
		If Sigla <> "" then 
			strSQL.Append(" and upper(TG02_SG_UF) like '%" & Sigla.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "TG02_ID_UF", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder

        strSQL.Append(" select TG02_ID_UF as CODIGO, TG02_NM_UF as DESCRICAO")
        strSQL.Append(" from DBGERAL.DBO.TG02_UF")
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

        strSQL.Append(" select max(TG02_ID_UF) from DBGERAL.DBO.TG02_UF")

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
        strSQL.Append(" from DBGERAL.DBO.TG02_UF")
        strSQL.Append(" where TG02_ID_UF = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class


