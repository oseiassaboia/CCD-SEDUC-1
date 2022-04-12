Imports Microsoft.VisualBasic
Imports System.Data

Public Class Cep
	Private TG55_ID_CEP as Integer
	Private TG55_NU_CEP as String

	Public Property Codigo() as Integer
		Get
			Return TG55_ID_CEP
		End Get
		Set(ByVal Value As Integer)
			TG55_ID_CEP = Value
		End Set
	End Property
	Public Property Numero() as String
		Get
			Return TG55_NU_CEP
		End Get
		Set(ByVal Value As String)
			TG55_NU_CEP = Value
		End Set
	End Property

	Public Sub New(Optional ByVal Codigo as Integer = 0)
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
        strSQL.Append(" from DBGERAL.DBO.TG55_CEP")
        strSQL.Append(" where TG55_ID_CEP = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("TG55_NU_CEP") = ProBanco(TG55_NU_CEP, eTipoValor.TEXTO)

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
        strSQL.Append(" from DBGERAL.DBO.TG55_CEP")
        strSQL.Append(" where TG55_ID_CEP = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG55_ID_CEP = DoBanco(dr("TG55_ID_CEP"), eTipoValor.CHAVE)
            TG55_NU_CEP = DoBanco(dr("TG55_NU_CEP"), eTipoValor.TEXTO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub ObterCep(ByVal Cep As String)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG55_CEP")
        strSQL.Append(" where TG55_NU_CEP = '" & Cep & "'")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG55_ID_CEP = DoBanco(dr("TG55_ID_CEP"), eTipoValor.CHAVE)
            TG55_NU_CEP = DoBanco(dr("TG55_NU_CEP"), eTipoValor.TEXTO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub


    Public Function Pesquisar(Optional ByVal Sort as String = "", Optional Codigo as Integer = 0, Optional Numero as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

        strSQL.Append(" select TG55.TG55_ID_CEP, TG55_NU_CEP, TG54_NM_LOGRADOURO, TG04_NM_BAIRRO, TG03_NM_MUNICIPIO, TG56.TG56_ID_LOGRADOURO_CEP ")
        strSQL.Append(" from DBGERAL.DBO.TG55_CEP As TG55")
        strSQL.Append(" LEFT JOIN DBGERAL.DBO.TG56_LOGRADOURO_CEP As TG56 On TG56.TG55_ID_CEP = TG55.TG55_ID_CEP ")
        strSQL.Append(" LEFT JOIN DBGERAL.DBO.TG54_LOGRADOURO As TG54 On TG54.TG54_ID_LOGRADOURO = TG56.TG54_ID_LOGRADOURO ")
        strSQL.Append(" LEFT JOIN DBGERAL.DBO.TG16_TIPO_LOGRADOURO As TG16 On TG16.TG16_ID_TIPO_LOGRADOURO = TG54.TG16_ID_TIPO_LOGRADOURO ")
        strSQL.Append(" LEFT JOIN DBGERAL.DBO.TG04_BAIRRO As TG04 On TG04.TG04_ID_BAIRRO = TG54.TG04_ID_BAIRRO  ")
        strSQL.Append(" LEFT JOIN DBGERAL.DBO.TG03_MUNICIPIO As TG03 On TG03.TG03_ID_MUNICIPIO = TG04.TG03_ID_MUNICIPIO ")
        strSQL.Append(" where TG55.TG55_ID_CEP is not null")

        If Codigo > 0 Then
            strSQL.Append(" and TG55.TG55_ID_CEP = " & Codigo)
        End If
		
		If Numero <> "" Then
            strSQL.Append(" and TG55_NU_CEP = '" & Numero & "'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "TG55.TG55_ID_CEP", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select TG55_ID_CEP as CODIGO, TG55_NU_CEP as DESCRICAO")
        strSQL.Append(" from DBGERAL.DBO.TG55_CEP")
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

        strSQL.Append(" select max(TG55_ID_CEP) from DBGERAL.DBO.TG55_CEP")

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
        strSQL.Append(" from DBGERAL.DBO.TG55_CEP")
        strSQL.Append(" where TG55_ID_CEP = " & Codigo)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class