Imports Microsoft.VisualBasic
Imports System.Data

Public Class LogradouroCep
	Private TG56_ID_LOGRADOURO_CEP as Integer
	Private TG54_ID_LOGRADOURO as Integer
	Private TG55_ID_CEP as Integer
    Private TG56_IN_CADASTRO_CORREIOS As Boolean

    Public Property Codigo() as Integer
		Get
			Return TG56_ID_LOGRADOURO_CEP
		End Get
		Set(ByVal Value As Integer)
			TG56_ID_LOGRADOURO_CEP = Value
		End Set
	End Property
	Public Property Logradouro() as Integer
		Get
			Return TG54_ID_LOGRADOURO
		End Get
		Set(ByVal Value As Integer)
			TG54_ID_LOGRADOURO = Value
		End Set
	End Property
	Public Property Cep() as Integer
		Get
			Return TG55_ID_CEP
		End Get
		Set(ByVal Value As Integer)
			TG55_ID_CEP = Value
		End Set
	End Property
    Public Property CadastroCorreios() As Boolean
        Get
            Return TG56_IN_CADASTRO_CORREIOS
        End Get
        Set(ByVal Value As Boolean)
            TG56_IN_CADASTRO_CORREIOS = Value
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
        strSQL.Append(" from DBGERAL.DBO.TG56_LOGRADOURO_CEP")
        strSQL.Append(" where TG56_ID_LOGRADOURO_CEP = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("TG54_ID_LOGRADOURO") = ProBanco(TG54_ID_LOGRADOURO, eTipoValor.CHAVE)
		dr("TG55_ID_CEP") = ProBanco(TG55_ID_CEP, eTipoValor.CHAVE)
        dr("TG56_IN_CADASTRO_CORREIOS") = ProBanco(TG56_IN_CADASTRO_CORREIOS, eTipoValor.BOOLEANO)

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
        strSQL.Append(" from DBGERAL.DBO.TG56_LOGRADOURO_CEP")
        strSQL.Append(" where TG56_ID_LOGRADOURO_CEP = " & Codigo)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			TG56_ID_LOGRADOURO_CEP = DoBanco(dr("TG56_ID_LOGRADOURO_CEP"), eTipoValor.CHAVE)
			TG54_ID_LOGRADOURO = DoBanco(dr("TG54_ID_LOGRADOURO"), eTipoValor.CHAVE)
			TG55_ID_CEP = DoBanco(dr("TG55_ID_CEP"), eTipoValor.CHAVE)
            TG56_IN_CADASTRO_CORREIOS = DoBanco(dr("TG56_IN_CADASTRO_CORREIOS"), eTipoValor.BOOLEANO)
        End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub


    Public Sub ObterCep(ByVal Codigo As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG56_LOGRADOURO_CEP")
        strSQL.Append(" where TG55_ID_CEP = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG56_ID_LOGRADOURO_CEP = DoBanco(dr("TG56_ID_LOGRADOURO_CEP"), eTipoValor.CHAVE)
            TG54_ID_LOGRADOURO = DoBanco(dr("TG54_ID_LOGRADOURO"), eTipoValor.CHAVE)
            TG55_ID_CEP = DoBanco(dr("TG55_ID_CEP"), eTipoValor.CHAVE)
            TG56_IN_CADASTRO_CORREIOS = DoBanco(dr("TG56_IN_CADASTRO_CORREIOS"), eTipoValor.BOOLEANO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub
    Public Function Pesquisar(Optional ByVal Sort as String = "", Optional Codigo as Integer = 0, Optional Logradouro as Integer = 0, Optional Cep as Integer = 0, Optional CadastroCorreios as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG56_LOGRADOURO_CEP")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where TG56_ID_LOGRADOURO_CEP is not null")
		
		If Codigo > 0 then 
			strSQL.Append(" and TG56_ID_LOGRADOURO_CEP = " & Codigo)
		End If
		
		If Logradouro > 0 then 
			strSQL.Append(" and TG54_ID_LOGRADOURO = " & Logradouro)
		End If
		
		If Cep > 0 then 
			strSQL.Append(" and TG55_ID_CEP = " & Cep)
		End If
		
		If CadastroCorreios <> "" then 
			strSQL.Append(" and upper(TG56_IN_CADASTRO_CORREIOS) like '%" & CadastroCorreios.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "TG56_ID_LOGRADOURO_CEP", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select TG56_ID_LOGRADOURO_CEP as CODIGO, ")
        strSQL.Append(" from DBGERAL.DBO.TG56_LOGRADOURO_CEP")
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

        strSQL.Append(" select max(TG56_ID_LOGRADOURO_CEP) from DBGERAL.DBO.TG56_LOGRADOURO_CEP")

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
        strSQL.Append(" from DBGERAL.DBO.TG56_LOGRADOURO_CEP")
        strSQL.Append(" where TG56_ID_LOGRADOURO_CEP = " & Codigo)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class

