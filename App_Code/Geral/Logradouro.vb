Imports Microsoft.VisualBasic
Imports System.Data

Public Class Logradouro
	Private TG54_ID_LOGRADOURO as Integer
	Private TG04_ID_BAIRRO as Integer
	Private TG16_ID_TIPO_LOGRADOURO as Integer
	Private TG54_NM_LOGRADOURO as String
    Private TG54_IN_CADASTRO_CORREIOS As Boolean

    Public Property Codigo() as Integer
		Get
			Return TG54_ID_LOGRADOURO
		End Get
		Set(ByVal Value As Integer)
			TG54_ID_LOGRADOURO = Value
		End Set
	End Property
	Public Property Bairro() as Integer
		Get
			Return TG04_ID_BAIRRO
		End Get
		Set(ByVal Value As Integer)
			TG04_ID_BAIRRO = Value
		End Set
	End Property
	Public Property TipoLogradouro() as Integer
		Get
			Return TG16_ID_TIPO_LOGRADOURO
		End Get
		Set(ByVal Value As Integer)
			TG16_ID_TIPO_LOGRADOURO = Value
		End Set
	End Property
	Public Property Nome() as String
		Get
			Return TG54_NM_LOGRADOURO
		End Get
		Set(ByVal Value As String)
			TG54_NM_LOGRADOURO = Value
		End Set
	End Property
    Public Property CadastroCorreios() As Boolean
        Get
            Return TG54_IN_CADASTRO_CORREIOS
        End Get
        Set(ByVal Value As Boolean)
            TG54_IN_CADASTRO_CORREIOS = Value
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
        strSQL.Append(" from DBGERAL.DBO.TG54_LOGRADOURO")
        strSQL.Append(" where TG54_ID_LOGRADOURO = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("TG04_ID_BAIRRO") = ProBanco(TG04_ID_BAIRRO, eTipoValor.CHAVE)
		dr("TG16_ID_TIPO_LOGRADOURO") = ProBanco(TG16_ID_TIPO_LOGRADOURO, eTipoValor.CHAVE)
		dr("TG54_NM_LOGRADOURO") = ProBanco(TG54_NM_LOGRADOURO, eTipoValor.TEXTO)
        dr("TG54_IN_CADASTRO_CORREIOS") = ProBanco(TG54_IN_CADASTRO_CORREIOS, eTipoValor.BOOLEANO)

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
        strSQL.Append(" from DBGERAL.DBO.TG54_LOGRADOURO")
        strSQL.Append(" where TG54_ID_LOGRADOURO = " & Codigo)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			TG54_ID_LOGRADOURO = DoBanco(dr("TG54_ID_LOGRADOURO"), eTipoValor.CHAVE)
			TG04_ID_BAIRRO = DoBanco(dr("TG04_ID_BAIRRO"), eTipoValor.CHAVE)
			TG16_ID_TIPO_LOGRADOURO = DoBanco(dr("TG16_ID_TIPO_LOGRADOURO"), eTipoValor.CHAVE)
			TG54_NM_LOGRADOURO = DoBanco(dr("TG54_NM_LOGRADOURO"), eTipoValor.TEXTO)
            TG54_IN_CADASTRO_CORREIOS = DoBanco(dr("TG54_IN_CADASTRO_CORREIOS"), eTipoValor.BOOLEANO)
        End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional Codigo as Integer = 0, Optional Bairro as Integer = 0, Optional TipoLogradouro as Integer = 0, Optional Nome as String = "", Optional CadastroCorreios as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG54_LOGRADOURO")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where TG54_ID_LOGRADOURO is not null")
		
		If Codigo > 0 then 
			strSQL.Append(" and TG54_ID_LOGRADOURO = " & Codigo)
		End If
		
		If Bairro > 0 then 
			strSQL.Append(" and TG04_ID_BAIRRO = " & Bairro)
		End If
		
		If TipoLogradouro > 0 then 
			strSQL.Append(" and TG16_ID_TIPO_LOGRADOURO = " & TipoLogradouro)
		End If
		
		If Nome <> "" then 
			strSQL.Append(" and upper(TG54_NM_LOGRADOURO) like '%" & Nome.toUpper & "%'")
		End If
		
		If CadastroCorreios <> "" then 
			strSQL.Append(" and upper(TG54_IN_CADASTRO_CORREIOS) like '%" & CadastroCorreios.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "TG54_ID_LOGRADOURO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select TG54_ID_LOGRADOURO as CODIGO, TG16_ID_TIPO_LOGRADOURO as DESCRICAO")
        strSQL.Append(" from DBGERAL.DBO.TG54_LOGRADOURO")
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

        strSQL.Append(" select max(TG54_ID_LOGRADOURO) from DBGERAL.DBO.TG54_LOGRADOURO")

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
        strSQL.Append(" from DBGERAL.DBO.TG54_LOGRADOURO")
        strSQL.Append(" where TG54_ID_LOGRADOURO = " & Codigo)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class

