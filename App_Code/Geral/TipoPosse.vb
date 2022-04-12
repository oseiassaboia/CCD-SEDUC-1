Imports Microsoft.VisualBasic
Imports System.Data

Public Class TipoPosse
	Private TG58_ID_TIPO_POSSE as Integer
	Private TG58_NM_TIPO_POSSE as String

	Public Property idTipoPosse() as Integer
		Get
			Return TG58_ID_TIPO_POSSE
		End Get
		Set(ByVal Value As Integer)
			TG58_ID_TIPO_POSSE = Value
		End Set
	End Property
	Public Property Descricao() as String
		Get
			Return TG58_NM_TIPO_POSSE
		End Get
		Set(ByVal Value As String)
			TG58_NM_TIPO_POSSE = Value
		End Set
	End Property

    Public Sub New(Optional ByVal idTipoPosse As Integer = 0)
        If idTipoPosse > 0 Then
            Obter(idTipoPosse)
        End If
    End Sub

	Public Sub Salvar()
        Dim cnn As New ConexaoGeral
        Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append("Then select * ")
		strSQL.Append(" from TG58_TIPO_POSSE")
		strSQL.Append(" where TG58_ID_TIPO_POSSE = " & idTipoPosse)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("TG58_NM_TIPO_POSSE") = ProBanco(TG58_NM_TIPO_POSSE, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

        cnn = Nothing
	End Sub

	Public Sub Obter(ByVal idTipoPosse as String)
        Dim cnn As New ConexaoGeral
        Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from TG58_TIPO_POSSE")
		strSQL.Append(" where TG58_ID_TIPO_POSSE = " & idTipoPosse)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			TG58_ID_TIPO_POSSE = DoBanco(dr("TG58_ID_TIPO_POSSE"), eTipoValor.CHAVE)
			TG58_NM_TIPO_POSSE = DoBanco(dr("TG58_NM_TIPO_POSSE"), eTipoValor.TEXTO)
		End If

        cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional idTipoPosse as Integer = 0, Optional Descricao as String = "") as DataTable
        Dim cnn As New ConexaoGeral
        Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from TG58_TIPO_POSSE")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where TG58_ID_TIPO_POSSE is not null")
		
		If idTipoPosse > 0 then 
			strSQL.Append(" and TG58_ID_TIPO_POSSE = " & idTipoPosse)
		End If
		
		If Descricao <> "" then 
			strSQL.Append(" and upper(TG58_NM_TIPO_POSSE) like '%" & Descricao.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "TG58_ID_TIPO_POSSE", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
        Dim cnn As New ConexaoGeral
        Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select TG58_ID_TIPO_POSSE as CODIGO, TG58_NM_TIPO_POSSE as DESCRICAO")
		strSQL.Append(" from TG58_TIPO_POSSE")
		strSQL.Append(" order by 2 ")

		dt = cnn.AbrirDataTable(strSQL.ToString)


        cnn = Nothing

		Return dt
	End Function

	Public Function ObterUltimo() as Integer
        Dim cnn As New ConexaoGeral
        Dim strSQL As New StringBuilder
		Dim CodigoUltimo As Integer
		
		strSQL.Append(" select max(TG58_ID_TIPO_POSSE) from TG58_TIPO_POSSE")

		With cnn.AbrirDataTable(strSQL.ToString)
			If Not IsDBNull(.Rows(0)(0)) Then
				CodigoUltimo = .Rows(0)(0)
			Else
				CodigoUltimo = 0
			End If
		End With


        cnn = Nothing

		Return CodigoUltimo

	End Function
	Public Function Excluir(ByVal idTipoPosse as String) As Integer
        Dim cnn As New ConexaoGeral
        Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from TG58_TIPO_POSSE")
		strSQL.Append(" where TG58_ID_TIPO_POSSE = " & idTipoPosse)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)


        cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class