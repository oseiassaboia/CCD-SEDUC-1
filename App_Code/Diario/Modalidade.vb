Imports Microsoft.VisualBasic
Imports System.Data

Public Class Modalidade
	Private DE05_ID_MODALIDADE as Integer
	Private DE30_ID_TIPO_ATENDIMENTO as Integer
	Private DE05_NM_MODALIDADE as String

	Public Property Codigo() as Integer
		Get
			Return DE05_ID_MODALIDADE
		End Get
		Set(ByVal Value As Integer)
			DE05_ID_MODALIDADE = Value
		End Set
	End Property
	Public Property TipoAtendimento() as Integer
		Get
			Return DE30_ID_TIPO_ATENDIMENTO
		End Get
		Set(ByVal Value As Integer)
			DE30_ID_TIPO_ATENDIMENTO = Value
		End Set
	End Property
	Public Property Nome() as String
		Get
			Return DE05_NM_MODALIDADE
		End Get
		Set(ByVal Value As String)
			DE05_NM_MODALIDADE = Value
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
        strSQL.Append(" from DBDIARIO..DE05_MODALIDADE")
        strSQL.Append(" where DE05_ID_MODALIDADE = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("DE30_ID_TIPO_ATENDIMENTO") = ProBanco(DE30_ID_TIPO_ATENDIMENTO, eTipoValor.CHAVE)
		dr("DE05_NM_MODALIDADE") = ProBanco(DE05_NM_MODALIDADE, eTipoValor.TEXTO)

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
        strSQL.Append(" from DBDIARIO..DE05_MODALIDADE")
        strSQL.Append(" where DE05_ID_MODALIDADE = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            DE05_ID_MODALIDADE = DoBanco(dr("DE05_ID_MODALIDADE"), eTipoValor.CHAVE)
            DE30_ID_TIPO_ATENDIMENTO = DoBanco(dr("DE30_ID_TIPO_ATENDIMENTO"), eTipoValor.CHAVE)
            DE05_NM_MODALIDADE = DoBanco(dr("DE05_NM_MODALIDADE"), eTipoValor.TEXTO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort as String = "", Optional Codigo as Integer = 0, Optional TipoAtendimento as Integer = 0, Optional Nome as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

        strSQL.Append(" select DE05.DE05_ID_MODALIDADE, DE05.DE05_NM_MODALIDADE, DE05.DE30_ID_TIPO_ATENDIMENTO ")
        strSQL.Append(" , DE30.DE30_NM_TIPO_ATENDIMENTO")
        strSQL.Append(" from DBDIARIO..DE05_MODALIDADE as DE05")
        strSQL.Append(" left join DE30_TIPO_ATENDIMENTO AS DE30 on DE30.DE30_ID_TIPO_ATENDIMENTO = DE05.DE30_ID_TIPO_ATENDIMENTO ")
        strSQL.Append(" where DE05.DE05_ID_MODALIDADE is not null")

        If Codigo > 0 Then
            strSQL.Append(" and DE05.DE05_ID_MODALIDADE = " & Codigo)
        End If
		
		If TipoAtendimento > 0 Then
            strSQL.Append(" and DE05.DE30_ID_TIPO_ATENDIMENTO = " & TipoAtendimento)
        End If
		
		If Nome <> "" Then
            strSQL.Append(" and upper(DE05.DE05_NM_MODALIDADE) like '%" & Nome.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "DE05.DE05_ID_MODALIDADE", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder

        strSQL.Append(" select DE05_ID_MODALIDADE as CODIGO, DE05_NM_MODALIDADE as DESCRICAO")
        strSQL.Append(" from DBDIARIO..DE05_MODALIDADE")
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

        strSQL.Append(" select max(DE05_ID_MODALIDADE) from DBDIARIO..DE05_MODALIDADE")

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
        strSQL.Append(" from DBDIARIO..DE05_MODALIDADE")
        strSQL.Append(" where DE05_ID_MODALIDADE = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class


