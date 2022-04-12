Imports Microsoft.VisualBasic
Imports System.Data

Public Class OrgaoEmissor
	Private TG43_ID_ORGAO_EMISSOR as Integer
	Private TG43_NM_ORGAO_EMISSOR as String
	Private TG43_NM_ORGAO_EMISSOR_ABREV as String
	Private TG43_IN_MATRICULA as String

	Public Property Codigo() as Integer
		Get
			Return TG43_ID_ORGAO_EMISSOR
		End Get
		Set(ByVal Value As Integer)
			TG43_ID_ORGAO_EMISSOR = Value
		End Set
	End Property
	Public Property Nome() as String
		Get
			Return TG43_NM_ORGAO_EMISSOR
		End Get
		Set(ByVal Value As String)
			TG43_NM_ORGAO_EMISSOR = Value
		End Set
	End Property
    Public Property NomeAbreviado() As String
        Get
            Return TG43_NM_ORGAO_EMISSOR_ABREV
        End Get
        Set(ByVal Value As String)
            TG43_NM_ORGAO_EMISSOR_ABREV = Value
        End Set
    End Property
    Public Property Matricula() as String
		Get
			Return TG43_IN_MATRICULA
		End Get
		Set(ByVal Value As String)
			TG43_IN_MATRICULA = Value
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
        strSQL.Append(" from DBGERAL.DBO.TG43_ORGAO_EMISSOR")
        strSQL.Append(" where TG43_ID_ORGAO_EMISSOR = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("TG43_NM_ORGAO_EMISSOR") = ProBanco(TG43_NM_ORGAO_EMISSOR, eTipoValor.TEXTO)
		dr("TG43_NM_ORGAO_EMISSOR_ABREV") = ProBanco(TG43_NM_ORGAO_EMISSOR_ABREV, eTipoValor.TEXTO)
		dr("TG43_IN_MATRICULA") = ProBanco(TG43_IN_MATRICULA, eTipoValor.TEXTO)

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
        strSQL.Append(" from DBGERAL.DBO.TG43_ORGAO_EMISSOR")
        strSQL.Append(" where TG43_ID_ORGAO_EMISSOR = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG43_ID_ORGAO_EMISSOR = DoBanco(dr("TG43_ID_ORGAO_EMISSOR"), eTipoValor.CHAVE)
            TG43_NM_ORGAO_EMISSOR = DoBanco(dr("TG43_NM_ORGAO_EMISSOR"), eTipoValor.TEXTO)
            TG43_NM_ORGAO_EMISSOR_ABREV = DoBanco(dr("TG43_NM_ORGAO_EMISSOR_ABREV"), eTipoValor.TEXTO)
            TG43_IN_MATRICULA = DoBanco(dr("TG43_IN_MATRICULA"), eTipoValor.TEXTO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort as String = "", Optional Codigo as Integer = 0, Optional Nome as String = "", Optional NomeAbreviado as String = "", Optional Matricula as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG43_ORGAO_EMISSOR")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where TG43_ID_ORGAO_EMISSOR is not null")
		
		If Codigo > 0 then 
			strSQL.Append(" and TG43_ID_ORGAO_EMISSOR = " & Codigo)
		End If
		
		If Nome <> "" then 
			strSQL.Append(" and upper(TG43_NM_ORGAO_EMISSOR) like '%" & Nome.toUpper & "%'")
		End If
		
		If NomeAbreviado <> "" then 
			strSQL.Append(" and upper(TG43_NM_ORGAO_EMISSOR_ABREV) like '%" & NomeAbreviado.toUpper & "%'")
		End If
		
		If Matricula <> "" then 
			strSQL.Append(" and upper(TG43_IN_MATRICULA) like '%" & Matricula.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "TG43_ID_ORGAO_EMISSOR", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select TG43_ID_ORGAO_EMISSOR as CODIGO, TG43_NM_ORGAO_EMISSOR as DESCRICAO")
        strSQL.Append(" from DBGERAL.DBO.TG43_ORGAO_EMISSOR")
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

        strSQL.Append(" select max(TG43_ID_ORGAO_EMISSOR) from DBGERAL.DBO.TG43_ORGAO_EMISSOR")

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
        strSQL.Append(" from DBGERAL.DBO.TG43_ORGAO_EMISSOR")
        strSQL.Append(" where TG43_ID_ORGAO_EMISSOR = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class


