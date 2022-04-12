Imports Microsoft.VisualBasic
Imports System.Data

Public Class Orgao
	Private RH04_ID_ORGAO as Integer
	Private RH04_NM_ORGAO as String
	Private RH04_CD_ORGAO as String

	Public Property IdOrgao() as Integer
		Get
			Return RH04_ID_ORGAO
		End Get
		Set(ByVal Value As Integer)
			RH04_ID_ORGAO = Value
		End Set
	End Property
	Public Property Descricao() as String
		Get
			Return RH04_NM_ORGAO
		End Get
		Set(ByVal Value As String)
			RH04_NM_ORGAO = Value
		End Set
	End Property
	Public Property Codigo() as String
		Get
			Return RH04_CD_ORGAO
		End Get
		Set(ByVal Value As String)
			RH04_CD_ORGAO = Value
		End Set
	End Property

	Public Sub New(Optional ByVal IdOrgao as Integer = 0)
		If IdOrgao >  0 Then
			Obter(IdOrgao)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH04_ORGAO")
		strSQL.Append(" where RH04_ID_ORGAO = " & IdOrgao)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH04_NM_ORGAO") = ProBanco(RH04_NM_ORGAO, eTipoValor.TEXTO)
		dr("RH04_CD_ORGAO") = ProBanco(RH04_CD_ORGAO, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal IdOrgao as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH04_ORGAO")
		strSQL.Append(" where RH04_ID_ORGAO = " & IdOrgao)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH04_ID_ORGAO = DoBanco(dr("RH04_ID_ORGAO"), eTipoValor.CHAVE)
			RH04_NM_ORGAO = DoBanco(dr("RH04_NM_ORGAO"), eTipoValor.TEXTO)
			RH04_CD_ORGAO = DoBanco(dr("RH04_CD_ORGAO"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional IdOrgao as Integer = 0, Optional Descricao as String = "", Optional Codigo as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH04_ORGAO")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH04_ID_ORGAO is not null")
		
		If IdOrgao > 0 then 
			strSQL.Append(" and RH04_ID_ORGAO = " & IdOrgao)
		End If
		
		If Descricao <> "" then 
			strSQL.Append(" and upper(RH04_NM_ORGAO) like '%" & Descricao.toUpper & "%'")
		End If
		
		If Codigo <> "" then 
			strSQL.Append(" and upper(RH04_CD_ORGAO) like '%" & Codigo.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH04_ID_ORGAO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH04_ID_ORGAO as CODIGO, RH04_NM_ORGAO + ' - ' + isnull(RH04_CD_ORGAO,'') as DESCRICAO")
		strSQL.Append(" from RH04_ORGAO")
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
		
		strSQL.Append(" select max(RH04_ID_ORGAO) from RH04_ORGAO")

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
	Public Function Excluir(ByVal IdOrgao as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH04_ORGAO")
		strSQL.Append(" where RH04_ID_ORGAO = " & IdOrgao)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class

'******************************************************************************
'*                                 26/04/2019                                 *
'*                                                                            *
'*          ESTE CÓDIGO FOI GERADO PELO GERA CODIGO VERSÃO 4.0                *
'*    SUPORTE PARA ASP.NET 2.0, AJAX, SQL SERVER COM ENTERPRISE LIBRARY       *
'*                                                                            *
'*  O Gera-Codigo gera um MODELO de código Página, Interface, Classe e Css    *
'*  cabe a cada programador fazer as adaptações quando NECESSÁRIAS.           *
'*                                                                            *
'*  Esta ferramenta é TOTALMENTE GRATUITA, por favor, não remova os créditos  *
'*                                                                            *
'*  O autor não se responsabiliza por qualquer evento acontecido com o uso    *
'*  desta ferramenta ou do sistema que ela vier a gerar.                      *
'*                                                                            *
'*          Desenvolvido por Nírondes Anglada Casanovas Tavares               *
'*                  E-Mail/MSN: nirondes@hotmail.com                          *
'*                                                                            *
'******************************************************************************

