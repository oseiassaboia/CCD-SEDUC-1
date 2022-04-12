Imports Microsoft.VisualBasic
Imports System.Data

Public Class TipoEscola
	Private RH48_ID_TIPO_ESCOLA as Integer
	Private RH48_NM_TIPO_ESCOLA as String

	Public Property IdTipoEscola() as Integer
		Get
			Return RH48_ID_TIPO_ESCOLA
		End Get
		Set(ByVal Value As Integer)
			RH48_ID_TIPO_ESCOLA = Value
		End Set
	End Property
	Public Property Descricao() as String
		Get
			Return RH48_NM_TIPO_ESCOLA
		End Get
		Set(ByVal Value As String)
			RH48_NM_TIPO_ESCOLA = Value
		End Set
	End Property

	Public Sub New(Optional ByVal IdTipoEscola as Integer = 0)
		If IdTipoEscola > 0 Then
			Obter(IdTipoEscola)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH48_TIPO_ESCOLA")
		strSQL.Append(" where RH48_ID_TIPO_ESCOLA = " & IdTipoEscola)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH48_NM_TIPO_ESCOLA") = ProBanco(RH48_NM_TIPO_ESCOLA, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal IdTipoEscola as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH48_TIPO_ESCOLA")
		strSQL.Append(" where RH48_ID_TIPO_ESCOLA = " & IdTipoEscola)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH48_ID_TIPO_ESCOLA = DoBanco(dr("RH48_ID_TIPO_ESCOLA"), eTipoValor.CHAVE)
			RH48_NM_TIPO_ESCOLA = DoBanco(dr("RH48_NM_TIPO_ESCOLA"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional IdTipoEscola as Integer = 0, Optional Descricao as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH48_TIPO_ESCOLA")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH48_ID_TIPO_ESCOLA is not null")
		
		If IdTipoEscola > 0 then 
			strSQL.Append(" and RH48_ID_TIPO_ESCOLA = " & IdTipoEscola)
		End If
		
		If Descricao <> "" then 
			strSQL.Append(" and upper(RH48_NM_TIPO_ESCOLA) like '%" & Descricao.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH48_ID_TIPO_ESCOLA", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH48_ID_TIPO_ESCOLA as CODIGO, RH48_NM_TIPO_ESCOLA as DESCRICAO")
		strSQL.Append(" from RH48_TIPO_ESCOLA")
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
		
		strSQL.Append(" select max(RH48_ID_TIPO_ESCOLA) from RH48_TIPO_ESCOLA")

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
	Public Function Excluir(ByVal IdTipoEscola as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH48_TIPO_ESCOLA")
		strSQL.Append(" where RH48_ID_TIPO_ESCOLA = " & IdTipoEscola)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class

'******************************************************************************
'*                                 09/04/2019                                 *
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

