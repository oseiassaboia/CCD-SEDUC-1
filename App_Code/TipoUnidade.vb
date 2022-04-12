Imports Microsoft.VisualBasic
Imports System.Data

Public Class TipoUnidade
	Private RH47_ID_TIPO_UNIDADE as Integer
	Private RH47_NM_TIPO_UNIDADE as String

	Public Property IdTipoUnidade() as Integer
		Get
			Return RH47_ID_TIPO_UNIDADE
		End Get
		Set(ByVal Value As Integer)
			RH47_ID_TIPO_UNIDADE = Value
		End Set
	End Property
	Public Property Descricao() as String
		Get
			Return RH47_NM_TIPO_UNIDADE
		End Get
		Set(ByVal Value As String)
			RH47_NM_TIPO_UNIDADE = Value
		End Set
	End Property

	Public Sub New(Optional ByVal IdTipoUnidade as Integer = 0)
		If IdTipoUnidade > 0 Then
			Obter(IdTipoUnidade)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH47_TIPO_UNIDADE")
		strSQL.Append(" where RH47_ID_TIPO_UNIDADE = " & IdTipoUnidade)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH47_NM_TIPO_UNIDADE") = ProBanco(RH47_NM_TIPO_UNIDADE, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal IdTipoUnidade as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH47_TIPO_UNIDADE")
		strSQL.Append(" where RH47_ID_TIPO_UNIDADE = " & IdTipoUnidade)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH47_ID_TIPO_UNIDADE = DoBanco(dr("RH47_ID_TIPO_UNIDADE"), eTipoValor.CHAVE)
			RH47_NM_TIPO_UNIDADE = DoBanco(dr("RH47_NM_TIPO_UNIDADE"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional IdTipoUnidade as Integer = 0, Optional Descricao as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH47_TIPO_UNIDADE")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH47_ID_TIPO_UNIDADE is not null")
		
		If IdTipoUnidade > 0 then 
			strSQL.Append(" and RH47_ID_TIPO_UNIDADE = " & IdTipoUnidade)
		End If
		
		If Descricao <> "" then 
			strSQL.Append(" and upper(RH47_NM_TIPO_UNIDADE) like '%" & Descricao.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH47_ID_TIPO_UNIDADE", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH47_ID_TIPO_UNIDADE as CODIGO, RH47_NM_TIPO_UNIDADE as DESCRICAO")
		strSQL.Append(" from RH47_TIPO_UNIDADE")
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
		
		strSQL.Append(" select max(RH47_ID_TIPO_UNIDADE) from RH47_TIPO_UNIDADE")

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
	Public Function Excluir(ByVal IdTipoUnidade as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH47_TIPO_UNIDADE")
		strSQL.Append(" where RH47_ID_TIPO_UNIDADE = " & IdTipoUnidade)

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

