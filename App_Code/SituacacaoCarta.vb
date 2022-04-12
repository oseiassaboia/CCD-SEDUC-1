Imports Microsoft.VisualBasic
Imports System.Data

Public Class SituacacaoCarta
	Private RH55_ID_SITUACAO_CARTA as Integer
	Private RH55_NM_SITUACAO_CARTA as String

	Public Property SituacaoCartaId() as Integer
		Get
			Return RH55_ID_SITUACAO_CARTA
		End Get
		Set(ByVal Value As Integer)
			RH55_ID_SITUACAO_CARTA = Value
		End Set
	End Property
	Public Property DescricaoCarta() as String
		Get
			Return RH55_NM_SITUACAO_CARTA
		End Get
		Set(ByVal Value As String)
			RH55_NM_SITUACAO_CARTA = Value
		End Set
	End Property

	Public Sub New(Optional ByVal SituacaoCartaId as integer = 0)
		If SituacaoCartaId > 0 Then
			Obter(SituacaoCartaId)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH55_SITUACAO_CARTA")
		strSQL.Append(" where RH55_ID_SITUACAO_CARTA = " & SituacaoCartaId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH55_NM_SITUACAO_CARTA") = ProBanco(RH55_NM_SITUACAO_CARTA, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal SituacaoCartaId as integer)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH55_SITUACAO_CARTA")
		strSQL.Append(" where RH55_ID_SITUACAO_CARTA = " & SituacaoCartaId)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH55_ID_SITUACAO_CARTA = DoBanco(dr("RH55_ID_SITUACAO_CARTA"), eTipoValor.CHAVE)
			RH55_NM_SITUACAO_CARTA = DoBanco(dr("RH55_NM_SITUACAO_CARTA"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional SituacaoCartaId as Integer = 0, Optional DescricaoCarta as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH55_SITUACAO_CARTA")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH55_ID_SITUACAO_CARTA is not null")
		
		If SituacaoCartaId > 0 then 
			strSQL.Append(" and RH55_ID_SITUACAO_CARTA = " & SituacaoCartaId)
		End If
		
		If DescricaoCarta <> "" then 
			strSQL.Append(" and upper(RH55_NM_SITUACAO_CARTA) like '%" & DescricaoCarta.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH55_ID_SITUACAO_CARTA", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH55_ID_SITUACAO_CARTA as CODIGO, RH55_NM_SITUACAO_CARTA as DESCRICAO")
		strSQL.Append(" from RH55_SITUACAO_CARTA")
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
		
		strSQL.Append(" select max(RH55_ID_SITUACAO_CARTA) from RH55_SITUACAO_CARTA")

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
	Public Function Excluir(ByVal SituacaoCartaId as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH55_SITUACAO_CARTA")
		strSQL.Append(" where RH55_ID_SITUACAO_CARTA = " & SituacaoCartaId)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class

'******************************************************************************
'*                                 14/01/2019                                 *
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

