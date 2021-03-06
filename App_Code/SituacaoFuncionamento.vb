Imports Microsoft.VisualBasic
Imports System.Data

Public Class SituacaoFuncionamento
	Private RH60_ID_SIT_FUNCIONAMENTO as Integer
	Private RH60_NM_SIT_FUNCIONAMENTO as String

	Public Property IdSituacaoFuncionamento() as Integer
		Get
			Return RH60_ID_SIT_FUNCIONAMENTO
		End Get
		Set(ByVal Value As Integer)
			RH60_ID_SIT_FUNCIONAMENTO = Value
		End Set
	End Property
	Public Property Descricao() as String
		Get
			Return RH60_NM_SIT_FUNCIONAMENTO
		End Get
		Set(ByVal Value As String)
			RH60_NM_SIT_FUNCIONAMENTO = Value
		End Set
	End Property

	Public Sub New(Optional ByVal IdSituacaoFuncionamento as integer = 0)
		If IdSituacaoFuncionamento > 0 Then
			Obter(IdSituacaoFuncionamento)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH60_SIT_FUNCIONAMENTO")
		strSQL.Append(" where RH60_ID_SIT_FUNCIONAMENTO = " & IdSituacaoFuncionamento)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH60_NM_SIT_FUNCIONAMENTO") = ProBanco(RH60_NM_SIT_FUNCIONAMENTO, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal IdSituacaoFuncionamento as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH60_SIT_FUNCIONAMENTO")
		strSQL.Append(" where RH60_ID_SIT_FUNCIONAMENTO = " & IdSituacaoFuncionamento)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH60_ID_SIT_FUNCIONAMENTO = DoBanco(dr("RH60_ID_SIT_FUNCIONAMENTO"), eTipoValor.CHAVE)
			RH60_NM_SIT_FUNCIONAMENTO = DoBanco(dr("RH60_NM_SIT_FUNCIONAMENTO"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional IdSituacaoFuncionamento as Integer = 0, Optional Descricao as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH60_SIT_FUNCIONAMENTO")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH60_ID_SIT_FUNCIONAMENTO is not null")
		
		If IdSituacaoFuncionamento > 0 then 
			strSQL.Append(" and RH60_ID_SIT_FUNCIONAMENTO = " & IdSituacaoFuncionamento)
		End If
		
		If Descricao <> "" then 
			strSQL.Append(" and upper(RH60_NM_SIT_FUNCIONAMENTO) like '%" & Descricao.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH60_ID_SIT_FUNCIONAMENTO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH60_ID_SIT_FUNCIONAMENTO as CODIGO, RH60_NM_SIT_FUNCIONAMENTO as DESCRICAO")
		strSQL.Append(" from RH60_SIT_FUNCIONAMENTO")
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
		
		strSQL.Append(" select max(RH60_ID_SIT_FUNCIONAMENTO) from RH60_SIT_FUNCIONAMENTO")

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
	Public Function Excluir(ByVal IdSituacaoFuncionamento as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH60_SIT_FUNCIONAMENTO")
		strSQL.Append(" where RH60_ID_SIT_FUNCIONAMENTO = " & IdSituacaoFuncionamento)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class

'******************************************************************************
'*                                 09/04/2019                                 *
'*                                                                            *
'*          ESTE C?DIGO FOI GERADO PELO GERA CODIGO VERS?O 4.0                *
'*    SUPORTE PARA ASP.NET 2.0, AJAX, SQL SERVER COM ENTERPRISE LIBRARY       *
'*                                                                            *
'*  O Gera-Codigo gera um MODELO de c?digo P?gina, Interface, Classe e Css    *
'*  cabe a cada programador fazer as adapta??es quando NECESS?RIAS.           *
'*                                                                            *
'*  Esta ferramenta ? TOTALMENTE GRATUITA, por favor, n?o remova os cr?ditos  *
'*                                                                            *
'*  O autor n?o se responsabiliza por qualquer evento acontecido com o uso    *
'*  desta ferramenta ou do sistema que ela vier a gerar.                      *
'*                                                                            *
'*          Desenvolvido por N?rondes Anglada Casanovas Tavares               *
'*                  E-Mail/MSN: nirondes@hotmail.com                          *
'*                                                                            *
'******************************************************************************

