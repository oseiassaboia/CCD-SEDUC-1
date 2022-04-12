Imports Microsoft.VisualBasic
Imports System.Data

Public Class Banco
	Private TG48_ID_BANCO as Integer
	Private TG48_NM_BANCO as String
	Private TG48_SG_BANCO as String

	Public Property BancoId() as Integer
		Get
			Return TG48_ID_BANCO
		End Get
		Set(ByVal Value As Integer)
			TG48_ID_BANCO = Value
		End Set
	End Property
	Public Property Banco() as String
		Get
			Return TG48_NM_BANCO
		End Get
		Set(ByVal Value As String)
			TG48_NM_BANCO = Value
		End Set
	End Property
	Public Property Sigla() as String
		Get
			Return TG48_SG_BANCO
		End Get
		Set(ByVal Value As String)
			TG48_SG_BANCO = Value
		End Set
	End Property

	Public Sub New(Optional ByVal BancoId as integer = 0)
		If BancoId > 0 Then
			Obter(BancoId)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from DBGERAL..TG48_BANCO")
		strSQL.Append(" where TG48_ID_BANCO = " & BancoId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("TG48_NM_BANCO") = ProBanco(TG48_NM_BANCO, eTipoValor.TEXTO)
		dr("TG48_SG_BANCO") = ProBanco(TG48_SG_BANCO, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal BancoId as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from DBGERAL..TG48_BANCO")
		strSQL.Append(" where TG48_ID_BANCO = " & BancoId)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			TG48_ID_BANCO = DoBanco(dr("TG48_ID_BANCO"), eTipoValor.CHAVE)
			TG48_NM_BANCO = DoBanco(dr("TG48_NM_BANCO"), eTipoValor.TEXTO)
			TG48_SG_BANCO = DoBanco(dr("TG48_SG_BANCO"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional BancoId as Integer = 0, Optional Banco as String = "", Optional Sigla as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from DBGERAL..TG48_BANCO")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where TG48_ID_BANCO is not null")
		
		If BancoId > 0 then 
			strSQL.Append(" and TG48_ID_BANCO = " & BancoId)
		End If
		
		If Banco <> "" then 
			strSQL.Append(" and upper(TG48_NM_BANCO) like '%" & Banco.toUpper & "%'")
		End If
		
		If Sigla <> "" then 
			strSQL.Append(" and upper(TG48_SG_BANCO) like '%" & Sigla.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "TG48_ID_BANCO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select TG48_ID_BANCO as CODIGO, TG48_SG_BANCO +' - '+TG48_NM_BANCO as DESCRICAO")
		strSQL.Append(" from DBGERAL..TG48_BANCO")
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
		
		strSQL.Append(" select max(TG48_ID_BANCO) from TG48_BANCO")

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
	Public Function Excluir(ByVal BancoId as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from TG48_BANCO")
		strSQL.Append(" where TG48_ID_BANCO = " & BancoId)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class

'******************************************************************************
'*                                 14/09/2018                                 *
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

