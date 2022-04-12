Imports Microsoft.VisualBasic
Imports System.Data

Public Class LancamentoFrequencia
	Private RH44_ID_LANCAMENTO_FREQ as Integer
	Private RH44_NM_LANCAMENTO_FREQ as String

	Public Property Codigo() as Integer
		Get
			Return RH44_ID_LANCAMENTO_FREQ
		End Get
		Set(ByVal Value As Integer)
			RH44_ID_LANCAMENTO_FREQ = Value
		End Set
	End Property
	Public Property DescricaoLancamento() as String
		Get
			Return RH44_NM_LANCAMENTO_FREQ
		End Get
		Set(ByVal Value As String)
			RH44_NM_LANCAMENTO_FREQ = Value
		End Set
	End Property

	Public Sub New(Optional ByVal Codigo as integer = 0)
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
		strSQL.Append(" from RH44_LANCAMENTO_FREQ")
		strSQL.Append(" where RH44_ID_LANCAMENTO_FREQ = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH44_NM_LANCAMENTO_FREQ") = ProBanco(RH44_NM_LANCAMENTO_FREQ, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal Codigo as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH44_LANCAMENTO_FREQ")
		strSQL.Append(" where RH44_ID_LANCAMENTO_FREQ = " & Codigo)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH44_ID_LANCAMENTO_FREQ = DoBanco(dr("RH44_ID_LANCAMENTO_FREQ"), eTipoValor.CHAVE)
			RH44_NM_LANCAMENTO_FREQ = DoBanco(dr("RH44_NM_LANCAMENTO_FREQ"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional Codigo as Integer = 0, Optional DescricaoLancamento as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH44_LANCAMENTO_FREQ")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH44_ID_LANCAMENTO_FREQ is not null")
		
		If Codigo > 0 then 
			strSQL.Append(" and RH44_ID_LANCAMENTO_FREQ = " & Codigo)
		End If
		
		If DescricaoLancamento <> "" then 
			strSQL.Append(" and upper(RH44_NM_LANCAMENTO_FREQ) like '%" & DescricaoLancamento.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH44_ID_LANCAMENTO_FREQ", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH44_ID_LANCAMENTO_FREQ as CODIGO, RH44_NM_LANCAMENTO_FREQ as DESCRICAO")
		strSQL.Append(" from RH44_LANCAMENTO_FREQ")
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
		
		strSQL.Append(" select max(RH44_ID_LANCAMENTO_FREQ) from RH44_LANCAMENTO_FREQ")

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
	Public Function Excluir(ByVal Codigo as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH44_LANCAMENTO_FREQ")
		strSQL.Append(" where RH44_ID_LANCAMENTO_FREQ = " & Codigo)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class

'******************************************************************************
'*                                 20/05/2019                                 *
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

