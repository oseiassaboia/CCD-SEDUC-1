Imports Microsoft.VisualBasic
Imports System.Data

Public Class PessoaDeficiencia
	Private RH46_ID_PESSOA_DEFICIENCIA as Integer
	Private RH01_ID_PESSOA as String
	Private TG15_ID_DEFICIENCIA as String
	Private RH46_DS_OBSERVACAO as String

	Public Property PessoDeficienciaId() as Integer
		Get
			Return RH46_ID_PESSOA_DEFICIENCIA
		End Get
		Set(ByVal Value As Integer)
			RH46_ID_PESSOA_DEFICIENCIA = Value
		End Set
	End Property
	Public Property PessoaId() as String
		Get
			Return RH01_ID_PESSOA
		End Get
		Set(ByVal Value As String)
			RH01_ID_PESSOA = Value
		End Set
	End Property
	Public Property DeficienciaId() as String
		Get
			Return TG15_ID_DEFICIENCIA
		End Get
		Set(ByVal Value As String)
			TG15_ID_DEFICIENCIA = Value
		End Set
	End Property
	Public Property Observacao() as String
		Get
			Return RH46_DS_OBSERVACAO
		End Get
		Set(ByVal Value As String)
			RH46_DS_OBSERVACAO = Value
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
		strSQL.Append(" from RH46_PESSOA_DEFICIENCIA")
		strSQL.Append(" where RH46_ID_PESSOA_DEFICIENCIA = " & PessoDeficienciaId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH01_ID_PESSOA") = ProBanco(RH01_ID_PESSOA, eTipoValor.CHAVE)
		dr("TG15_ID_DEFICIENCIA") = ProBanco(TG15_ID_DEFICIENCIA, eTipoValor.CHAVE)
		dr("RH46_DS_OBSERVACAO") = ProBanco(RH46_DS_OBSERVACAO, eTipoValor.TEXTO)

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
		strSQL.Append(" from RH46_PESSOA_DEFICIENCIA")
		strSQL.Append(" where RH46_ID_PESSOA_DEFICIENCIA = " & Codigo)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH46_ID_PESSOA_DEFICIENCIA = DoBanco(dr("RH46_ID_PESSOA_DEFICIENCIA"), eTipoValor.CHAVE)
			RH01_ID_PESSOA = DoBanco(dr("RH01_ID_PESSOA"), eTipoValor.chave)
			TG15_ID_DEFICIENCIA = DoBanco(dr("TG15_ID_DEFICIENCIA"), eTipoValor.chave)
			RH46_DS_OBSERVACAO = DoBanco(dr("RH46_DS_OBSERVACAO"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional PessoDeficienciaId as Integer = 0, Optional PessoaId as integer = 0, Optional DeficienciaId as Integer = 0, Optional Observacao as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH46_PESSOA_DEFICIENCIA PessiaDeficiencia")
		strSQL.Append(" left join DBGERAL..TG15_DEFICIENCIA Deficiencia on Deficiencia.TG15_ID_DEFICIENCIA = PessiaDeficiencia.TG15_ID_DEFICIENCIA ")
		strSQL.Append(" where RH46_ID_PESSOA_DEFICIENCIA is not null")
		
		If PessoDeficienciaId > 0 then 
			strSQL.Append(" and RH46_ID_PESSOA_DEFICIENCIA = " & PessoDeficienciaId)
		End If
		
		If PessoaId > 0 then
			strSQL.Append(" and RH01_ID_PESSOA = " & PessoaId )
		End If
		
		If  DeficienciaId > 0 then
			strSQL.Append(" and TG15_ID_DEFICIENCIA = " & DeficienciaId )
		End If
		
		If Observacao <> "" then 
			strSQL.Append(" and upper(RH46_DS_OBSERVACAO) like '%" & Observacao.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH46_ID_PESSOA_DEFICIENCIA", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH46_ID_PESSOA_DEFICIENCIA as CODIGO, RH01_ID_PESSOA as DESCRICAO")
		strSQL.Append(" from RH46_PESSOA_DEFICIENCIA")
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
		
		strSQL.Append(" select max(RH46_ID_PESSOA_DEFICIENCIA) from RH46_PESSOA_DEFICIENCIA")

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
	Public Function Excluir(ByVal PessoDeficienciaId as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH46_PESSOA_DEFICIENCIA")
		strSQL.Append(" where RH46_ID_PESSOA_DEFICIENCIA = " & PessoDeficienciaId)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class

'******************************************************************************
'*                                 12/09/2018                                 *
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

