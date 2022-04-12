Imports Microsoft.VisualBasic
Imports System.Data

Public Class TipoDependenciaAdministrativa
	Private RH61_ID_TIPO_DEPEND_ADM as Integer
	Private RH61_NM_TIPO_DEPEND_ADM as String

	Public Property IdTipoDependenciaAdministrativa() as Integer
		Get
			Return RH61_ID_TIPO_DEPEND_ADM
		End Get
		Set(ByVal Value As Integer)
			RH61_ID_TIPO_DEPEND_ADM = Value
		End Set
	End Property
	Public Property Descricao() as String
		Get
			Return RH61_NM_TIPO_DEPEND_ADM
		End Get
		Set(ByVal Value As String)
			RH61_NM_TIPO_DEPEND_ADM = Value
		End Set
	End Property

	Public Sub New(Optional ByVal IdTipoDependenciaAdministrativa as Integer = 0)
		If IdTipoDependenciaAdministrativa > 0 Then
			Obter(IdTipoDependenciaAdministrativa)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH61_TIPO_DEPEND_ADM")
		strSQL.Append(" where RH61_ID_TIPO_DEPEND_ADM = " & IdTipoDependenciaAdministrativa)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH61_NM_TIPO_DEPEND_ADM") = ProBanco(RH61_NM_TIPO_DEPEND_ADM, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal IdTipoDependenciaAdministrativa as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH61_TIPO_DEPEND_ADM")
		strSQL.Append(" where RH61_ID_TIPO_DEPEND_ADM = " & IdTipoDependenciaAdministrativa)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH61_ID_TIPO_DEPEND_ADM = DoBanco(dr("RH61_ID_TIPO_DEPEND_ADM"), eTipoValor.CHAVE)
			RH61_NM_TIPO_DEPEND_ADM = DoBanco(dr("RH61_NM_TIPO_DEPEND_ADM"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional IdTipoDependenciaAdministrativa as Integer = 0, Optional Descricao as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH61_TIPO_DEPEND_ADM")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH61_ID_TIPO_DEPEND_ADM is not null")
		
		If IdTipoDependenciaAdministrativa > 0 then 
			strSQL.Append(" and RH61_ID_TIPO_DEPEND_ADM = " & IdTipoDependenciaAdministrativa)
		End If
		
		If Descricao <> "" then 
			strSQL.Append(" and upper(RH61_NM_TIPO_DEPEND_ADM) like '%" & Descricao.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH61_ID_TIPO_DEPEND_ADM", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH61_ID_TIPO_DEPEND_ADM as CODIGO, RH61_NM_TIPO_DEPEND_ADM as DESCRICAO")
		strSQL.Append(" from RH61_TIPO_DEPEND_ADM")
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
		
		strSQL.Append(" select max(RH61_ID_TIPO_DEPEND_ADM) from RH61_TIPO_DEPEND_ADM")

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
	Public Function Excluir(ByVal IdTipoDependenciaAdministrativa as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH61_TIPO_DEPEND_ADM")
		strSQL.Append(" where RH61_ID_TIPO_DEPEND_ADM = " & IdTipoDependenciaAdministrativa)

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

