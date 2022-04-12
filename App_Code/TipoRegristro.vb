Imports Microsoft.VisualBasic
Imports System.Data

Public Class TipoRegristro
	Private RH23_ID_TIPO_REGISTRO as Integer
	Private RH23_NM_TIPO_REGISTRO as String
	Private RH23_SG_TIPO_REGISTRO as String

	Public Property Codigo() as Integer
		Get
			Return RH23_ID_TIPO_REGISTRO
		End Get
		Set(ByVal Value As Integer)
			RH23_ID_TIPO_REGISTRO = Value
		End Set
	End Property
	Public Property Descricao() as String
		Get
			Return RH23_NM_TIPO_REGISTRO
		End Get
		Set(ByVal Value As String)
			RH23_NM_TIPO_REGISTRO = Value
		End Set
	End Property
	Public Property SiglaTipoRegistro() as String
		Get
			Return RH23_SG_TIPO_REGISTRO
		End Get
		Set(ByVal Value As String)
			RH23_SG_TIPO_REGISTRO = Value
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
		strSQL.Append(" from RH23_TIPO_REGISTRO")
		strSQL.Append(" where RH23_ID_TIPO_REGISTRO = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH23_NM_TIPO_REGISTRO") = ProBanco(RH23_NM_TIPO_REGISTRO, eTipoValor.TEXTO)
		dr("RH23_SG_TIPO_REGISTRO") = ProBanco(RH23_SG_TIPO_REGISTRO, eTipoValor.TEXTO)

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
		strSQL.Append(" from RH23_TIPO_REGISTRO")
		strSQL.Append(" where RH23_ID_TIPO_REGISTRO = " & Codigo)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH23_ID_TIPO_REGISTRO = DoBanco(dr("RH23_ID_TIPO_REGISTRO"), eTipoValor.CHAVE)
			RH23_NM_TIPO_REGISTRO = DoBanco(dr("RH23_NM_TIPO_REGISTRO"), eTipoValor.TEXTO)
			RH23_SG_TIPO_REGISTRO = DoBanco(dr("RH23_SG_TIPO_REGISTRO"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional Codigo as Integer = 0, Optional Descricao as String = "", Optional SiglaTipoRegistro as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH23_TIPO_REGISTRO")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH23_ID_TIPO_REGISTRO is not null")
		
		If Codigo > 0 then 
			strSQL.Append(" and RH23_ID_TIPO_REGISTRO = " & Codigo)
		End If
		
		If Descricao <> "" then 
			strSQL.Append(" and upper(RH23_NM_TIPO_REGISTRO) like '%" & Descricao.toUpper & "%'")
		End If
		
		If SiglaTipoRegistro <> "" then 
			strSQL.Append(" and upper(RH23_SG_TIPO_REGISTRO) like '%" & SiglaTipoRegistro.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH23_ID_TIPO_REGISTRO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH23_ID_TIPO_REGISTRO as CODIGO, RH23_NM_TIPO_REGISTRO as DESCRICAO")
		strSQL.Append(" from RH23_TIPO_REGISTRO")
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
		
		strSQL.Append(" select max(RH23_ID_TIPO_REGISTRO) from RH23_TIPO_REGISTRO")

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
		strSQL.Append(" from RH23_TIPO_REGISTRO")
		strSQL.Append(" where RH23_ID_TIPO_REGISTRO = " & Codigo)

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

