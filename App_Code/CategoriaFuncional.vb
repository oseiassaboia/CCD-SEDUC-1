Imports Microsoft.VisualBasic
Imports System.Data

Public Class CategoriaFuncional
	Private RH39_ID_CATEGORIA_FUNCIONAL as Integer
	Private RH39_NM_CATEGORIA_FUNCIONAL as String
	Private RH39_SG_CATEGORIA_FUNCIONAL as String

	Public Property CategoriaFuncionalId() as Integer
		Get
			Return RH39_ID_CATEGORIA_FUNCIONAL
		End Get
		Set(ByVal Value As Integer)
			RH39_ID_CATEGORIA_FUNCIONAL = Value
		End Set
	End Property
	Public Property CategoriaFuncional() as String
		Get
			Return RH39_NM_CATEGORIA_FUNCIONAL
		End Get
		Set(ByVal Value As String)
			RH39_NM_CATEGORIA_FUNCIONAL = Value
		End Set
	End Property
	Public Property Sigla() as String
		Get
			Return RH39_SG_CATEGORIA_FUNCIONAL
		End Get
		Set(ByVal Value As String)
			RH39_SG_CATEGORIA_FUNCIONAL = Value
		End Set
	End Property

    Public Sub New(Optional ByVal CategoriaFuncionalId As Integer = 0)
        If CategoriaFuncionalId > 0 Then
            Obter(CategoriaFuncionalId)
        End If
    End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append("Then select * ")
		strSQL.Append(" from RH39_CATEGORIA_FUNCIONAL")
		strSQL.Append(" where RH39_ID_CATEGORIA_FUNCIONAL = " & CategoriaFuncionalId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH39_NM_CATEGORIA_FUNCIONAL") = ProBanco(RH39_NM_CATEGORIA_FUNCIONAL, eTipoValor.TEXTO)
		dr("RH39_SG_CATEGORIA_FUNCIONAL") = ProBanco(RH39_SG_CATEGORIA_FUNCIONAL, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal CategoriaFuncionalId as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH39_CATEGORIA_FUNCIONAL")
		strSQL.Append(" where RH39_ID_CATEGORIA_FUNCIONAL = " & CategoriaFuncionalId)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH39_ID_CATEGORIA_FUNCIONAL = DoBanco(dr("RH39_ID_CATEGORIA_FUNCIONAL"), eTipoValor.CHAVE)
			RH39_NM_CATEGORIA_FUNCIONAL = DoBanco(dr("RH39_NM_CATEGORIA_FUNCIONAL"), eTipoValor.TEXTO)
			RH39_SG_CATEGORIA_FUNCIONAL = DoBanco(dr("RH39_SG_CATEGORIA_FUNCIONAL"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional CategoriaFuncionalId as Integer = 0, Optional CategoriaFuncional as String = "", Optional Sigla as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH39_CATEGORIA_FUNCIONAL")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH39_ID_CATEGORIA_FUNCIONAL is not null")
		
		If CategoriaFuncionalId > 0 then 
			strSQL.Append(" and RH39_ID_CATEGORIA_FUNCIONAL = " & CategoriaFuncionalId)
		End If
		
		If CategoriaFuncional <> "" then 
			strSQL.Append(" and upper(RH39_NM_CATEGORIA_FUNCIONAL) like '%" & CategoriaFuncional.toUpper & "%'")
		End If
		
		If Sigla <> "" then 
			strSQL.Append(" and upper(RH39_SG_CATEGORIA_FUNCIONAL) like '%" & Sigla.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH39_ID_CATEGORIA_FUNCIONAL", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH39_ID_CATEGORIA_FUNCIONAL as CODIGO, RH39_NM_CATEGORIA_FUNCIONAL as DESCRICAO")
		strSQL.Append(" from RH39_CATEGORIA_FUNCIONAL")
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
		
		strSQL.Append(" select max(RH39_ID_CATEGORIA_FUNCIONAL) from RH39_CATEGORIA_FUNCIONAL")

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
	Public Function Excluir(ByVal CategoriaFuncionalId as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH39_CATEGORIA_FUNCIONAL")
		strSQL.Append(" where RH39_ID_CATEGORIA_FUNCIONAL = " & CategoriaFuncionalId)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class

'******************************************************************************
'*                                 11/09/2018                                 *
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

