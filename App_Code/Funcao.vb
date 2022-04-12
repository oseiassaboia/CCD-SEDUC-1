Imports Microsoft.VisualBasic
Imports System.Data

Public Class Funcao
	Private RH06_ID_FUNCAO as Integer
	Private RH06_NM_FUNCAO as String
	Private RH06_CD_FOLHA as String

	Public Property FuncaoId() as Integer
		Get
			Return RH06_ID_FUNCAO
		End Get
		Set(ByVal Value As Integer)
			RH06_ID_FUNCAO = Value
		End Set
	End Property
	Public Property Funcao() as String
		Get
			Return RH06_NM_FUNCAO
		End Get
		Set(ByVal Value As String)
			RH06_NM_FUNCAO = Value
		End Set
	End Property
	Public Property FolhaId() as String
		Get
			Return RH06_CD_FOLHA
		End Get
		Set(ByVal Value As String)
			RH06_CD_FOLHA = Value
		End Set
	End Property

    Public Sub New(Optional ByVal FuncaoId As Integer = 0)
        If FuncaoId > 0 Then
            Obter(FuncaoId)
        End If
    End Sub

    Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append("Then select * ")
		strSQL.Append(" from RH06_FUNCAO")
		strSQL.Append(" where RH06_ID_FUNCAO = " & FuncaoId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH06_NM_FUNCAO") = ProBanco(RH06_NM_FUNCAO, eTipoValor.TEXTO)
		dr("RH06_CD_FOLHA") = ProBanco(RH06_CD_FOLHA, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal FuncaoId as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH06_FUNCAO")
		strSQL.Append(" where RH06_ID_FUNCAO = " & FuncaoId)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH06_ID_FUNCAO = DoBanco(dr("RH06_ID_FUNCAO"), eTipoValor.CHAVE)
			RH06_NM_FUNCAO = DoBanco(dr("RH06_NM_FUNCAO"), eTipoValor.TEXTO)
			RH06_CD_FOLHA = DoBanco(dr("RH06_CD_FOLHA"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional FuncaoId as Integer = 0, Optional Funcao as String = "", Optional FolhaId as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH06_FUNCAO")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH06_ID_FUNCAO is not null")
		
		If FuncaoId > 0 then 
			strSQL.Append(" and RH06_ID_FUNCAO = " & FuncaoId)
		End If
		
		If Funcao <> "" then 
			strSQL.Append(" and upper(RH06_NM_FUNCAO) like '%" & Funcao.toUpper & "%'")
		End If
		
		If FolhaId <> "" then 
			strSQL.Append(" and upper(RH06_CD_FOLHA) like '%" & FolhaId.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH06_ID_FUNCAO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH06_ID_FUNCAO as CODIGO, RH06_NM_FUNCAO as DESCRICAO")
		strSQL.Append(" from RH06_FUNCAO")
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
		
		strSQL.Append(" select max(RH06_ID_FUNCAO) from RH06_FUNCAO")

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
	Public Function Excluir(ByVal FuncaoId as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH06_FUNCAO")
		strSQL.Append(" where RH06_ID_FUNCAO = " & FuncaoId)

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

