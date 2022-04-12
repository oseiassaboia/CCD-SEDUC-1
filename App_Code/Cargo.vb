Imports Microsoft.VisualBasic
Imports System.Data

Public Class Cargo
	Private RH16_ID_CARGO as Integer
	Private RH04_ID_ORGAO as String
	Private RH16_NM_CARGO as String
	'Private RH16_IN_COMISSIONADO as Integer
	Private RH16_CD_FOLHA As String
	Private RH85_ID_CATEGORIA_CARGO As Integer

	Public Property CargoId() as Integer
		Get
			Return RH16_ID_CARGO
		End Get
		Set(ByVal Value As Integer)
			RH16_ID_CARGO = Value
		End Set
	End Property
	Public Property OrgaoId() as String
		Get
			Return RH04_ID_ORGAO
		End Get
		Set(ByVal Value As String)
			RH04_ID_ORGAO = Value
		End Set
	End Property
	Public Property Cargo() as String
		Get
			Return RH16_NM_CARGO
		End Get
		Set(ByVal Value As String)
			RH16_NM_CARGO = Value
		End Set
	End Property
	'Public Property Comissionado() as Integer
	'	Get
	'		Return RH16_IN_COMISSIONADO
	'	End Get
	'	Set(ByVal Value As Integer)
	'		RH16_IN_COMISSIONADO = Value
	'	End Set
	'End Property
	Public Property FolhaId() As String
		Get
			Return RH16_CD_FOLHA
		End Get
		Set(ByVal Value As String)
			RH16_CD_FOLHA = Value
		End Set
	End Property
	Public Property CategoriaCargo() As Integer
		Get
			Return RH85_ID_CATEGORIA_CARGO
		End Get
		Set(value As Integer)
			RH85_ID_CATEGORIA_CARGO = value
		End Set
	End Property

	Public Sub New(Optional ByVal CargoId as integer = 0)
		If CargoId > 0 Then
			Obter(CargoId)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH16_CARGO")
		strSQL.Append(" where RH16_ID_CARGO = " & CargoId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

        dr("RH04_ID_ORGAO") = ProBanco(RH04_ID_ORGAO, eTipoValor.CHAVE)
        dr("RH16_NM_CARGO") = ProBanco(RH16_NM_CARGO, eTipoValor.TEXTO)
		'dr("RH16_IN_COMISSIONADO") = ProBanco(RH16_IN_COMISSIONADO, eTipoValor.CHAVE)
		dr("RH16_CD_FOLHA") = ProBanco(RH16_CD_FOLHA, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal CargoId as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH16_CARGO")
		strSQL.Append(" where RH16_ID_CARGO = " & CargoId)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH16_ID_CARGO = DoBanco(dr("RH16_ID_CARGO"), eTipoValor.CHAVE)
            RH04_ID_ORGAO = DoBanco(dr("RH04_ID_ORGAO"), eTipoValor.CHAVE)
            RH16_NM_CARGO = DoBanco(dr("RH16_NM_CARGO"), eTipoValor.TEXTO)
			'RH16_IN_COMISSIONADO = DoBanco(dr("RH16_IN_COMISSIONADO"), eTipoValor.BOOLEANO)
			RH16_CD_FOLHA = DoBanco(dr("RH16_CD_FOLHA"), eTipoValor.TEXTO)
			RH85_ID_CATEGORIA_CARGO = DoBanco(dr("RH85_ID_CATEGORIA_CARGO"), eTipoValor.CHAVE)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional CargoId as Integer = 0, Optional OrgaoId as Integer = 0, Optional Cargo as string = "", Optional Comissionado as Integer = 0, Optional FolhaId as Integer = 0) as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH16_CARGO")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH16_ID_CARGO is not null")
		
		If CargoId > 0 then 
			strSQL.Append(" and RH16_ID_CARGO = " & CargoId)
		End If
		
		If OrgaoId > 0 then
			strSQL.Append(" and RH04_ID_ORGAO = " & OrgaoId)
		End If
		
		If Cargo <> "" then 
			strSQL.Append(" and upper(RH16_NM_CARGO) like '%" & Cargo.toUpper & "%'")
		End If
		
		If Comissionado > 0 then 
			strSQL.Append(" and RH16_IN_COMISSIONADO = " & Comissionado)
		End If
		
		If FolhaId > 0 then 
			strSQL.Append(" and upper(RH16_CD_FOLHA)  = " & FolhaId )
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH16_ID_CARGO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela(Optional Orgao As Integer = 0) as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH16_ID_CARGO as CODIGO, RH16_NM_CARGO as DESCRICAO")
		strSQL.Append(" from RH16_CARGO")
	    strSQL.Append(" where RH16_ID_CARGO is not null ")

        If Orgao > 0 Then
            strSQL.Append(" and RH04_ID_ORGAO=" & Orgao)
        End If


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
		
		strSQL.Append(" select max(RH16_ID_CARGO) from RH16_CARGO")

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
	Public Function Excluir(ByVal CargoId as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH16_CARGO")
		strSQL.Append(" where RH16_ID_CARGO = " & CargoId)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class

'******************************************************************************
'*                                 10/09/2018                                 *
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

