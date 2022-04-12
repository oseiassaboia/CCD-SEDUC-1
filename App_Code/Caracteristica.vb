Imports System.Data

Public Class Caracteristica
	Private RH68_ID_CARACTERISTICA as Integer
	Private RH01_ID_PESSOA as String
	Private TG09_ID_GRUPO_SANGUINEO as String
	Private TG10_ID_RELIGIAO as String
	Private TG42_ID_ETNIA as String
	Private TG11_ID_RACA_COR as String

	Public Property CaracteristicaId() as Integer
		Get
			Return RH68_ID_CARACTERISTICA
		End Get
		Set(ByVal Value As Integer)
			RH68_ID_CARACTERISTICA = Value
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
	Public Property GrupoSanguineoId() as String
		Get
			Return TG09_ID_GRUPO_SANGUINEO
		End Get
		Set(ByVal Value As String)
			TG09_ID_GRUPO_SANGUINEO = Value
		End Set
	End Property
	Public Property ReligiaoId() as String
		Get
			Return TG10_ID_RELIGIAO
		End Get
		Set(ByVal Value As String)
			TG10_ID_RELIGIAO = Value
		End Set
	End Property
	Public Property EtniaId() as String
		Get
			Return TG42_ID_ETNIA
		End Get
		Set(ByVal Value As String)
			TG42_ID_ETNIA = Value
		End Set
	End Property
	Public Property RacaCorId() as String
		Get
			Return TG11_ID_RACA_COR
		End Get
		Set(ByVal Value As String)
			TG11_ID_RACA_COR = Value
		End Set
	End Property

	Public Sub New(Optional ByVal CaracteristicaId as Integer = 0 )
		If CaracteristicaId >0 Then
			Obter(CaracteristicaId)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH68_CARACTERISTICA")
		strSQL.Append(" where RH68_ID_CARACTERISTICA = " & CaracteristicaId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

        dr("RH01_ID_PESSOA") = ProBanco(RH01_ID_PESSOA, eTipoValor.CHAVE)
        dr("TG09_ID_GRUPO_SANGUINEO") = ProBanco(TG09_ID_GRUPO_SANGUINEO, eTipoValor.CHAVE)
        dr("TG10_ID_RELIGIAO") = ProBanco(TG10_ID_RELIGIAO, eTipoValor.CHAVE)
        dr("TG42_ID_ETNIA") = ProBanco(TG42_ID_ETNIA, eTipoValor.CHAVE)
        dr("TG11_ID_RACA_COR") = ProBanco(TG11_ID_RACA_COR, eTipoValor.CHAVE)

        cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal CodigoPessoa as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH68_CARACTERISTICA")
		strSQL.Append(" where RH01_ID_PESSOA = " & CodigoPessoa)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH68_ID_CARACTERISTICA = DoBanco(dr("RH68_ID_CARACTERISTICA"), eTipoValor.CHAVE)
            RH01_ID_PESSOA = DoBanco(dr("RH01_ID_PESSOA"), eTipoValor.CHAVE)
            TG09_ID_GRUPO_SANGUINEO = DoBanco(dr("TG09_ID_GRUPO_SANGUINEO"), eTipoValor.CHAVE)
            TG10_ID_RELIGIAO = DoBanco(dr("TG10_ID_RELIGIAO"), eTipoValor.CHAVE)
            TG42_ID_ETNIA = DoBanco(dr("TG42_ID_ETNIA"), eTipoValor.CHAVE)
            TG11_ID_RACA_COR = DoBanco(dr("TG11_ID_RACA_COR"), eTipoValor.CHAVE)
        End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional CaracteristicaId as Integer = 0, Optional PessoaId as Integer = 0, Optional GrupoSanguineo as Integer = 0, Optional Religiao as Integer = 0, Optional Etinia as Integer = 0, Optional RacaCor as Integer = 0) as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH68_CARACTERISTICA")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH68_ID_CARACTERISTICA is not null")
		
		If CaracteristicaId > 0 then 
			strSQL.Append(" and RH68_ID_CARACTERISTICA = " & CaracteristicaId)
		End If
		
		If PessoaId> 0  then
			strSQL.Append(" and RH01_ID_PESSOA = " & PessoaId)
		End If
		
		If GrupoSanguineo> 0 then
			strSQL.Append(" and TG09_ID_GRUPO_SANGUINEO = " & GrupoSanguineo)
		End If
		
		If Religiao> 0 then
			strSQL.Append(" and TG10_ID_RELIGIAO = " & Religiao)
		End If
		
		If Etinia > 0 then
			strSQL.Append(" and TG42_ID_ETNIA = " & Etinia)
		End If
		
		If  RacaCor > 0  then
			strSQL.Append(" and TG11_ID_RACA_COR = " & RacaCor)
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH68_ID_CARACTERISTICA", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH68_ID_CARACTERISTICA as CODIGO, RH01_ID_PESSOA as DESCRICAO")
		strSQL.Append(" from RH68_CARACTERISTICA")
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
		
		strSQL.Append(" select max(RH68_ID_CARACTERISTICA) from RH68_CARACTERISTICA")

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
	Public Function Excluir(ByVal CaracteristicaId as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH68_CARACTERISTICA")
		strSQL.Append(" where RH68_ID_CARACTERISTICA = " & CaracteristicaId)

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

