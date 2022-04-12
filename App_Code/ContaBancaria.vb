Imports Microsoft.VisualBasic
Imports System.Data

Public Class ContaBancaria
	Private RH33_ID_CONTA_BANCARIA as Integer
	Private RH01_ID_PESSOA as Integer
	Private TG48_ID_BANCO as Integer
	Private CA04_ID_USUARIO as Integer
	Private RH33_NU_AGENCIA as String
	Private RH33_NU_DV_AGENCIA as String
	Private RH33_NU_CONTA as String
	Private RH33_NU_DV_CONTA as String
	Private RH33_TP_CONTA as String
	Private RH33_DH_CADASTRO as String
	Private RH33_DH_DESATIVACAO as String

	Public Property ContabancariaId() as Integer
		Get
			Return RH33_ID_CONTA_BANCARIA
		End Get
		Set(ByVal Value As Integer)
			RH33_ID_CONTA_BANCARIA = Value
		End Set
	End Property
	Public Property PessoaId() as Integer
		Get
			Return RH01_ID_PESSOA
		End Get
		Set(ByVal Value As Integer)
			RH01_ID_PESSOA = Value
		End Set
	End Property
	Public Property BancoId() as Integer
		Get
			Return TG48_ID_BANCO
		End Get
		Set(ByVal Value As Integer)
			TG48_ID_BANCO = Value
		End Set
	End Property
	Public Property UsuarioId() as Integer
		Get
			Return CA04_ID_USUARIO
		End Get
		Set(ByVal Value As Integer)
			CA04_ID_USUARIO = Value
		End Set
	End Property
	Public Property Agencia() as String
		Get
			Return RH33_NU_AGENCIA
		End Get
		Set(ByVal Value As String)
			RH33_NU_AGENCIA = Value
		End Set
	End Property
	Public Property DigitoAgencia() as String
		Get
			Return RH33_NU_DV_AGENCIA
		End Get
		Set(ByVal Value As String)
			RH33_NU_DV_AGENCIA = Value
		End Set
	End Property
	Public Property ContaCorrente() as String
		Get
			Return RH33_NU_CONTA
		End Get
		Set(ByVal Value As String)
			RH33_NU_CONTA = Value
		End Set
	End Property
	Public Property DigitoContaCorrente() as String
		Get
			Return RH33_NU_DV_CONTA
		End Get
		Set(ByVal Value As String)
			RH33_NU_DV_CONTA = Value
		End Set
	End Property
	Public Property TipoConta() as String
		Get
			Return RH33_TP_CONTA
		End Get
		Set(ByVal Value As String)
			RH33_TP_CONTA = Value
		End Set
	End Property
	Public Property DataHoraCadastro() as String
		Get
			Return RH33_DH_CADASTRO
		End Get
		Set(ByVal Value As String)
			RH33_DH_CADASTRO = Value
		End Set
	End Property
	Public Property DataHoraDesativacao() as String
		Get
			Return RH33_DH_DESATIVACAO
		End Get
		Set(ByVal Value As String)
			RH33_DH_DESATIVACAO = Value
		End Set
	End Property

	Public Sub New(Optional ByVal ContabancariaId as Integer = 0)
		If ContabancariaId > 0 Then
			Obter(ContabancariaId)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH33_CONTA_BANCARIA")
		strSQL.Append(" where RH33_ID_CONTA_BANCARIA = " & ContabancariaId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

        dr("RH01_ID_PESSOA") = ProBanco(RH01_ID_PESSOA, eTipoValor.CHAVE)
        dr("TG48_ID_BANCO") = ProBanco(TG48_ID_BANCO, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("RH33_NU_AGENCIA") = ProBanco(RH33_NU_AGENCIA, eTipoValor.TEXTO)
		dr("RH33_NU_DV_AGENCIA") = ProBanco(RH33_NU_DV_AGENCIA, eTipoValor.TEXTO)
		dr("RH33_NU_CONTA") = ProBanco(RH33_NU_CONTA, eTipoValor.TEXTO)
		dr("RH33_NU_DV_CONTA") = ProBanco(RH33_NU_DV_CONTA, eTipoValor.TEXTO)
		dr("RH33_TP_CONTA") = ProBanco(RH33_TP_CONTA, eTipoValor.TEXTO)
		dr("RH33_DH_CADASTRO") = ProBanco(RH33_DH_CADASTRO, eTipoValor.DATA)
		dr("RH33_DH_DESATIVACAO") = ProBanco(RH33_DH_DESATIVACAO, eTipoValor.DATA)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal ContabancariaId as Integer)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH33_CONTA_BANCARIA")
		strSQL.Append(" where RH33_ID_CONTA_BANCARIA = " & ContabancariaId)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH33_ID_CONTA_BANCARIA = DoBanco(dr("RH33_ID_CONTA_BANCARIA"), eTipoValor.CHAVE)
            RH01_ID_PESSOA = DoBanco(dr("RH01_ID_PESSOA"), eTipoValor.CHAVE)
            TG48_ID_BANCO = DoBanco(dr("TG48_ID_BANCO"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            RH33_NU_AGENCIA = DoBanco(dr("RH33_NU_AGENCIA"), eTipoValor.TEXTO)
			RH33_NU_DV_AGENCIA = DoBanco(dr("RH33_NU_DV_AGENCIA"), eTipoValor.TEXTO)
			RH33_NU_CONTA = DoBanco(dr("RH33_NU_CONTA"), eTipoValor.TEXTO)
			RH33_NU_DV_CONTA = DoBanco(dr("RH33_NU_DV_CONTA"), eTipoValor.TEXTO)
			RH33_TP_CONTA = DoBanco(dr("RH33_TP_CONTA"), eTipoValor.TEXTO)
			RH33_DH_CADASTRO = DoBanco(dr("RH33_DH_CADASTRO"), eTipoValor.DATA)
			RH33_DH_DESATIVACAO = DoBanco(dr("RH33_DH_DESATIVACAO"), eTipoValor.DATA)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional ContabancariaId as Integer = 0, Optional PessoaId as Integer = 0, Optional BancoId as Integer = 0, Optional UsuarioId as Integer = 0, Optional Agencia as String = "", Optional DigitoAgencia as String = "", Optional ContaCorrente as String = "", Optional DigitoContaCorrente as String = "", Optional TipoConta as String = "", Optional DataHoraCadastro as String = "", Optional DataHoraDesativacao as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH33_CONTA_BANCARIA")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH33_ID_CONTA_BANCARIA is not null")
		
		If ContabancariaId > 0 then 
			strSQL.Append(" and RH33_ID_CONTA_BANCARIA = " & ContabancariaId)
		End If
		
		If PessoaId> 0 then
			strSQL.Append(" and RH01_ID_PESSOA = " & PessoaId )
		End If
		
		If  BancoId> 0  then
			strSQL.Append(" and TG48_ID_BANCO = " & BancoId )
		End If
		
		If UsuarioId > 0 then 
			strSQL.Append(" and CA04_ID_USUARIO = " & UsuarioId)
		End If
		
		If Agencia <> "" then 
			strSQL.Append(" and upper(RH33_NU_AGENCIA) ='" & Agencia.toUpper & "'")
		End If
		
		If DigitoAgencia <> "" then 
			strSQL.Append(" and upper(RH33_NU_DV_AGENCIA) ='" & DigitoAgencia.toUpper & "'")
		End If
		
		If ContaCorrente <> "" then 
			strSQL.Append(" and upper(RH33_NU_CONTA) ='" & ContaCorrente.toUpper & "'")
		End If
		
		If DigitoContaCorrente <> "" then 
			strSQL.Append(" and upper(RH33_NU_DV_CONTA) ='" & DigitoContaCorrente.toUpper & "'")
		End If
		
		If TipoConta <> "" then 
			strSQL.Append(" and upper(RH33_TP_CONTA) ='" & TipoConta.toUpper & "'")
		End If
		
		If isDate(DataHoraCadastro) then 
			strSQL.Append(" and RH33_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
		End If
		
		If isDate(DataHoraDesativacao) then 
			strSQL.Append(" and RH33_DH_DESATIVACAO = Convert(DateTime, '" & DataHoraDesativacao & "', 103)")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH33_ID_CONTA_BANCARIA", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH33_ID_CONTA_BANCARIA as CODIGO, RH01_ID_PESSOA as DESCRICAO")
		strSQL.Append(" from RH33_CONTA_BANCARIA")
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
		
		strSQL.Append(" select max(RH33_ID_CONTA_BANCARIA) from RH33_CONTA_BANCARIA")

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
	Public Function Excluir(ByVal ContabancariaId as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH33_CONTA_BANCARIA")
		strSQL.Append(" where RH33_ID_CONTA_BANCARIA = " & ContabancariaId)

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

