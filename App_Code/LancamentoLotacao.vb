Imports Microsoft.VisualBasic
Imports System.Data

Public Class LancamentoLotacao
	Private RH45_ID_LANCAMENTO_LOTACAO as Integer
	Private RH36_ID_LOTACAO as Integer
	Private RH44_ID_LANCAMENTO_FREQ as Integer
	Private CA04_ID_USUARIO as Integer
	Private CA04_ID_USUARIO_ALT as Integer
	Private RH45_DH_CADASTRO as String
	Private RH45_DH_DESATIVACAO as String

	Public Property Codigo() as Integer
		Get
			Return RH45_ID_LANCAMENTO_LOTACAO
		End Get
		Set(ByVal Value As Integer)
			RH45_ID_LANCAMENTO_LOTACAO = Value
		End Set
	End Property
	Public Property IdLotacao() as Integer
		Get
			Return RH36_ID_LOTACAO
		End Get
		Set(ByVal Value As Integer)
			RH36_ID_LOTACAO = Value
		End Set
	End Property
	Public Property IdLancamento() as Integer
		Get
			Return RH44_ID_LANCAMENTO_FREQ
		End Get
		Set(ByVal Value As Integer)
			RH44_ID_LANCAMENTO_FREQ = Value
		End Set
	End Property
	Public Property IdUsuario() as Integer
		Get
			Return CA04_ID_USUARIO
		End Get
		Set(ByVal Value As Integer)
			CA04_ID_USUARIO = Value
		End Set
	End Property
	Public Property IdUsuarioAlteracao() as Integer
		Get
			Return CA04_ID_USUARIO_ALT
		End Get
		Set(ByVal Value As Integer)
			CA04_ID_USUARIO_ALT = Value
		End Set
	End Property
	Public Property DataHoraCadastro() as String
		Get
			Return RH45_DH_CADASTRO
		End Get
		Set(ByVal Value As String)
			RH45_DH_CADASTRO = Value
		End Set
	End Property
	Public Property DataHoraDesativacao() as String
		Get
			Return RH45_DH_DESATIVACAO
		End Get
		Set(ByVal Value As String)
			RH45_DH_DESATIVACAO = Value
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
		strSQL.Append(" from RH45_LANCAMENTO_LOTACAO")
		strSQL.Append(" where RH45_ID_LANCAMENTO_LOTACAO = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH36_ID_LOTACAO") = ProBanco(RH36_ID_LOTACAO, eTipoValor.CHAVE)
		dr("RH44_ID_LANCAMENTO_FREQ") = ProBanco(RH44_ID_LANCAMENTO_FREQ, eTipoValor.CHAVE)
		dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
		dr("CA04_ID_USUARIO_ALT") = ProBanco(CA04_ID_USUARIO_ALT, eTipoValor.CHAVE)
		dr("RH45_DH_CADASTRO") = ProBanco(RH45_DH_CADASTRO, eTipoValor.DATA)
		dr("RH45_DH_DESATIVACAO") = ProBanco(RH45_DH_DESATIVACAO, eTipoValor.DATA)

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
		strSQL.Append(" from RH45_LANCAMENTO_LOTACAO")
		strSQL.Append(" where RH45_ID_LANCAMENTO_LOTACAO = " & Codigo)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH45_ID_LANCAMENTO_LOTACAO = DoBanco(dr("RH45_ID_LANCAMENTO_LOTACAO"), eTipoValor.CHAVE)
			RH36_ID_LOTACAO = DoBanco(dr("RH36_ID_LOTACAO"), eTipoValor.CHAVE)
			RH44_ID_LANCAMENTO_FREQ = DoBanco(dr("RH44_ID_LANCAMENTO_FREQ"), eTipoValor.CHAVE)
			CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
			CA04_ID_USUARIO_ALT = DoBanco(dr("CA04_ID_USUARIO_ALT"), eTipoValor.CHAVE)
			RH45_DH_CADASTRO = DoBanco(dr("RH45_DH_CADASTRO"), eTipoValor.DATA)
			RH45_DH_DESATIVACAO = DoBanco(dr("RH45_DH_DESATIVACAO"), eTipoValor.DATA)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional Codigo as Integer = 0, Optional IdLotacao as Integer = 0, Optional IdLancamento as Integer = 0, Optional IdUsuario as Integer = 0, Optional IdUsuarioAlteracao as Integer = 0, Optional DataHoraCadastro as String = "", Optional DataHoraDesativacao as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH45_LANCAMENTO_LOTACAO")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH45_ID_LANCAMENTO_LOTACAO is not null")
		
		If Codigo > 0 then 
			strSQL.Append(" and RH45_ID_LANCAMENTO_LOTACAO = " & Codigo)
		End If
		
		If IdLotacao > 0 then 
			strSQL.Append(" and RH36_ID_LOTACAO = " & IdLotacao)
		End If
		
		If IdLancamento > 0 then 
			strSQL.Append(" and RH44_ID_LANCAMENTO_FREQ = " & IdLancamento)
		End If
		
		If IdUsuario > 0 then 
			strSQL.Append(" and CA04_ID_USUARIO = " & IdUsuario)
		End If
		
		If IdUsuarioAlteracao > 0 then 
			strSQL.Append(" and CA04_ID_USUARIO_ALT = " & IdUsuarioAlteracao)
		End If
		
		If isDate(DataHoraCadastro) then 
			strSQL.Append(" and RH45_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
		End If
		
		If isDate(DataHoraDesativacao) then 
			strSQL.Append(" and RH45_DH_DESATIVACAO = Convert(DateTime, '" & DataHoraDesativacao & "', 103)")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH45_ID_LANCAMENTO_LOTACAO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH45_ID_LANCAMENTO_LOTACAO as CODIGO, RH36_ID_LOTACAO as DESCRICAO")
		strSQL.Append(" from RH45_LANCAMENTO_LOTACAO")
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
		
		strSQL.Append(" select max(RH45_ID_LANCAMENTO_LOTACAO) from RH45_LANCAMENTO_LOTACAO")

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
		strSQL.Append(" from RH45_LANCAMENTO_LOTACAO")
		strSQL.Append(" where RH45_ID_LANCAMENTO_LOTACAO = " & Codigo)

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

