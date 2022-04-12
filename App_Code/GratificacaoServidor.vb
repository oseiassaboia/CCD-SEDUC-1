Imports Microsoft.VisualBasic
Imports System.Data

Public Class GratificacaoServidor
	Implements IDisposable

	Private RH91_ID_GRATIFICACAO_SERVIDOR As Integer
	Private RH02_ID_SERVIDOR As Integer
	Private RH90_ID_TIPO_GRATIFICACAO As Integer
	Private CA04_ID_USUARIO As Integer
	Private RH91_DH_CADASTRO As String
	Private CA04_ID_USUARIO_ALT As Integer
	Private RH91_DH_DESATIVACAO As String
	Private RH88_ID_PERIODO As Integer
	Private disposedValue As Boolean

	Public Property IdGratificacao() As Integer
		Get
			Return RH91_ID_GRATIFICACAO_SERVIDOR
		End Get
		Set(ByVal Value As Integer)
			RH91_ID_GRATIFICACAO_SERVIDOR = Value
		End Set
	End Property
	Public Property IdServidor() As Integer
		Get
			Return RH02_ID_SERVIDOR
		End Get
		Set(ByVal Value As Integer)
			RH02_ID_SERVIDOR = Value
		End Set
	End Property
	Public Property IdTipoGratificacao() As Integer
		Get
			Return RH90_ID_TIPO_GRATIFICACAO
		End Get
		Set(ByVal Value As Integer)
			RH90_ID_TIPO_GRATIFICACAO = Value
		End Set
	End Property
	Public Property IdUsuario() As Integer
		Get
			Return CA04_ID_USUARIO
		End Get
		Set(ByVal Value As Integer)
			CA04_ID_USUARIO = Value
		End Set
	End Property
	Public Property DataHoraCadastro() As String
		Get
			Return RH91_DH_CADASTRO
		End Get
		Set(ByVal Value As String)
			RH91_DH_CADASTRO = Value
		End Set
	End Property
	Public Property IdUsuarioAlteracao() As Integer
		Get
			Return CA04_ID_USUARIO_ALT
		End Get
		Set(ByVal Value As Integer)
			CA04_ID_USUARIO_ALT = Value
		End Set
	End Property
	Public Property DataHoraDesativacao() As String
		Get
			Return RH91_DH_DESATIVACAO
		End Get
		Set(ByVal Value As String)
			RH91_DH_DESATIVACAO = Value
		End Set
	End Property
	Public Property IdPeriodo() As Integer
		Get
			Return RH88_ID_PERIODO
		End Get
		Set(ByVal Value As Integer)
			RH88_ID_PERIODO = Value
		End Set
	End Property

	Public Sub New(Optional ByVal IdGratificacao As Integer = 0)
		If IdGratificacao > 0 Then
			Obter(IdGratificacao)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder

		strSQL.Append("Then select * ")
		strSQL.Append(" from RH91_GRATIFICACAO_SERVIDOR")
		strSQL.Append(" where RH91_ID_GRATIFICACAO_SERVIDOR = " & IdGratificacao)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH02_ID_SERVIDOR") = ProBanco(RH02_ID_SERVIDOR, eTipoValor.CHAVE)
		dr("RH90_ID_TIPO_GRATIFICACAO") = ProBanco(RH90_ID_TIPO_GRATIFICACAO, eTipoValor.CHAVE)
		dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
		dr("RH91_DH_CADASTRO") = ProBanco(RH91_DH_CADASTRO, eTipoValor.DATA)
		dr("CA04_ID_USUARIO_ALT") = ProBanco(CA04_ID_USUARIO_ALT, eTipoValor.CHAVE)
		dr("RH91_DH_DESATIVACAO") = ProBanco(RH91_DH_DESATIVACAO, eTipoValor.DATA)
		dr("RH88_ID_PERIODO") = ProBanco(RH88_ID_PERIODO, eTipoValor.CHAVE)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal IdGratificacao As String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder

		strSQL.Append(" select * ")
		strSQL.Append(" from RH91_GRATIFICACAO_SERVIDOR")
		strSQL.Append(" where RH91_ID_GRATIFICACAO_SERVIDOR = " & IdGratificacao)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)

			RH91_ID_GRATIFICACAO_SERVIDOR = DoBanco(dr("RH91_ID_GRATIFICACAO_SERVIDOR"), eTipoValor.CHAVE)
			RH02_ID_SERVIDOR = DoBanco(dr("RH02_ID_SERVIDOR"), eTipoValor.CHAVE)
			RH90_ID_TIPO_GRATIFICACAO = DoBanco(dr("RH90_ID_TIPO_GRATIFICACAO"), eTipoValor.CHAVE)
			CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
			RH91_DH_CADASTRO = DoBanco(dr("RH91_DH_CADASTRO"), eTipoValor.DATA)
			CA04_ID_USUARIO_ALT = DoBanco(dr("CA04_ID_USUARIO_ALT"), eTipoValor.CHAVE)
			RH91_DH_DESATIVACAO = DoBanco(dr("RH91_DH_DESATIVACAO"), eTipoValor.DATA)
			RH88_ID_PERIODO = DoBanco(dr("RH88_ID_PERIODO"), eTipoValor.CHAVE)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort As String = "", Optional IdGratificacao As Integer = 0, Optional IdServidor As Integer = 0, Optional IdTipoGratificacao As Integer = 0,
								Optional IdUsuario As Integer = 0, Optional DataHoraCadastro As String = "", Optional IdUsuarioAlteracao As Integer = 0,
								Optional DataHoraDesativacao As String = "", Optional IdPeriodo As Integer = 0, Optional ByVal Matricula As String = "", Optional ByVal Nome As String = "") As DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append(" select * ")
		strSQL.Append(" from RH91_GRATIFICACAO_SERVIDOR RH91 ")
		strSQL.Append(" right join RH90_TIPO_GRATIFICACAO rh90 on RH91.RH90_ID_TIPO_GRATIFICACAO = RH90.RH90_ID_TIPO_GRATIFICACAO ")
		strSQL.Append(" right join RH02_SERVIDOR RH02 on RH91.RH02_ID_SERVIDOR = RH02.RH02_ID_SERVIDOR ")
		strSQL.Append(" right join RH01_PESSOA RH01 on RH01.RH01_ID_PESSOA = RH02.RH01_ID_PESSOA ")
		'strSQL.Append(" where RH91_ID_GRATIFICACAO_SERVIDOR is not null")
		strSQL.Append(" where RH02.RH07_ID_SITUACAO_SERVIDOR IN (1,10,11)")
		strSQL.Append(" AND RH04_ID_ORGAO IN (5,6,7,8) ")

		If Matricula <> "" Then
			strSQL.Append(" and rh02.RH02_CD_MATRICULA = '" & Matricula & "'")
		End If

		If Nome <> "" Then
			strSQL.Append(" And upper(rh01.RH01_NM_PESSOA) Like '%" & Nome.ToUpper & "%' collate sql_latin1_general_cp1251_cs_as")
		End If


		If IdGratificacao > 0 Then
			strSQL.Append(" and RH91_ID_GRATIFICACAO_SERVIDOR = " & IdGratificacao)
		End If

		If IdServidor > 0 Then
			strSQL.Append(" and RH02_ID_SERVIDOR = " & IdServidor)
		End If

		If IdTipoGratificacao > 0 Then
			strSQL.Append(" and RH90_ID_TIPO_GRATIFICACAO = " & IdTipoGratificacao)
		End If

		If IdUsuario > 0 Then
			strSQL.Append(" and CA04_ID_USUARIO = " & IdUsuario)
		End If

		If IsDate(DataHoraCadastro) Then
			strSQL.Append(" and RH91_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
		End If

		If IdUsuarioAlteracao > 0 Then
			strSQL.Append(" and CA04_ID_USUARIO_ALT = " & IdUsuarioAlteracao)
		End If

		If IsDate(DataHoraDesativacao) Then
			strSQL.Append(" and RH91_DH_DESATIVACAO = Convert(DateTime, '" & DataHoraDesativacao & "', 103)")
		End If

		If IdPeriodo > 0 Then
			strSQL.Append(" and RH88_ID_PERIODO = " & IdPeriodo)
		End If

		strSQL.Append(" Order By " & IIf(Sort = "", "RH91_ID_GRATIFICACAO_SERVIDOR", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() As DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder

		strSQL.Append(" select RH91_ID_GRATIFICACAO_SERVIDOR as CODIGO, RH02_ID_SERVIDOR as DESCRICAO")
		strSQL.Append(" from RH91_GRATIFICACAO_SERVIDOR")
		strSQL.Append(" order by 2 ")

		dt = cnn.AbrirDataTable(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return dt
	End Function

	Public Function ObterUltimo() As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim CodigoUltimo As Integer

		strSQL.Append(" select max(RH91_ID_GRATIFICACAO_SERVIDOR) from RH91_GRATIFICACAO_SERVIDOR")

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
	Public Function Excluir(ByVal IdGratificacao As String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer

		strSQL.Append(" delete ")
		strSQL.Append(" from RH91_GRATIFICACAO_SERVIDOR")
		strSQL.Append(" where RH91_ID_GRATIFICACAO_SERVIDOR = " & IdGratificacao)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

	Protected Overridable Sub Dispose(disposing As Boolean)
		If Not disposedValue Then
			If disposing Then
				' Tarefa pendente: descartar o estado gerenciado (objetos gerenciados)
			End If

			' Tarefa pendente: liberar recursos não gerenciados (objetos não gerenciados) e substituir o finalizador
			' Tarefa pendente: definir campos grandes como nulos
			disposedValue = True
		End If
	End Sub

	' ' Tarefa pendente: substituir o finalizador somente se 'Dispose(disposing As Boolean)' tiver o código para liberar recursos não gerenciados
	' Protected Overrides Sub Finalize()
	'     ' Não altere este código. Coloque o código de limpeza no método 'Dispose(disposing As Boolean)'
	'     Dispose(disposing:=False)
	'     MyBase.Finalize()
	' End Sub

	Public Sub Dispose() Implements IDisposable.Dispose
		' Não altere este código. Coloque o código de limpeza no método 'Dispose(disposing As Boolean)'
		Dispose(disposing:=True)
		GC.SuppressFinalize(Me)
	End Sub
End Class

'******************************************************************************
'*                                 27/07/2020                                 *
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

