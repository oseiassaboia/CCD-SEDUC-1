Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Runtime.InteropServices.WindowsRuntime

Public Class ServidorCargaHoraria
    Implements IDisposable
	Private RH78_ID_SERVIDOR_CARGA_HORARIA as Integer
	Private RH02_ID_SERVIDOR as Integer
	Private RH77_ID_CARGA_HORARIA as Integer
	Private RH78_DT_INICIO_VIGENCIA as String
	Private RH78_DT_TERMINO_VIGENCIA as String
	Private CA04_ID_USUARIO as Integer
	Private RH78_DH_CADASTRO as String
    Private RH78_DH_DESATIVACAO as String
	Private CA04_ID_USUARIO_ALT As Integer
	Private RH88_ID_PERIODO As Integer

	Public Property IdServidorCargaHoraria() as Integer
		Get
			Return RH78_ID_SERVIDOR_CARGA_HORARIA
		End Get
		Set(ByVal Value As Integer)
			RH78_ID_SERVIDOR_CARGA_HORARIA = Value
		End Set
	End Property
	Public Property IdServidor() as Integer
		Get
			Return RH02_ID_SERVIDOR
		End Get
		Set(ByVal Value As Integer)
			RH02_ID_SERVIDOR = Value
		End Set
	End Property
	Public Property IdCargaHoraria() as Integer
		Get
			Return RH77_ID_CARGA_HORARIA
		End Get
		Set(ByVal Value As Integer)
		    RH77_ID_CARGA_HORARIA = Value
		End Set
	End Property
	Public Property dataInicioVigencia() as String
		Get
			Return RH78_DT_INICIO_VIGENCIA
		End Get
		Set(ByVal Value As String)
			RH78_DT_INICIO_VIGENCIA = Value
		End Set
	End Property
	Public Property DataTerminoVigencia() as String
		Get
			Return RH78_DT_TERMINO_VIGENCIA
		End Get
		Set(ByVal Value As String)
			RH78_DT_TERMINO_VIGENCIA = Value
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
	Public Property DataHoraCadastro() as String
		Get
			Return RH78_DH_CADASTRO
		End Get
		Set(ByVal Value As String)
			RH78_DH_CADASTRO = Value
		End Set
	End Property

    Public Property  IdUsuarioAlteracao() As Integer
    get
        Return CA04_ID_USUARIO_ALT
    End Get
        Set(value As Integer)
            CA04_ID_USUARIO_ALT = value
        End Set
    End Property
	Public Property DataHoraDesativacao() As String
		Get
			Return RH78_DH_DESATIVACAO
		End Get
		Set(value As String)
			RH78_DH_DESATIVACAO = value
		End Set
	End Property
	Public Property Periodo() As Integer
		Get
			Return RH88_ID_PERIODO
		End Get
		Set(value As Integer)
			RH88_ID_PERIODO = value
		End Set
	End Property

	Public Sub New(Optional ByVal IdServidorCargaHoraria as integer = 0)
		If IdServidorCargaHoraria > 0 Then
			Obter(IdServidorCargaHoraria)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH78_SERVIDOR_CARGA_HORARIA")
		strSQL.Append(" where RH78_ID_SERVIDOR_CARGA_HORARIA = " & IdServidorCargaHoraria)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH02_ID_SERVIDOR") = ProBanco(RH02_ID_SERVIDOR, eTipoValor.chave)
		dr("RH77_ID_CARGA_HORARIA") = ProBanco(RH77_ID_CARGA_HORARIA, eTipoValor.chave)
		dr("RH78_DT_INICIO_VIGENCIA") = ProBanco(RH78_DT_INICIO_VIGENCIA, eTipoValor.data)
		dr("RH78_DT_TERMINO_VIGENCIA") = ProBanco(RH78_DT_TERMINO_VIGENCIA, eTipoValor.data)
		dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
		dr("RH78_DH_CADASTRO") = ProBanco(RH78_DH_CADASTRO, eTipoValor.DATA_COMPLETA)
	    dr("RH78_DH_DESATIVACAO") = ProBanco(RH78_DH_DESATIVACAO, eTipoValor.DATA_COMPLETA)
		dr("CA04_ID_USUARIO_ALT") = ProBanco(CA04_ID_USUARIO_ALT, eTipoValor.CHAVE)
		dr("RH88_ID_PERIODO") = ProBanco(RH88_ID_PERIODO, eTipoValor.CHAVE)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal IdServidorCargaHoraria as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH78_SERVIDOR_CARGA_HORARIA")
		strSQL.Append(" where RH78_ID_SERVIDOR_CARGA_HORARIA = " & IdServidorCargaHoraria)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH78_ID_SERVIDOR_CARGA_HORARIA = DoBanco(dr("RH78_ID_SERVIDOR_CARGA_HORARIA"), eTipoValor.CHAVE)
			RH02_ID_SERVIDOR = DoBanco(dr("RH02_ID_SERVIDOR"), eTipoValor.chave)
			RH77_ID_CARGA_HORARIA = DoBanco(dr("RH77_ID_CARGA_HORARIA"), eTipoValor.chave)
			RH78_DT_INICIO_VIGENCIA = DoBanco(dr("RH78_DT_INICIO_VIGENCIA"), eTipoValor.TEXTO)
			RH78_DT_TERMINO_VIGENCIA = DoBanco(dr("RH78_DT_TERMINO_VIGENCIA"), eTipoValor.TEXTO)
			CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.chave)
			RH78_DH_CADASTRO = DoBanco(dr("RH78_DH_CADASTRO"), eTipoValor.DATA_COMPLETA)
		    RH78_DH_DESATIVACAO = DoBanco(dr("RH78_DH_DESATIVACAO"), eTipoValor.DATA_COMPLETA)
			CA04_ID_USUARIO_ALT = DoBanco(dr("CA04_ID_USUARIO_ALT"), eTipoValor.CHAVE)
			RH88_ID_PERIODO = DoBanco(dr("RH88_ID_PERIODO"), eTipoValor.CHAVE)

		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort As String = "", Optional IdServidorCargaHoraria As Integer = 0, Optional IdServidor As Integer = 0, Optional IdCargaHoraria As Integer = 0, Optional dataInicioVigencia As String = "", Optional DataTerminoVigencia As String = "", Optional IdUsuario As Integer = 0, Optional DataCadastro As String = "", Optional ByVal CodigoPessoa As Integer = 0, Optional ByVal CodigoServidor As Integer = 0, Optional ByVal CargaHorariaAtiva As Boolean = True, Optional ByVal CodigoUsuarioAlteracao As Integer = 0, Optional ByVal ServidorAtivo As Boolean = True, Optional Periodo As String = "") As DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append(" select *, RH79_NM_TIPO_CARGA_HORARIA+' - CONTRATAÇÃO: '+ case when  RH77_TP_CONTRATACAO = 'E' then 'EFETIVO' WHEN RH77_TP_CONTRATACAO ='C' THEN  'CONTRATO' END  + ' - SERVIDOR: '+isnull(RH02_CD_MATRICULA,'')+ ' - '+RH01_NM_PESSOA as DESCRICAO,case when isnull(RH78_DH_DESATIVACAO,0)   = ''  then 0 else 1 end as Desativado  ")
		strSQL.Append(" from RH78_SERVIDOR_CARGA_HORARIA ServidorCargaHoraria")
		strSQL.Append(" inner join RH02_SERVIDOR SERVIDOR  on Servidor.RH02_ID_SERVIDOR = ServidorCargaHoraria.RH02_ID_SERVIDOR ")
		strSQL.Append(" inner join RH01_PESSOA Pessoa  on Pessoa.RH01_ID_Pessoa = Servidor.RH01_ID_Pessoa ")
		strSQL.Append(" inner join RH77_CARGA_HORARIA CargaHoraria  on CargaHoraria.RH77_ID_CARGA_HORARIA = ServidorCargaHoraria.RH77_ID_CARGA_HORARIA ")
		strSQL.Append(" inner join RH79_TIPO_CARGA_HORARIA TipoCargaHoraria  on TipoCargaHoraria.RH79_ID_TIPO_CARGA_HORARIA = CargaHoraria.RH79_ID_TIPO_CARGA_HORARIA ")
		strSQL.Append(" inner join	RH88_PERIODO	rh88	on	ServidorCargaHoraria.RH88_ID_PERIODO = rh88.RH88_ID_PERIODO ")
		strSQL.Append(" where RH78_ID_SERVIDOR_CARGA_HORARIA is not null ")
		strSQL.Append(" and Servidor.RH07_ID_SITUACAO_SERVIDOR in (1,11,10 )")

		'strSQL.Append(" and ServidorCargaHoraria.RH78_DH_DESATIVACAO is null")

		If Periodo <> "" Then
			strSQL.Append(" and rh88.RH88_NM_PERIODO = " & Periodo)
		End If


		If CargaHorariaAtiva Then
			strSQL.Append(" and isnull(ServidorCargaHoraria.RH78_DH_DESATIVACAO,'') = '' ")
		End If

		If IdServidorCargaHoraria > 0 Then
			strSQL.Append(" and ServidorCargaHoraria.RH78_ID_SERVIDOR_CARGA_HORARIA = " & IdServidorCargaHoraria)
		End If

		If IdServidor > 0 Then
			strSQL.Append(" and SERVIDOR.RH02_ID_SERVIDOR = " & IdServidor)
		End If

		'If CargaHorariaVigente Then ' Campo RH78_DT_TERMINO_VIGENCIA Nulo, por isso não é verificado
		'    strSQL.Append(" and RH78_DT_TERMINO_VIGENCIA > getdate() ")
		'End If

		If IdCargaHoraria > 0 Then
			strSQL.Append(" and CargaHoraria.RH77_ID_CARGA_HORARIA = " & IdCargaHoraria)
		End If

		If dataInicioVigencia <> "" Then
			strSQL.Append(" and upper(RH78_DT_INICIO_VIGENCIA) like '%" & dataInicioVigencia.ToUpper & "%'")
		End If

		If DataTerminoVigencia <> "" Then
			strSQL.Append(" and upper(RH78_DT_TERMINO_VIGENCIA) like '%" & DataTerminoVigencia.ToUpper & "%'")
		End If

		If IdUsuario > 0 Then
			strSQL.Append(" and CA04_ID_USUARIO = " & IdUsuario)
		End If

		If DataCadastro <> "" Then
			strSQL.Append(" and upper(RH78_DH_CADASTRO) like '%" & RH78_DH_CADASTRO.ToUpper & "%'")
		End If

		If CodigoPessoa > 0 Then
			strSQL.Append(" and Pessoa.RH01_ID_Pessoa = " & CodigoPessoa)
		End If

		If CodigoUsuarioAlteracao > 0 Then
			strSQL.Append(" and ServidorCargaHoraria.CA04_ID_USUARIO_ALT = " & CodigoUsuarioAlteracao)
		End If

		strSQL.Append(" Order By " & IIf(Sort = "", "RH88_NM_PERIODO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH78_ID_SERVIDOR_CARGA_HORARIA as CODIGO, RH02_ID_SERVIDOR as DESCRICAO")
		strSQL.Append(" from RH78_SERVIDOR_CARGA_HORARIA")
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
		
		strSQL.Append(" select max(RH78_ID_SERVIDOR_CARGA_HORARIA) from RH78_SERVIDOR_CARGA_HORARIA")

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
	Public Function Excluir(ByVal IdServidorCargaHoraria as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH78_SERVIDOR_CARGA_HORARIA")
		strSQL.Append(" where RH78_ID_SERVIDOR_CARGA_HORARIA = " & IdServidorCargaHoraria)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function


#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

'******************************************************************************
'*                                 26/02/2019                                 *
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

