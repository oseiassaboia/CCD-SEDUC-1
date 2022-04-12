Imports Microsoft.VisualBasic
Imports System.Data

Public Class FrequenciaServidor
    Implements  IDisposable

	Private RH18_ID_FREQ_SERVIDOR as Integer
	Private RH14_ID_LOTACAO_SERVIDOR as Integer
	Private RH24_ID_FREQ_LOTACAO as Integer
	Private RH23_ID_TIPO_REGISTRO as Integer
	Private CA04_ID_USUARIO as Integer
	Private RH18_DT_FREQUENCIA as String
	Private RH18_DH_ENVIO as String
	Private RH18_DH_RETIFICACAO as String
	Private RH18_DH_CADASTRO as String
	Private RH24_ST_FREQ_LOTACAO as String
	Private RH24_DH_ST_FREQ_LOTACAO as String

	Public Property Codigo() as Integer
		Get
			Return RH18_ID_FREQ_SERVIDOR
		End Get
		Set(ByVal Value As Integer)
			RH18_ID_FREQ_SERVIDOR = Value
		End Set
	End Property
	Public Property IdLotacaoServidor() as Integer
		Get
			Return RH14_ID_LOTACAO_SERVIDOR
		End Get
		Set(ByVal Value As Integer)
			RH14_ID_LOTACAO_SERVIDOR = Value
		End Set
	End Property
	Public Property IdFrequenciaLotacao() as Integer
		Get
			Return RH24_ID_FREQ_LOTACAO
		End Get
		Set(ByVal Value As Integer)
			RH24_ID_FREQ_LOTACAO = Value
		End Set
	End Property
	Public Property idTipoRegistro() as Integer
		Get
			Return RH23_ID_TIPO_REGISTRO
		End Get
		Set(ByVal Value As Integer)
			RH23_ID_TIPO_REGISTRO = Value
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
	Public Property DataFrequencia() as String
		Get
			Return RH18_DT_FREQUENCIA
		End Get
		Set(ByVal Value As String)
			RH18_DT_FREQUENCIA = Value
		End Set
	End Property
	Public Property DataHoraEnvio() as String
		Get
			Return RH18_DH_ENVIO
		End Get
		Set(ByVal Value As String)
			RH18_DH_ENVIO = Value
		End Set
	End Property
	Public Property DataHoraRetificacao() as String
		Get
			Return RH18_DH_RETIFICACAO
		End Get
		Set(ByVal Value As String)
			RH18_DH_RETIFICACAO = Value
		End Set
	End Property
	Public Property DataHoraCadastro() as String
		Get
			Return RH18_DH_CADASTRO
		End Get
		Set(ByVal Value As String)
			RH18_DH_CADASTRO = Value
		End Set
	End Property
	Public Property SituacaoFrequencia() as String
		Get
			Return RH24_ST_FREQ_LOTACAO
		End Get
		Set(ByVal Value As String)
			RH24_ST_FREQ_LOTACAO = Value
		End Set
	End Property
	Public Property DataHoraSituacaoFrequencia() as String
		Get
			Return RH24_DH_ST_FREQ_LOTACAO
		End Get
		Set(ByVal Value As String)
			RH24_DH_ST_FREQ_LOTACAO = Value
		End Set
	End Property

	Public Sub New(Optional ByVal Codigo as integer = 0, optional CodigoLotacaoServidor as Integer = 0, optional CodigoFrequenciaLotacao As Integer = 0, Optional Data As String = "")
		If Codigo > 0 or CodigoLotacaoServidor > 0 or CodigoFrequenciaLotacao > 0 or  Data <> "" Then
			Obter(Codigo,CodigoLotacaoServidor,CodigoFrequenciaLotacao,Data )
		End If
	End Sub

    Public Function Salvar(Optional ByRef tran As Transacao = Nothing) As Boolean
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH18_FREQ_SERVIDOR")
        strSQL.Append(" where RH18_ID_FREQ_SERVIDOR = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH14_ID_LOTACAO_SERVIDOR") = ProBanco(RH14_ID_LOTACAO_SERVIDOR, eTipoValor.CHAVE)
        dr("RH24_ID_FREQ_LOTACAO") = ProBanco(RH24_ID_FREQ_LOTACAO, eTipoValor.CHAVE)
        dr("RH23_ID_TIPO_REGISTRO") = ProBanco(RH23_ID_TIPO_REGISTRO, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("RH18_DT_FREQUENCIA") = ProBanco(RH18_DT_FREQUENCIA, eTipoValor.TEXTO)
        dr("RH18_DH_ENVIO") = ProBanco(RH18_DH_ENVIO, eTipoValor.DATA)
        dr("RH18_DH_RETIFICACAO") = ProBanco(RH18_DH_RETIFICACAO, eTipoValor.DATA)
        dr("RH18_DH_CADASTRO") = ProBanco(RH18_DH_CADASTRO, eTipoValor.DATA)
        dr("RH24_ST_FREQ_LOTACAO") = ProBanco(RH24_ST_FREQ_LOTACAO, eTipoValor.TEXTO)
        dr("RH24_DH_ST_FREQ_LOTACAO") = ProBanco(RH24_DH_ST_FREQ_LOTACAO, eTipoValor.DATA)

        Dim flag As Boolean = True
        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing

        Return flag
    End Function

    Public Sub Obter(Optional ByVal Codigo as integer = 0, optional CodigoLotacaoServidor as Integer = 0, optional CodigoFrequenciaLotacao As Integer = 0, Optional Data As String = "")
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH18_FREQ_SERVIDOR")
		strSQL.Append(" where RH18_ID_FREQ_SERVIDOR is not null" )

        If Codigo> 0 Then
            strSQL.Append(" and RH18_ID_FREQ_SERVIDOR = " & Codigo )
        End If

        If CodigoLotacaoServidor > 0 Then
            strSQL.Append(" and RH14_ID_LOTACAO_SERVIDOR = " & CodigoLotacaoServidor )
        End If

        If CodigoFrequenciaLotacao > 0 Then
            strSQL.Append(" and RH24_ID_FREQ_LOTACAO = " & CodigoFrequenciaLotacao )
        End If

        If Data <> "" Then
            strSQL.Append(" and convert(varchar,RH18_DT_FREQUENCIA,103) = '" & Data & "'" )
        End If

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH18_ID_FREQ_SERVIDOR = DoBanco(dr("RH18_ID_FREQ_SERVIDOR"), eTipoValor.CHAVE)
			RH14_ID_LOTACAO_SERVIDOR = DoBanco(dr("RH14_ID_LOTACAO_SERVIDOR"), eTipoValor.CHAVE)
			RH24_ID_FREQ_LOTACAO = DoBanco(dr("RH24_ID_FREQ_LOTACAO"), eTipoValor.CHAVE)
			RH23_ID_TIPO_REGISTRO = DoBanco(dr("RH23_ID_TIPO_REGISTRO"), eTipoValor.CHAVE)
			CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
			RH18_DT_FREQUENCIA = DoBanco(dr("RH18_DT_FREQUENCIA"), eTipoValor.TEXTO)
			RH18_DH_ENVIO = DoBanco(dr("RH18_DH_ENVIO"), eTipoValor.DATA)
			RH18_DH_RETIFICACAO = DoBanco(dr("RH18_DH_RETIFICACAO"), eTipoValor.DATA)
			RH18_DH_CADASTRO = DoBanco(dr("RH18_DH_CADASTRO"), eTipoValor.DATA)
			RH24_ST_FREQ_LOTACAO = DoBanco(dr("RH24_ST_FREQ_LOTACAO"), eTipoValor.TEXTO)
			RH24_DH_ST_FREQ_LOTACAO = DoBanco(dr("RH24_DH_ST_FREQ_LOTACAO"), eTipoValor.DATA)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

    Public Function buscarFrequenciaAbertaServidor(codigoServidor As Integer) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select RH18_ID_FREQ_SERVIDOR, RH14_ID_LOTACAO_SERVIDOR, RH24.RH24_ID_FREQ_LOTACAO, RH17_ID_PERIODO_FREQ, RH36_ID_LOTACAO   ")
        strSQL.Append(" from RH18_FREQ_SERVIDOR as rh18 ")
        strSQL.Append(" inner join RH24_FREQ_LOTACAO as rh24 on rh24.RH24_ID_FREQ_LOTACAO = rh18.RH24_ID_FREQ_LOTACAO ")
        strSQL.Append(" where RH14_ID_LOTACAO_SERVIDOR = " & codigoServidor)
        strSQL.Append(" and rh24.RH24_ST_FREQ_LOTACAO in (1,2,5) ")

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Sub atualizarLotacaoFrequencia(codigo As Integer, frquenciaLotacaoCod As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH18_FREQ_SERVIDOR")
        strSQL.Append(" where RH18_ID_FREQ_SERVIDOR = " & codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH24_ID_FREQ_LOTACAO") = ProBanco(frquenciaLotacaoCod, eTipoValor.CHAVE)

        Dim flag As Boolean = True
        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing


    End Sub


    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional IdLotacaoServidor As Integer = 0, Optional IdFrequenciaLotacao As Integer = 0, Optional idTipoRegistro As Integer = 0, Optional IdUsuario As Integer = 0, Optional DataFrequencia As String = "", Optional DataHoraEnvio As String = "", Optional DataHoraRetificacao As String = "", Optional DataHoraCadastro As String = "", Optional SituacaoFrequencia As String = "", Optional DataHoraSituacaoFrequencia As String = "", Optional ByVal Lotacao As Integer = 0, Optional ByVal Periodo As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH18_FREQ_SERVIDOR RH18")
        strSQL.Append(" INNER join RH24_FREQ_LOTACAO RH24 on RH24.RH24_ID_FREQ_LOTACAO = RH18.RH24_ID_FREQ_LOTACAO ")
        strSQL.Append(" where RH18_ID_FREQ_SERVIDOR is not null")

        If Codigo > 0 Then
            strSQL.Append(" and RH18_ID_FREQ_SERVIDOR = " & Codigo)
        End If

        If Lotacao > 0 Then
            strSQL.Append(" and RH24.RH36_ID_LOTACAO = " & Lotacao)

        End If

        If IdLotacaoServidor > 0 Then
            strSQL.Append(" and RH14_ID_LOTACAO_SERVIDOR = " & IdLotacaoServidor)
        End If

        If IdFrequenciaLotacao > 0 Then
            strSQL.Append(" and RH24.RH24_ID_FREQ_LOTACAO = " & IdFrequenciaLotacao)
        End If

        If idTipoRegistro > 0 Then
            strSQL.Append(" and RH23_ID_TIPO_REGISTRO = " & idTipoRegistro)
        End If

        If IdUsuario > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & IdUsuario)
        End If

        If DataFrequencia <> "" Then
            strSQL.Append(" and upper(RH18_DT_FREQUENCIA) like '%" & DataFrequencia.ToUpper & "%'")
        End If

        If IsDate(DataHoraEnvio) Then
            strSQL.Append(" and RH18_DH_ENVIO = Convert(DateTime, '" & DataHoraEnvio & "', 103)")
        End If

        If IsDate(DataHoraRetificacao) Then
            strSQL.Append(" and RH18_DH_RETIFICACAO = Convert(DateTime, '" & DataHoraRetificacao & "', 103)")
        End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" and RH18_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
        End If

        If SituacaoFrequencia <> "" Then
            strSQL.Append(" and upper(RH24_ST_FREQ_LOTACAO) like '%" & SituacaoFrequencia.ToUpper & "%'")
        End If

        If IsDate(DataHoraSituacaoFrequencia) Then
            strSQL.Append(" and RH24_DH_ST_FREQ_LOTACAO = Convert(DateTime, '" & DataHoraSituacaoFrequencia & "', 103)")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH18_ID_FREQ_SERVIDOR", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH18_ID_FREQ_SERVIDOR as CODIGO, RH14_ID_LOTACAO_SERVIDOR as DESCRICAO")
		strSQL.Append(" from RH18_FREQ_SERVIDOR")
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
		
		strSQL.Append(" select max(RH18_ID_FREQ_SERVIDOR) from RH18_FREQ_SERVIDOR")

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
		strSQL.Append(" from RH18_FREQ_SERVIDOR")
		strSQL.Append(" where RH18_ID_FREQ_SERVIDOR = " & Codigo)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

    Public Function excluir(codigoLotacaoFrequencia As Integer, codigoServidorLotacao As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from RH18_FREQ_SERVIDOR")
        strSQL.Append(" where RH24_ID_FREQ_LOTACAO = " & codigoLotacaoFrequencia)
        strSQL.Append(" and RH14_ID_LOTACAO_SERVIDOR =  " & codigoServidorLotacao)

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

