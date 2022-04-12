Imports Microsoft.VisualBasic
Imports System.Data

Public Class FrequenciaLotacao

    Implements IDisposable

	Private RH24_ID_FREQ_LOTACAO as Integer
	Private RH17_ID_PERIODO_FREQ as Integer
	Private RH36_ID_LOTACAO as Integer
	Private RH44_ID_LANCAMENTO_FREQ as Integer
	Private CA04_ID_USUARIO as Integer
	Private CA04_ID_USUARIO_ALT as Integer
	Private RH24_DH_CADASTRO as String
	Private RH24_ST_FREQ_LOTACAO as String
	Private RH24_DH_ST_FREQ_LOTACAO as String

	Public Property Codigo() as Integer
		Get
			Return RH24_ID_FREQ_LOTACAO
		End Get
		Set(ByVal Value As Integer)
			RH24_ID_FREQ_LOTACAO = Value
		End Set
	End Property
	Public Property IdPeriodoFrequencia() as Integer
		Get
			Return RH17_ID_PERIODO_FREQ
		End Get
		Set(ByVal Value As Integer)
			RH17_ID_PERIODO_FREQ = Value
		End Set
	End Property
	Public Property idLotacao() as Integer
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
	Public Property idUsuario() as Integer
		Get
			Return CA04_ID_USUARIO
		End Get
		Set(ByVal Value As Integer)
			CA04_ID_USUARIO = Value
		End Set
	End Property
	Public Property idUsuarioAltrecacao() as Integer
		Get
			Return CA04_ID_USUARIO_ALT
		End Get
		Set(ByVal Value As Integer)
			CA04_ID_USUARIO_ALT = Value
		End Set
	End Property
	Public Property DataHoraCadastro() as String
		Get
			Return RH24_DH_CADASTRO
		End Get
		Set(ByVal Value As String)
			RH24_DH_CADASTRO = Value
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
		strSQL.Append(" from RH24_FREQ_LOTACAO")
		strSQL.Append(" where RH24_ID_FREQ_LOTACAO = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH17_ID_PERIODO_FREQ") = ProBanco(RH17_ID_PERIODO_FREQ, eTipoValor.CHAVE)
		dr("RH36_ID_LOTACAO") = ProBanco(RH36_ID_LOTACAO, eTipoValor.CHAVE)
		dr("RH44_ID_LANCAMENTO_FREQ") = ProBanco(RH44_ID_LANCAMENTO_FREQ, eTipoValor.CHAVE)
		dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
		dr("CA04_ID_USUARIO_ALT") = ProBanco(CA04_ID_USUARIO_ALT, eTipoValor.CHAVE)
		dr("RH24_DH_CADASTRO") = ProBanco(RH24_DH_CADASTRO, eTipoValor.DATA)
		dr("RH24_ST_FREQ_LOTACAO") = ProBanco(RH24_ST_FREQ_LOTACAO, eTipoValor.NUMERO_DECIMAL)
		dr("RH24_DH_ST_FREQ_LOTACAO") = ProBanco(RH24_DH_ST_FREQ_LOTACAO, eTipoValor.DATA)

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
		strSQL.Append(" from RH24_FREQ_LOTACAO")
		strSQL.Append(" where RH24_ID_FREQ_LOTACAO = " & Codigo)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH24_ID_FREQ_LOTACAO = DoBanco(dr("RH24_ID_FREQ_LOTACAO"), eTipoValor.CHAVE)
			RH17_ID_PERIODO_FREQ = DoBanco(dr("RH17_ID_PERIODO_FREQ"), eTipoValor.CHAVE)
			RH36_ID_LOTACAO = DoBanco(dr("RH36_ID_LOTACAO"), eTipoValor.CHAVE)
			RH44_ID_LANCAMENTO_FREQ = DoBanco(dr("RH44_ID_LANCAMENTO_FREQ"), eTipoValor.CHAVE)
			CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
			CA04_ID_USUARIO_ALT = DoBanco(dr("CA04_ID_USUARIO_ALT"), eTipoValor.CHAVE)
			RH24_DH_CADASTRO = DoBanco(dr("RH24_DH_CADASTRO"), eTipoValor.DATA)
			RH24_ST_FREQ_LOTACAO = DoBanco(dr("RH24_ST_FREQ_LOTACAO"), eTipoValor.NUMERO_DECIMAL)
			RH24_DH_ST_FREQ_LOTACAO = DoBanco(dr("RH24_DH_ST_FREQ_LOTACAO"), eTipoValor.DATA)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub
    Public Function PesquisarFrequenciaAberta(Optional ByVal Sort As String = "", Optional CodigoLotacao As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select *,RH24.RH24_ID_FREQ_LOTACAO as CODIGO,   ")
        strSQL.Append(" convert(varchar,RH17.RH17_NR_MES) +'/'+convert(varchar,RH17_NR_ANO) + ' - ' +   ")
        strSQL.Append(" case RH24_ST_FREQ_LOTACAO   ")
        strSQL.Append(" when 1 then 'CRIADO'  ")
        strSQL.Append(" WHEN 2 THEN 'ABERTO'  ")
        strSQL.Append(" WHEN 3 THEN 'ENVIADO'   ")
        strSQL.Append(" WHEN 4 THEN 'FECHADO'  ")
        strSQL.Append(" WHEN 5 THEN 'ABERTO PARA RETIFICAÇÃO'  ")
        strSQL.Append(" WHEN 6 THEN 'FECHADO COM RETIFICAÇÃO' END AS DESCRICAO  ")
        strSQL.Append(" from RH24_FREQ_LOTACAO RH24")
        strSQL.Append(" inner join RH17_PERIODO_FREQ RH17 ON  RH17.RH17_ID_PERIODO_FREQ  = RH24.RH17_ID_PERIODO_FREQ ")
        strSQL.Append(" where RH24_ID_FREQ_LOTACAO is not null")
        strSQL.Append(" and RH24_ST_FREQ_LOTACAO in (1,2,5) ")

        If CodigoLotacao > 0 Then
            strSQL.Append(" and RH24.RH36_ID_LOTACAO = " & CodigoLotacao)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH24_ID_FREQ_LOTACAO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function buscarLotacaoAbertaServidor(codigoLotacaoAntigo As Integer) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder


        strSQL.Append(" select RH24_ID_FREQ_LOTACAO, RH17_ID_PERIODO_FREQ, RH36_ID_LOTACAO, RH44_ID_LANCAMENTO_FREQ, RH24_ST_FREQ_LOTACAO from rh24_FREQ_LOTACAO  ")
        strSQL.Append(" where RH24_ST_FREQ_LOTACAO in(1,2,5) and RH36_ID_LOTACAO = " + codigoLotacaoAntigo)

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Sub buscarLotacaoPeriodo(periodoFrequencia As Integer, codigoLotacaoNovo As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * from RH24_FREQ_LOTACAO ")
        strSQL.Append(" where RH36_ID_LOTACAO = " & codigoLotacaoNovo)
        strSQL.Append(" and RH17_ID_PERIODO_FREQ = " & periodoFrequencia)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH24_ID_FREQ_LOTACAO = DoBanco(dr("RH24_ID_FREQ_LOTACAO"), eTipoValor.CHAVE)
            RH17_ID_PERIODO_FREQ = DoBanco(dr("RH17_ID_PERIODO_FREQ"), eTipoValor.CHAVE)
            RH36_ID_LOTACAO = DoBanco(dr("RH36_ID_LOTACAO"), eTipoValor.CHAVE)
            RH44_ID_LANCAMENTO_FREQ = DoBanco(dr("RH44_ID_LANCAMENTO_FREQ"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            CA04_ID_USUARIO_ALT = DoBanco(dr("CA04_ID_USUARIO_ALT"), eTipoValor.CHAVE)
            RH24_DH_CADASTRO = DoBanco(dr("RH24_DH_CADASTRO"), eTipoValor.DATA)
            RH24_ST_FREQ_LOTACAO = DoBanco(dr("RH24_ST_FREQ_LOTACAO"), eTipoValor.NUMERO_DECIMAL)
            RH24_DH_ST_FREQ_LOTACAO = DoBanco(dr("RH24_DH_ST_FREQ_LOTACAO"), eTipoValor.DATA)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

	Public Function atualizarFrequenciaParaFechadoQuandoFecharPeriodo(codigoPeriodoFrequencia As Integer, statusFrequenciaLotacao As Integer) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer

		strSQL.Append(" update RH24_FREQ_LOTACAO set")

		strSQL.Append(" RH24_ST_FREQ_LOTACAO = " & statusFrequenciaLotacao)

		strSQL.Append(" where RH17_ID_PERIODO_FREQ = " & codigoPeriodoFrequencia)
		strSQL.Append(" and RH24_ST_FREQ_LOTACAO in (1,4,5)")

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function
	Public Function PesquisarTodasAsFrequenciasPeriodo(Optional ByVal Sort As String = "", Optional CodigoLotacao As Integer = 0) As DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append(" select *,RH17.RH17_ID_PERIODO_FREQ as CODIGO,   ")
		strSQL.Append(" convert(varchar,RH17.RH17_NR_MES) +'/'+convert(varchar,RH17_NR_ANO) + ' - ' +   ")
		strSQL.Append(" case RH24_ST_FREQ_LOTACAO   ")
		strSQL.Append(" when 1 then 'CRIADO'  ")
		strSQL.Append(" WHEN 2 THEN 'ABERTO'  ")
		strSQL.Append(" WHEN 3 THEN 'ENVIADO'   ")
		strSQL.Append(" WHEN 4 THEN 'FECHADO'  ")
		strSQL.Append(" WHEN 5 THEN 'ABERTO PARA RETIFICAÇÃO'  ")
		strSQL.Append(" WHEN 6 THEN 'FECHADO COM RETIFICAÇÃO' END AS DESCRICAO  ")
		strSQL.Append(" from RH24_FREQ_LOTACAO RH24")
		strSQL.Append(" inner join RH17_PERIODO_FREQ RH17 ON  RH17.RH17_ID_PERIODO_FREQ  = RH24.RH17_ID_PERIODO_FREQ ")
		strSQL.Append(" where RH24_ID_FREQ_LOTACAO is not null")


		If CodigoLotacao > 0 Then
			strSQL.Append(" and RH24.RH36_ID_LOTACAO = " & CodigoLotacao)
		End If

		strSQL.Append(" Order By " & IIf(Sort = "", "RH24_ID_FREQ_LOTACAO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function PesquisarTodasAsFrequencias(Optional ByVal Sort As String = "", Optional CodigoLotacao As Integer = 0) As DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append(" select *,RH24.RH24_ID_FREQ_LOTACAO as CODIGO,   ")
		strSQL.Append(" convert(varchar,RH17.RH17_NR_MES) +'/'+convert(varchar,RH17_NR_ANO) + ' - ' +   ")
		strSQL.Append(" case RH24_ST_FREQ_LOTACAO   ")
		strSQL.Append(" when 1 then 'CRIADO'  ")
		strSQL.Append(" WHEN 2 THEN 'ABERTO'  ")
		strSQL.Append(" WHEN 3 THEN 'ENVIADO'   ")
		strSQL.Append(" WHEN 4 THEN 'FECHADO'  ")
		strSQL.Append(" WHEN 5 THEN 'ABERTO PARA RETIFICAÇÃO'  ")
		strSQL.Append(" WHEN 6 THEN 'FECHADO COM RETIFICAÇÃO' END AS DESCRICAO  ")
		strSQL.Append(" from RH24_FREQ_LOTACAO RH24")
		strSQL.Append(" inner join RH17_PERIODO_FREQ RH17 ON  RH17.RH17_ID_PERIODO_FREQ  = RH24.RH17_ID_PERIODO_FREQ ")
		strSQL.Append(" where RH24_ID_FREQ_LOTACAO is not null")


		If CodigoLotacao > 0 Then
			strSQL.Append(" and RH24.RH36_ID_LOTACAO = " & CodigoLotacao)
		End If

		strSQL.Append(" Order By " & IIf(Sort = "", "RH24_ID_FREQ_LOTACAO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function
	Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional IdPeriodoFrequencia As Integer = 0 _
                              , Optional idLotacao As Integer = 0, Optional IdLancamento As Integer = 0, Optional idUsuario As Integer = 0 _
                              , Optional idUsuarioAltrecacao As Integer = 0, Optional DataHoraCadastro As String = "", Optional SituacaoFrequencia As String = "" _
                              , Optional DataHoraSituacaoFrequencia As String = "", Optional ByVal Regional As Integer = 0, Optional ByVal Cidade As Integer = 0, Optional ByVal PeriodoAtivo As Boolean = True) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select *,isnull(RH48.RH48_NM_TIPO_ESCOLA,'') + ' ' + RH36.RH36_NM_LOTACAO as LOTACAO, ")
        strSQL.Append("case RH24_ST_FREQ_LOTACAO   ")
        strSQL.Append("when 1 then 'CRIADO'  ")
        strSQL.Append("when 2 then 'ABERTO'  ")
        strSQL.Append("when 3 then 'ENVIADO'  ")
        strSQL.Append("when 4 then 'FECHADO'  ")
        strSQL.Append("when 5 then 'ABERTO PARA RETIFICACAO'  ")
        strSQL.Append("when 6 then 'FECHADO COM RETIFICACAO'  ")
        strSQL.Append(" end as Situacao ")
        strSQL.Append(" from RH24_FREQ_LOTACAO RH24")
        strSQL.Append(" inner join RH36_LOTACAO RH36 on RH24.rh36_ID_LOTACAO = RH36.rh36_ID_LOTACAO")
        strSQL.Append(" Left Join  RH48_TIPO_ESCOLA as RH48 on RH48.RH48_ID_TIPO_ESCOLA = RH36.RH48_ID_TIPO_ESCOLA ")
        strSQL.Append(" inner join DBGERAL..TG03_MUNICIPIO TG03 on TG03.TG03_ID_MUNICIPIO  = RH36.TG03_ID_MUNICIPIO ")
        strSQL.Append(" inner join DBGERAL..TG05_REGIONAL TG05 ON   TG05.TG05_ID_REGIONAL  = TG03.TG05_ID_REGIONAL ")
        strSQL.Append(" inner join RH17_PERIODO_FREQ RH17 ON  RH17.RH17_ID_PERIODO_FREQ  = RH24.RH17_ID_PERIODO_FREQ ")
        strSQL.Append(" where RH24_ID_FREQ_LOTACAO is not null")

        If Codigo > 0 Then
            strSQL.Append(" and RH24_ID_FREQ_LOTACAO = " & Codigo)
        End If

        If PeriodoAtivo Then
            strSQL.Append(" and RH17_ST_PERIODO_FREQ = 'A'")
        End If

        If IdPeriodoFrequencia > 0 Then
            strSQL.Append(" and RH17.RH17_ID_PERIODO_FREQ = " & IdPeriodoFrequencia)
        End If

        If idLotacao > 0 Then
            strSQL.Append(" and RH24.RH36_ID_LOTACAO = " & idLotacao)
        End If

        If IdLancamento > 0 Then
            strSQL.Append(" and RH44_ID_LANCAMENTO_FREQ = " & IdLancamento)
        End If

        If idUsuario > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & idUsuario)
        End If

        If idUsuarioAltrecacao > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO_ALT = " & idUsuarioAltrecacao)
        End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" and RH24_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
        End If

        If Cidade > 0 Then
            strSQL.Append(" and TG03.TG03_ID_MUNICIPIO = " & Cidade)
        End If

        If Regional > 0 Then
            strSQL.Append(" and TG05.TG05_ID_REGIONAL = " & Regional)
        End If


        If IsNumeric(SituacaoFrequencia.Replace(".", "")) And SituacaoFrequencia > 0 Then
            strSQL.Append(" and RH24_ST_FREQ_LOTACAO = " & SituacaoFrequencia.Replace(".", "").Replace(",", "."))
        End If

        If IsDate(DataHoraSituacaoFrequencia) Then
            strSQL.Append(" and RH24_DH_ST_FREQ_LOTACAO = Convert(DateTime, '" & DataHoraSituacaoFrequencia & "', 103)")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH24_ID_FREQ_LOTACAO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH24_ID_FREQ_LOTACAO as CODIGO, RH17_ID_PERIODO_FREQ as DESCRICAO")
		strSQL.Append(" from RH24_FREQ_LOTACAO")
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
		
		strSQL.Append(" select max(RH24_ID_FREQ_LOTACAO) from RH24_FREQ_LOTACAO")

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
		strSQL.Append(" from RH24_FREQ_LOTACAO")
		strSQL.Append(" where RH24_ID_FREQ_LOTACAO = " & Codigo)

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

