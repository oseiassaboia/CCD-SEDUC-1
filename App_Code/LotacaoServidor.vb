Imports Microsoft.VisualBasic
Imports System.Data

Public Class LotacaoServidor
	Implements IDisposable

	Private RH14_ID_LOTACAO_SERVIDOR As Integer
	Private RH36_ID_LOTACAO As String
	Private RH02_ID_SERVIDOR As String
	Private RH06_ID_FUNCAO As String
	Private CA04_ID_USUARIO As Integer
	Private CA04_ID_USUARIO_DESLIGAMENTO As Integer
	Private RH14_TP_LOTACAO_SERVIDOR As String
	Private RH14_DT_LOTACAO_SERVIDOR As String
	Private RH14_DT_DESLIGAMENTO As String
	Private RH14_DH_CADASTRO As String
	Private RH88_ID_PERIODO As Integer

	Public Property LotacaoServidorId() As Integer
		Get
			Return RH14_ID_LOTACAO_SERVIDOR
		End Get
		Set(ByVal Value As Integer)
			RH14_ID_LOTACAO_SERVIDOR = Value
		End Set
	End Property
	Public Property LotacaoId() As String
		Get
			Return RH36_ID_LOTACAO
		End Get
		Set(ByVal Value As String)
			RH36_ID_LOTACAO = Value
		End Set
	End Property
	Public Property ServidorId() As String
		Get
			Return RH02_ID_SERVIDOR
		End Get
		Set(ByVal Value As String)
			RH02_ID_SERVIDOR = Value
		End Set
	End Property
	Public Property FuncaoId() As String
		Get
			Return RH06_ID_FUNCAO
		End Get
		Set(ByVal Value As String)
			RH06_ID_FUNCAO = Value
		End Set
	End Property
	Public Property UsuarioId() As Integer
		Get
			Return CA04_ID_USUARIO
		End Get
		Set(ByVal Value As Integer)
			CA04_ID_USUARIO = Value
		End Set
	End Property
	Public Property UsuarioIdDesligamento() As Integer
		Get
			Return CA04_ID_USUARIO_DESLIGAMENTO
		End Get
		Set(ByVal Value As Integer)
			CA04_ID_USUARIO_DESLIGAMENTO = Value
		End Set
	End Property
	Public Property TipoLotacao() As String
		Get
			Return RH14_TP_LOTACAO_SERVIDOR
		End Get
		Set(ByVal Value As String)
			RH14_TP_LOTACAO_SERVIDOR = Value
		End Set
	End Property
	Public Property DataLotacao() As String
		Get
			Return RH14_DT_LOTACAO_SERVIDOR
		End Get
		Set(ByVal Value As String)
			RH14_DT_LOTACAO_SERVIDOR = Value
		End Set
	End Property
	Public Property DataDesligamento() As String
		Get
			Return RH14_DT_DESLIGAMENTO
		End Get
		Set(ByVal Value As String)
			RH14_DT_DESLIGAMENTO = Value
		End Set
	End Property
	Public Property DataHoraCadastro() As String
		Get
			Return RH14_DH_CADASTRO
		End Get
		Set(ByVal Value As String)
			RH14_DH_CADASTRO = Value
		End Set
	End Property
	Public Property PeriodoId() As Integer
		Get
			Return RH88_ID_PERIODO
		End Get
		Set(value As Integer)
			RH88_ID_PERIODO = value
		End Set
	End Property

	Public Sub New(Optional ByVal LotacaoServidorId As Integer = 0)
		If LotacaoServidorId > 0 Then
			Obter(LotacaoServidorId)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder

		strSQL.Append(" select * ")
		strSQL.Append(" from RH14_LOTACAO_SERVIDOR")
		strSQL.Append(" where RH14_ID_LOTACAO_SERVIDOR = " & LotacaoServidorId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH36_ID_LOTACAO") = ProBanco(RH36_ID_LOTACAO, eTipoValor.CHAVE)
		dr("RH02_ID_SERVIDOR") = ProBanco(RH02_ID_SERVIDOR, eTipoValor.CHAVE)
		dr("RH06_ID_FUNCAO") = ProBanco(RH06_ID_FUNCAO, eTipoValor.CHAVE)
		dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
		dr("CA04_ID_USUARIO_DESLIGAMENTO") = ProBanco(CA04_ID_USUARIO_DESLIGAMENTO, eTipoValor.CHAVE)
		dr("RH14_TP_LOTACAO_SERVIDOR") = ProBanco(RH14_TP_LOTACAO_SERVIDOR, eTipoValor.TEXTO)
		dr("RH14_DT_LOTACAO_SERVIDOR") = ProBanco(RH14_DT_LOTACAO_SERVIDOR, eTipoValor.DATA)
		dr("RH14_DT_DESLIGAMENTO") = ProBanco(RH14_DT_DESLIGAMENTO, eTipoValor.DATA)
		dr("RH14_DH_CADASTRO") = ProBanco(RH14_DH_CADASTRO, eTipoValor.DATA_COMPLETA)
		dr("RH88_ID_PERIODO") = ProBanco(RH88_ID_PERIODO, eTipoValor.CHAVE)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal LotacaoServidorId As String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder

		strSQL.Append(" select * ")
		strSQL.Append(" from RH14_LOTACAO_SERVIDOR")
		strSQL.Append(" where RH14_ID_LOTACAO_SERVIDOR = " & LotacaoServidorId)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)

			RH14_ID_LOTACAO_SERVIDOR = DoBanco(dr("RH14_ID_LOTACAO_SERVIDOR"), eTipoValor.CHAVE)
			RH36_ID_LOTACAO = DoBanco(dr("RH36_ID_LOTACAO"), eTipoValor.CHAVE)
			RH02_ID_SERVIDOR = DoBanco(dr("RH02_ID_SERVIDOR"), eTipoValor.CHAVE)
			RH06_ID_FUNCAO = DoBanco(dr("RH06_ID_FUNCAO"), eTipoValor.CHAVE)
			CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
			CA04_ID_USUARIO_DESLIGAMENTO = DoBanco(dr("CA04_ID_USUARIO_DESLIGAMENTO"), eTipoValor.CHAVE)
			RH14_TP_LOTACAO_SERVIDOR = DoBanco(dr("RH14_TP_LOTACAO_SERVIDOR"), eTipoValor.TEXTO)
			RH14_DT_LOTACAO_SERVIDOR = DoBanco(dr("RH14_DT_LOTACAO_SERVIDOR"), eTipoValor.TEXTO)
			RH14_DT_DESLIGAMENTO = DoBanco(dr("RH14_DT_DESLIGAMENTO"), eTipoValor.TEXTO)
			RH14_DH_CADASTRO = DoBanco(dr("RH14_DH_CADASTRO"), eTipoValor.DATA)
			RH88_ID_PERIODO = DoBanco(dr("RH88_ID_PERIODO"), eTipoValor.CHAVE)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub
	Public Function PesquisarServidores(Optional ByVal Sort As String = "", Optional ByVal Regional As Integer = 0, Optional ByVal Cidade As Integer = 0, Optional ByVal Lotacao As Integer = 0, Optional ByVal Enturmados As Boolean = True) As DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append(" select *, RH40_NM_SUBLOTACAO + ' - ' + RH52_NM_FUNCAO_MAPEAMENTO as FuncaoLotado  ")
		strSQL.Append(" from RH14_LOTACAO_SERVIDOR RH14   ")
		strSQL.Append(" inner join RH02_SERVIDOR RH02 on RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR   ")
		strSQL.Append(" inner join RH01_PESSOA RH01 ON RH01.RH01_ID_PESSOA = RH02.RH01_ID_PESSOA ")
		strSQL.Append(" inner join RH05_TIPO_VINCULO RH05 ON RH02.RH05_ID_TIPO_VINCULO = RH05.RH05_ID_TIPO_VINCULO ")
		strSQL.Append(" left join RH06_FUNCAO RH06 on RH06.RH06_ID_FUNCAO = RH14.RH06_ID_FUNCAO   ")
		strSQL.Append(" inner join RH16_CARGO RH16 ON RH16.RH16_ID_CARGO = RH02.RH16_ID_CARGO   ")
		strSQL.Append(" inner join RH36_LOTACAO RH36 on RH36.RH36_ID_LOTACAO = RH14.RH36_ID_LOTACAO   ")
		strSQL.Append(" inner join dbgeral..TG03_MUNICIPIO TG03 ON RH36.TG03_ID_MUNICIPIO = TG03.TG03_ID_MUNICIPIO   ")
		strSQL.Append(" inner join dbgeral..TG05_REGIONAL TG05 ON TG05.TG05_ID_REGIONAL = TG03.TG05_ID_REGIONAL   ")
		strSQL.Append(" left join RH53_FUNCAO_LOTADO RH53 ON RH53.RH14_ID_LOTACAO_SERVIDOR = RH14.RH14_ID_LOTACAO_SERVIDOR ")
		strSQL.Append(" left join RH52_FUNCAO_MAPEAMENTO RH52 ON RH52.RH52_ID_FUNCAO_MAPEAMENTO = RH53.RH52_ID_FUNCAO_MAPEAMENTO ")
		strSQL.Append(" LEFT JOIN RH40_SUBLOTACAO RH40 ON RH40.RH40_ID_SUBLOTACAO = RH52.RH40_ID_SUBLOTACAO ")
		strSQL.Append(" where RH14.RH14_ID_LOTACAO_SERVIDOR is not null   ")
		strSQL.Append(" AND RH02.RH07_ID_SITUACAO_SERVIDOR IN (1,10,11) and Rh14.RH14_DT_DESLIGAMENTO is null ")
		If Enturmados Then
			strSQL.Append(" AND  not exists   ")
			strSQL.Append(" 	(SELECT RH80.RH14_ID_LOTACAO_SERVIDOR  FROM RH80_ALOCACAO_CARGA_HORARIA RH80  ")
			strSQL.Append(" 		LEFT JOIN DBDIARIO..DE13_HORARIO_TURMA  DE13 ON DE13.RH80_ID_ALOCACAO_CARGA_HORARIA = RH80.RH80_ID_ALOCACAO_CARGA_HORARIA ")
			strSQL.Append(" 		LEFT JOIN RH74_HABILITACAO RH74 ON RH80.RH80_ID_ALOCACAO_CARGA_HORARIA = RH74.RH80_ID_ALOCACAO_CARGA_HORARIA ")
			strSQL.Append(" 		WHERE RH80.RH14_ID_LOTACAO_SERVIDOR = RH14.RH14_ID_LOTACAO_SERVIDOR AND DE13.DE13_DH_EXCLUSAO IS NULL AND RH74.RH74_DH_DESATIVACAO IS NULL ")
			strSQL.Append("             and RH80.RH80_DH_DESATIVACAO is null ) ")
		End If


		If Regional > 0 Then
			strSQL.Append(" and TG05.TG05_ID_REGIONAL = " & Regional)
		End If

		If Cidade > 0 Then
			strSQL.Append(" And TG03.TG03_ID_CIDADE = " & Cidade)
		End If

		If Lotacao > 0 Then
			strSQL.Append(" And RH36.RH36_ID_LOTACAO = " & Lotacao)
		End If

		strSQL.Append(" Order By " & IIf(Sort = "", "RH01_NM_PESSOA", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function
	Public Function PesquisarServidoresAtivos(Optional ByVal Sort As String = "", Optional ByVal Regional As Integer = 0, Optional ByVal Cidade As Integer = 0 _
											  , Optional ByVal Lotacao As Integer = 0, Optional NomeServidor As String = "", Optional ByVal Periodo As String = "" _
											  , Optional LotacaoServidor As Integer = 0, Optional ByVal Servidor As Integer = 0) As DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append(" select *   ")
		strSQL.Append(" from RH14_LOTACAO_SERVIDOR RH14   ")
		strSQL.Append(" inner join RH02_SERVIDOR RH02 on RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR   ")
		strSQL.Append(" inner join RH01_PESSOA RH01 ON RH01.RH01_ID_PESSOA = RH02.RH01_ID_PESSOA ")
		strSQL.Append(" inner join RH05_TIPO_VINCULO RH05 ON RH02.RH05_ID_TIPO_VINCULO = RH05.RH05_ID_TIPO_VINCULO ")
		strSQL.Append(" left join RH06_FUNCAO RH06 on RH06.RH06_ID_FUNCAO = RH14.RH06_ID_FUNCAO   ")
		strSQL.Append(" inner join RH16_CARGO RH16 ON RH16.RH16_ID_CARGO = RH02.RH16_ID_CARGO   ")
		strSQL.Append(" inner join RH36_LOTACAO RH36 on RH36.RH36_ID_LOTACAO = RH14.RH36_ID_LOTACAO   ")
		strSQL.Append(" inner join dbgeral..TG03_MUNICIPIO TG03 ON RH36.TG03_ID_MUNICIPIO = TG03.TG03_ID_MUNICIPIO   ")
		strSQL.Append(" inner join dbgeral..TG05_REGIONAL TG05 ON TG05.TG05_ID_REGIONAL = TG03.TG05_ID_REGIONAL   ")
		strSQL.Append(" left join	RH88_PERIODO			rh88	on	rh14.RH88_ID_PERIODO = rh88.RH88_ID_PERIODO ")
		strSQL.Append(" where RH14.RH14_ID_LOTACAO_SERVIDOR is not null   ")
		strSQL.Append(" AND RH02.RH07_ID_SITUACAO_SERVIDOR IN (1,10,11) and Rh14.RH14_DT_DESLIGAMENTO is null ")

		If Regional > 0 Then
			strSQL.Append(" and TG05.TG05_ID_REGIONAL = " & Regional)
		End If

		If Servidor > 0 Then
			strSQL.Append(" and rh14.rh02_id_servidor = " & LotacaoServidor)
		End If

		If LotacaoServidor > 0 Then
			strSQL.Append(" and rh14.RH14_ID_LOTACAO_SERVIDOR = " & LotacaoServidor)
		End If

		If Cidade > 0 Then
			strSQL.Append(" And TG03.TG03_ID_CIDADE = " & Cidade)
		End If

		If Lotacao > 0 Then
			strSQL.Append(" And RH36.RH36_ID_LOTACAO = " & Lotacao)
		End If

		If NomeServidor <> "" Then
			strSQL.Append(" And  RH01_NM_PESSOA Like '%" & NomeServidor.ToUpper & "%' collate sql_latin1_general_cp1251_cs_as")
		End If

		If Periodo <> "" Then
			strSQL.Append(" and rh88.RH88_NM_PERIODO = " & Periodo)
		End If


		strSQL.Append(" Order By " & IIf(Sort = "", "RH01_NM_PESSOA", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function
	Public Function PesquisarServidoresAtviso(Optional ByVal Sort As String = "", Optional ByVal Regional As Integer = 0, Optional ByVal Cidade As Integer = 0, Optional ByVal Lotacao As Integer = 0, Optional NomeServidor As String = "", Optional ByVal Periodo As String = "") As DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append(" select *, RH40_NM_SUBLOTACAO + ' - ' + RH52_NM_FUNCAO_MAPEAMENTO as FuncaoLotado  ")
		strSQL.Append(" from RH14_LOTACAO_SERVIDOR RH14   ")
		strSQL.Append(" inner join RH02_SERVIDOR RH02 on RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR   ")
		strSQL.Append(" inner join RH01_PESSOA RH01 ON RH01.RH01_ID_PESSOA = RH02.RH01_ID_PESSOA ")
		strSQL.Append(" inner join RH05_TIPO_VINCULO RH05 ON RH02.RH05_ID_TIPO_VINCULO = RH05.RH05_ID_TIPO_VINCULO ")
		strSQL.Append(" left join RH06_FUNCAO RH06 on RH06.RH06_ID_FUNCAO = RH14.RH06_ID_FUNCAO   ")
		strSQL.Append(" inner join RH16_CARGO RH16 ON RH16.RH16_ID_CARGO = RH02.RH16_ID_CARGO   ")
		strSQL.Append(" inner join RH36_LOTACAO RH36 on RH36.RH36_ID_LOTACAO = RH14.RH36_ID_LOTACAO   ")
		strSQL.Append(" inner join dbgeral..TG03_MUNICIPIO TG03 ON RH36.TG03_ID_MUNICIPIO = TG03.TG03_ID_MUNICIPIO   ")
		strSQL.Append(" inner join dbgeral..TG05_REGIONAL TG05 ON TG05.TG05_ID_REGIONAL = TG03.TG05_ID_REGIONAL   ")
		strSQL.Append(" left join RH53_FUNCAO_LOTADO RH53 ON RH53.RH14_ID_LOTACAO_SERVIDOR = RH14.RH14_ID_LOTACAO_SERVIDOR ")
		strSQL.Append(" left join RH52_FUNCAO_MAPEAMENTO RH52 ON RH52.RH52_ID_FUNCAO_MAPEAMENTO = RH53.RH52_ID_FUNCAO_MAPEAMENTO ")
		strSQL.Append(" left join	RH88_PERIODO			rh88	on	rh14.RH88_ID_PERIODO = rh88.RH88_ID_PERIODO ")
		strSQL.Append(" LEFT JOIN RH40_SUBLOTACAO RH40 ON RH40.RH40_ID_SUBLOTACAO = RH52.RH40_ID_SUBLOTACAO ")
		strSQL.Append(" where RH14.RH14_ID_LOTACAO_SERVIDOR is not null   ")
		strSQL.Append(" AND RH02.RH07_ID_SITUACAO_SERVIDOR IN (1,10,11) and Rh14.RH14_DT_DESLIGAMENTO is null ")

		If Regional > 0 Then
			strSQL.Append(" and TG05.TG05_ID_REGIONAL = " & Regional)
		End If

		If Cidade > 0 Then
			strSQL.Append(" And TG03.TG03_ID_CIDADE = " & Cidade)
		End If

		If Lotacao > 0 Then
			strSQL.Append(" And RH36.RH36_ID_LOTACAO = " & Lotacao)
		End If

		If NomeServidor <> "" Then
			strSQL.Append(" And  RH01_NM_PESSOA Like '%" & NomeServidor.ToUpper & "%' collate sql_latin1_general_cp1251_cs_as")
		End If

		If Periodo <> "" Then
			strSQL.Append(" and rh88.RH88_NM_PERIODO = " & Periodo)
		End If


		strSQL.Append(" Order By " & IIf(Sort = "", "RH01_NM_PESSOA", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function
	Public Function PesquisarMapeamento(Optional ByVal Sort As String = "", Optional ByVal LotacaoServidor As Integer = 0, Optional ByVal Periodo As String = "") As Data.DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append(" select rh06.RH06_NM_FUNCAO, RH52.RH52_NM_FUNCAO_MAPEAMENTO, rh88.RH88_NM_PERIODO, rh14.RH14_ID_LOTACAO_SERVIDOR, isnull(rh53.RH53_ID_FUNCAO_LOTADO,0) as RH53_ID_FUNCAO_LOTADO, TG06_NM_TURNO ")
		strSQL.Append(" from RH14_LOTACAO_SERVIDOR RH14 ")
		strSQL.Append(" left join	RH06_FUNCAO				RH06	on	rh14.RH06_ID_FUNCAO				=	rh06.RH06_ID_FUNCAO ")
		strSQL.Append(" left join	RH53_FUNCAO_LOTADO		RH53	on	RH14.RH14_ID_LOTACAO_SERVIDOR	=	rh53.RH14_ID_LOTACAO_SERVIDOR ")
		strSQL.Append(" LEFT JOIN	RH52_FUNCAO_MAPEAMENTO	RH52	ON	RH53.RH52_ID_FUNCAO_MAPEAMENTO	=	RH52.RH52_ID_FUNCAO_MAPEAMENTO ")
		strSQL.Append(" LEFT JOIN	DBGERAL..TG06_turno	TG06	ON	TG06.TG06_ID_TURNO	=	rh53.TG06_ID_TURNO ")
		strSQL.Append(" left join	RH88_PERIODO			rh88	on	rh14.RH88_ID_PERIODO = rh88.RH88_ID_PERIODO ")
		strSQL.Append(" where rh14.RH14_DT_DESLIGAMENTO is null  ")

		If LotacaoServidor > 0 Then
			strSQL.Append(" and rh14.RH14_ID_LOTACAO_SERVIDOR = " & LotacaoServidor)
		End If

		If Periodo <> "" Then
			strSQL.Append(" and rh88.RH88_NM_PERIODO = " & Periodo)
		End If

		strSQL.Append(" Order By " & IIf(Sort = "", "RH14.RH02_ID_SERVIDOR", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function Pesquisar(Optional ByVal Sort As String = "", Optional LotacaoServidorId As Integer = 0, Optional LotacaoId As Integer = 0 _
							  , Optional ServidorId As Integer = 0, Optional FuncaoId As Integer = 0, Optional UsuarioId As Integer = 0 _
							  , Optional UsuarioIdDesligamento As Integer = 0, Optional TipoLotacao As String = "", Optional DataLotacao As String = "" _
							  , Optional DataDesligamento As String = "", Optional DataHoraCadastro As String = "", Optional PessoaId As Integer = 0 _
							  , Optional ByVal LotacaoServidorAtiva As Boolean = False, Optional FrequenciaLotacao As Integer = 0 _
							  , Optional Ferias As Boolean = False, Optional ByVal Periodo As String = "", Optional ByVal IDPeriodo As Integer = 0) As DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append(" Select *,  'LOTAÇÃO: '+RH36_NM_LOTACAO+ ' - Tipo de lotação: '+case when RH14_TP_LOTACAO_SERVIDOR ='P' then 'PRINCIPAL' when RH14_TP_LOTACAO_SERVIDOR ='C' then 'COMPLEMENTAR' end as DESCRICAO, case  isnull(RH14_DT_DESLIGAMENTO,'')  when null then 0 when '' then 0 else  1  end  as Desativado ")

		If FrequenciaLotacao > 0 Then

			strSQL.Append(" ,(select distinct top 1 case  RH24_ST_FREQ_LOTACAO  when  1  then 'LANCADO'  ELSE 'NAO LANCADO' end as FREQUENCIA ")
			strSQL.Append(" from RH18_FREQ_SERVIDOR rh18 where ServidorLotacao.RH14_ID_LOTACAO_SERVIDOR = rh18.RH14_ID_LOTACAO_SERVIDOR AND rh18.RH24_ID_FREQ_LOTACAO =" & FrequenciaLotacao & ") AS FREQUENCIA, ")
			strSQL.Append(" 		(select  count(isnull(rh18.RH23_ID_TIPO_REGISTRO,0)) ")
			strSQL.Append(" 			from RH18_FREQ_SERVIDOR rh18  ")
			strSQL.Append(" 				where ServidorLotacao.RH14_ID_LOTACAO_SERVIDOR = rh18.RH14_ID_LOTACAO_SERVIDOR  ")
			strSQL.Append(" 				AND rh18.RH24_ID_FREQ_LOTACAO = " & FrequenciaLotacao & "  and rh18.RH23_ID_TIPO_REGISTRO = 1) AS Presenca,   ")
			strSQL.Append(" 		(select  count(isnull(rh18.RH23_ID_TIPO_REGISTRO,0)) ")
			strSQL.Append(" 			from RH18_FREQ_SERVIDOR rh18  ")
			strSQL.Append(" 				where ServidorLotacao.RH14_ID_LOTACAO_SERVIDOR = rh18.RH14_ID_LOTACAO_SERVIDOR  ")
			strSQL.Append(" 				AND rh18.RH24_ID_FREQ_LOTACAO =" & FrequenciaLotacao & " and rh18.RH23_ID_TIPO_REGISTRO = 2) AS Falta, ")
			strSQL.Append(" 		(select  count(isnull(rh18.RH23_ID_TIPO_REGISTRO,0)) ")
			strSQL.Append(" 			from RH18_FREQ_SERVIDOR rh18  ")
			strSQL.Append(" 				where ServidorLotacao.RH14_ID_LOTACAO_SERVIDOR = rh18.RH14_ID_LOTACAO_SERVIDOR  ")
			strSQL.Append(" 				AND rh18.RH24_ID_FREQ_LOTACAO =" & FrequenciaLotacao & " and rh18.RH23_ID_TIPO_REGISTRO = 3) AS Faltas_Justificada,     ")
			strSQL.Append(" 		(select  count(isnull(rh18.RH23_ID_TIPO_REGISTRO,0)) ")
			strSQL.Append(" 			from RH18_FREQ_SERVIDOR rh18  ")
			strSQL.Append(" 				where ServidorLotacao.RH14_ID_LOTACAO_SERVIDOR = rh18.RH14_ID_LOTACAO_SERVIDOR  ")
			strSQL.Append(" 				AND rh18.RH24_ID_FREQ_LOTACAO =" & FrequenciaLotacao & " and rh18.RH23_ID_TIPO_REGISTRO = 4) AS Ferias ,    ")
			strSQL.Append(" 		FREQ_LOTACAO = " & FrequenciaLotacao)
		End If



		If Ferias Then

			strSQL.Append(" ,(select convert(varchar,RH28_DT_INICIO_GOZO,103) +' - '+  convert(varchar,RH28_DT_TERMINO_GOZO,103)     ")
			strSQL.Append(" 	from RH28_FERIAS rh28	 ")
			strSQL.Append(" 	left join RH87_PERIODO_FERIAS rh87 on rh28.RH87_ID_PERIODO_FERIAS = rh87.RH87_ID_PERIODO_FERIAS ")
			strSQL.Append(" 	where RH87_NR_ANO_REFERENCIA = year(getdate()) and Servidor.RH02_ID_SERVIDOR = RH28.RH02_ID_SERVIDOR   ) as FERIAS ")

			strSQL.Append(" ,(select RH28_ID_FERIAS    ")
			strSQL.Append(" 	from RH28_FERIAS rh28	 ")
			strSQL.Append(" 	left join RH87_PERIODO_FERIAS rh87 on rh28.RH87_ID_PERIODO_FERIAS = rh87.RH87_ID_PERIODO_FERIAS ")
			strSQL.Append(" 	where RH87_NR_ANO_REFERENCIA = year(getdate()) and Servidor.RH02_ID_SERVIDOR = RH28.RH02_ID_SERVIDOR   ) as RH28_ID_FERIAS ")

		End If

		strSQL.Append(" from RH14_LOTACAO_SERVIDOR ServidorLotacao")
        strSQL.Append(" left join RH02_SERVIDOR Servidor On Servidor.RH02_ID_SERVIDOR = ServidorLotacao.RH02_ID_SERVIDOR ")


		strSQL.Append(" left join RH06_FUNCAO Funcao On Funcao.RH06_ID_FUNCAO = ServidorLotacao.RH06_ID_FUNCAO ")
		strSQL.Append(" left join RH36_LOTACAO Lotacao On Lotacao.RH36_ID_LOTACAO = ServidorLotacao.RH36_ID_LOTACAO ")
		strSQL.Append(" left join RH01_PESSOA RH01 On RH01.RH01_ID_PESSOA = Servidor.RH01_ID_PESSOA ")
		strSQL.Append(" left join	RH88_PERIODO			rh88	on	ServidorLotacao.RH88_ID_PERIODO = rh88.RH88_ID_PERIODO ")
		strSQL.Append(" where RH14_ID_LOTACAO_SERVIDOR Is Not null ")

		If PessoaId > 0 Then
			strSQL.Append(" And Servidor.RH01_ID_PESSOA = " & PessoaId)
		End If


		If IDPeriodo > 0 Then
			strSQL.Append(" And rh88.RH88_ID_PERIODO = " & IDPeriodo)
		End If


		If LotacaoServidorId > 0 Then
			strSQL.Append(" And RH14_ID_LOTACAO_SERVIDOR = " & LotacaoServidorId)
		End If

		If LotacaoId > 0 Then
			strSQL.Append(" And Lotacao.RH36_ID_LOTACAO = " & LotacaoId)
		End If

		If ServidorId > 0 Then
			strSQL.Append(" And Servidor.RH02_ID_SERVIDOR = " & ServidorId)
		End If

		If Periodo <> "" Then
			strSQL.Append(" and rh88.RH88_NM_PERIODO = " & Periodo)
		End If

		If FuncaoId > 0 Then
			strSQL.Append(" And Funcao.RH06_ID_FUNCAO = " & FuncaoId)
		End If

		If UsuarioId > 0 Then
			strSQL.Append(" And CA04_ID_USUARIO = " & UsuarioId)
		End If

		If UsuarioIdDesligamento > 0 Then
			strSQL.Append(" And CA04_ID_USUARIO_DESLIGAMENTO = " & UsuarioIdDesligamento)
		End If

		If TipoLotacao <> "" Then
			strSQL.Append(" And upper(RH14_TP_LOTACAO_SERVIDOR) Like '%" & TipoLotacao.ToUpper & "%'")
		End If

		If DataLotacao <> "" Then
			strSQL.Append(" and upper(RH14_DT_LOTACAO_SERVIDOR) like '%" & DataLotacao.ToUpper & "%'")
		End If

		If DataDesligamento <> "" Then
			strSQL.Append(" and upper(RH14_DT_DESLIGAMENTO) like '%" & DataDesligamento.ToUpper & "%'")
		End If

		If IsDate(DataHoraCadastro) Then
			strSQL.Append(" and RH14_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
		End If

		If LotacaoServidorAtiva Then
			strSQL.Append(" AND isnull(ServidorLotacao.RH14_DT_DESLIGAMENTO,'') ='' ")
		End If

		strSQL.Append(" Order By " & IIf(Sort = "", "rh01_nm_pessoa", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)

	End Function

	Public Function ObterTabela() As DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder

		strSQL.Append(" select RH14_ID_LOTACAO_SERVIDOR as CODIGO, RH36_ID_LOTACAO as DESCRICAO")
		strSQL.Append(" from RH14_LOTACAO_SERVIDOR")
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

		strSQL.Append(" select max(RH14_ID_LOTACAO_SERVIDOR) from RH14_LOTACAO_SERVIDOR")

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
	Public Function Excluir(ByVal LotacaoServidorId As String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer

		strSQL.Append(" delete ")
		strSQL.Append(" from RH14_LOTACAO_SERVIDOR")
		strSQL.Append(" where RH14_ID_LOTACAO_SERVIDOR = " & LotacaoServidorId)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

#Region "IDisposable Support"
	Private disposedValue As Boolean ' Para detectar chamadas redundantes

	' IDisposable
	Protected Overridable Sub Dispose(disposing As Boolean)
		If Not disposedValue Then
			If disposing Then
				' TODO: descartar estado gerenciado (objetos gerenciados).
			End If

			' TODO: liberar recursos não gerenciados (objetos não gerenciados) e substituir um Finalize() abaixo.
			' TODO: definir campos grandes como nulos.
		End If
		disposedValue = True
	End Sub

	' TODO: substituir Finalize() somente se Dispose(disposing As Boolean) acima tiver o código para liberar recursos não gerenciados.
	'Protected Overrides Sub Finalize()
	'    ' Não altere este código. Coloque o código de limpeza em Dispose(disposing As Boolean) acima.
	'    Dispose(False)
	'    MyBase.Finalize()
	'End Sub

	' Código adicionado pelo Visual Basic para implementar corretamente o padrão descartável.
	Public Sub Dispose() Implements IDisposable.Dispose
		' Não altere este código. Coloque o código de limpeza em Dispose(disposing As Boolean) acima.
		Dispose(True)
		' TODO: remover marca de comentário da linha a seguir se Finalize() for substituído acima.
		GC.SuppressFinalize(Me)
	End Sub
#End Region

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

