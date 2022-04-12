Imports Microsoft.VisualBasic
Imports System.Data

Public Class Pessoa
    Implements IDisposable

    Private RH01_ID_PESSOA as Integer
	Private RH01_NU_CPF as String
	Private RH01_NM_PESSOA as String
	Private RH01_NM_MAE as String
	Private RH01_NM_PAI as String
	Private RH01_TP_SEXO as String
	Private RH01_NU_DDD_TELEFONE as String
	Private RH01_NU_TELEFONE as String
	Private RH01_NU_DDD_CELULAR as String
	Private RH01_NU_CELULAR as String
	Private RH01_NM_EMAIL as String
	Private RH01_DT_NASCIMENTO as String
    Private RH01_IN_ATESTADO_FISICO_MENTAL As String
    Private RH01_IN_DECLARA_ACUMULACAO_CARGO as String
	Private RH01_IN_DECLARACAO_BENS as String
	Private RH01_DH_CADASTRO as String
	Private CA04_ID_USUARIO as Integer
	Private RH01_IN_CONTRIBUI_FUNBEM as String
	Private CA04_ID_USUARIO_LOGIN As Integer

	Public Property PessoaId() as Integer
		Get
			Return RH01_ID_PESSOA
		End Get
		Set(ByVal Value As Integer)
			RH01_ID_PESSOA = Value
		End Set
	End Property
	Public Property Cpf() as String
		Get
			Return RH01_NU_CPF
		End Get
		Set(ByVal Value As String)
			RH01_NU_CPF = Value
		End Set
	End Property
	Public Property Nome() as String
		Get
			Return RH01_NM_PESSOA
		End Get
		Set(ByVal Value As String)
			RH01_NM_PESSOA = Value
		End Set
	End Property
	Public Property Mae() as String
		Get
			Return RH01_NM_MAE
		End Get
		Set(ByVal Value As String)
			RH01_NM_MAE = Value
		End Set
	End Property
	Public Property Pai() as String
		Get
			Return RH01_NM_PAI
		End Get
		Set(ByVal Value As String)
			RH01_NM_PAI = Value
		End Set
	End Property
	Public Property Sexo() as String
		Get
			Return RH01_TP_SEXO
		End Get
		Set(ByVal Value As String)
			RH01_TP_SEXO = Value
		End Set
	End Property
	Public Property DddTelefone() as String
		Get
			Return RH01_NU_DDD_TELEFONE
		End Get
		Set(ByVal Value As String)
			RH01_NU_DDD_TELEFONE = Value
		End Set
	End Property
	Public Property Telefone() as String
		Get
			Return RH01_NU_TELEFONE
		End Get
		Set(ByVal Value As String)
			RH01_NU_TELEFONE = Value
		End Set
	End Property
	Public Property DddCelular() as String
		Get
			Return RH01_NU_DDD_CELULAR
		End Get
		Set(ByVal Value As String)
			RH01_NU_DDD_CELULAR = Value
		End Set
	End Property
	Public Property Celular() as String
		Get
			Return RH01_NU_CELULAR
		End Get
		Set(ByVal Value As String)
			RH01_NU_CELULAR = Value
		End Set
	End Property
	Public Property Email() as String
		Get
			Return RH01_NM_EMAIL
		End Get
		Set(ByVal Value As String)
			RH01_NM_EMAIL = Value
		End Set
	End Property
	Public Property DataNascimento() as String
		Get
			Return RH01_DT_NASCIMENTO
		End Get
		Set(ByVal Value As String)
			RH01_DT_NASCIMENTO = Value
		End Set
	End Property
	Public Property AtestadoFisicoMental() as String
		Get
            Return RH01_IN_ATESTADO_FISICO_MENTAL
        End Get
		Set(ByVal Value As String)
            RH01_IN_ATESTADO_FISICO_MENTAL = Value
        End Set
	End Property
	Public Property DeclaraAcumulacaoCargo() as String
		Get
			Return RH01_IN_DECLARA_ACUMULACAO_CARGO
		End Get
		Set(ByVal Value As String)
			RH01_IN_DECLARA_ACUMULACAO_CARGO = Value
		End Set
	End Property
	Public Property DeclaracaoBens() as String
		Get
			Return RH01_IN_DECLARACAO_BENS
		End Get
		Set(ByVal Value As String)
			RH01_IN_DECLARACAO_BENS = Value
		End Set
	End Property
	Public Property DataHoraCadastro() as String
		Get
			Return RH01_DH_CADASTRO
		End Get
		Set(ByVal Value As String)
			RH01_DH_CADASTRO = Value
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
	Public Property ContribuicaoFunbem() as String
		Get
			Return RH01_IN_CONTRIBUI_FUNBEM
		End Get
		Set(ByVal Value As String)
			RH01_IN_CONTRIBUI_FUNBEM = Value
		End Set
	End Property
	Public Property UsuarioLogin() As Integer
	    Get
	        Return CA04_ID_USUARIO_LOGIN
	    End Get
	    Set(ByVal Value As Integer)
	        CA04_ID_USUARIO_LOGIN = Value
	    End Set
	End Property

	Public Sub New(Optional ByVal PessoaId as Integer = 0)
		If PessoaId >0 Then
			Obter(PessoaId)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH01_PESSOA")
		strSQL.Append(" where RH01_ID_PESSOA = " & PessoaId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH01_NU_CPF") = ProBanco(RH01_NU_CPF, eTipoValor.TEXTO)
		dr("RH01_NM_PESSOA") = ProBanco(RH01_NM_PESSOA, eTipoValor.TEXTO)
		dr("RH01_NM_MAE") = ProBanco(RH01_NM_MAE, eTipoValor.TEXTO)
		dr("RH01_NM_PAI") = ProBanco(RH01_NM_PAI, eTipoValor.TEXTO)
		dr("RH01_TP_SEXO") = ProBanco(RH01_TP_SEXO, eTipoValor.TEXTO)
		dr("RH01_NU_DDD_TELEFONE") = ProBanco(RH01_NU_DDD_TELEFONE, eTipoValor.TEXTO)
		dr("RH01_NU_TELEFONE") = ProBanco(RH01_NU_TELEFONE, eTipoValor.TEXTO)
		dr("RH01_NU_DDD_CELULAR") = ProBanco(RH01_NU_DDD_CELULAR, eTipoValor.TEXTO)
		dr("RH01_NU_CELULAR") = ProBanco(RH01_NU_CELULAR, eTipoValor.TEXTO)
		dr("RH01_NM_EMAIL") = ProBanco(RH01_NM_EMAIL, eTipoValor.TEXTO)
		dr("RH01_DT_NASCIMENTO") = ProBanco(RH01_DT_NASCIMENTO, eTipoValor.DATA)
        dr("RH01_IN_ATESTADO_FISICO_MENTAL") = ProBanco(RH01_IN_ATESTADO_FISICO_MENTAL, eTipoValor.BOOLEANO)
        dr("RH01_IN_DECLARA_ACUMULACAO_CARGO") = ProBanco(RH01_IN_DECLARA_ACUMULACAO_CARGO, eTipoValor.BOOLEANO)
		dr("RH01_IN_DECLARACAO_BENS") = ProBanco(RH01_IN_DECLARACAO_BENS, eTipoValor.BOOLEANO)
		dr("RH01_DH_CADASTRO") = ProBanco(RH01_DH_CADASTRO, eTipoValor.DATA)
		dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
		dr("RH01_IN_CONTRIBUI_FUNBEM") = ProBanco(RH01_IN_CONTRIBUI_FUNBEM, eTipoValor.BOOLEANO)
	    dr("CA04_ID_USUARIO_LOGIN") = ProBanco(CA04_ID_USUARIO_LOGIN, eTipoValor.CHAVE)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

    Public Sub Obter(ByVal PessoaId As String)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH01_PESSOA")
        strSQL.Append(" where RH01_ID_PESSOA = " & PessoaId)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH01_ID_PESSOA = DoBanco(dr("RH01_ID_PESSOA"), eTipoValor.CHAVE)
            RH01_NU_CPF = DoBanco(dr("RH01_NU_CPF"), eTipoValor.TEXTO)
            RH01_NM_PESSOA = DoBanco(dr("RH01_NM_PESSOA"), eTipoValor.TEXTO)
            RH01_NM_MAE = DoBanco(dr("RH01_NM_MAE"), eTipoValor.TEXTO)
            RH01_NM_PAI = DoBanco(dr("RH01_NM_PAI"), eTipoValor.TEXTO)
            RH01_TP_SEXO = DoBanco(dr("RH01_TP_SEXO"), eTipoValor.TEXTO)
            RH01_NU_DDD_TELEFONE = DoBanco(dr("RH01_NU_DDD_TELEFONE"), eTipoValor.TEXTO)
            RH01_NU_TELEFONE = DoBanco(dr("RH01_NU_TELEFONE"), eTipoValor.TEXTO)
            RH01_NU_DDD_CELULAR = DoBanco(dr("RH01_NU_DDD_CELULAR"), eTipoValor.TEXTO)
            RH01_NU_CELULAR = DoBanco(dr("RH01_NU_CELULAR"), eTipoValor.TEXTO)
            RH01_NM_EMAIL = DoBanco(dr("RH01_NM_EMAIL"), eTipoValor.TEXTO)
            RH01_DT_NASCIMENTO = DoBanco(dr("RH01_DT_NASCIMENTO"), eTipoValor.DATA)
            RH01_IN_ATESTADO_FISICO_MENTAL = DoBanco(dr("RH01_IN_ATESTADO_FISICO_MENTAL"), eTipoValor.BOOLEANO)
            RH01_IN_DECLARA_ACUMULACAO_CARGO = DoBanco(dr("RH01_IN_DECLARA_ACUMULACAO_CARGO"), eTipoValor.BOOLEANO)
            RH01_IN_DECLARACAO_BENS = DoBanco(dr("RH01_IN_DECLARACAO_BENS"), eTipoValor.BOOLEANO)
            RH01_DH_CADASTRO = DoBanco(dr("RH01_DH_CADASTRO"), eTipoValor.DATA)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            RH01_IN_CONTRIBUI_FUNBEM = DoBanco(dr("RH01_IN_CONTRIBUI_FUNBEM"), eTipoValor.BOOLEANO)
            CA04_ID_USUARIO_LOGIN = DoBanco(dr("CA04_ID_USUARIO_LOGIN"), eTipoValor.CHAVE)

        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function pesquisaPrincipal(Optional ByVal Sort As String = "", Optional ByVal Regional As Integer = 0, Optional ByVal Cidade As Integer = 0, Optional ByVal Lotacao As Integer = 0, Optional ByVal Cargo As Integer = 0, Optional ByVal CargaHoraria As Integer = 0, Optional ByVal Vinculo As Integer = 0, Optional ByVal Turno As Integer = 0, Optional ByVal Professor As Boolean = False, Optional ByVal Reducao As Boolean = False, Optional Ampliacao As Boolean = False, Optional Lotados As Integer = 0, Optional Enturmados As Integer = 0, Optional ByVal Orgao As Integer = 0, Optional RestoCargaHoraria As Boolean = False, Optional Cpf As String = "", Optional Situacao As Integer = 0, Optional Nome As String = "") As Data.DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select distinct * , isnull(QTD_HORAS_ALOCADA,0)-isnull(QTD_HORAS_ENTURMADO,0) as QTD_HORAS_VAGAS from ( ")
        strSQL.Append(" select ")
        strSQL.Append(" tg05.TG05_NM_REGIONAL as REGIONAL, ")
        strSQL.Append(" tg03.TG03_NM_MUNICIPIO AS CIDADE, ")
        strSQL.Append(" rh36.RH36_CD_INEP_LOTACAO AS INEP, ")
        strSQL.Append(" ISNULL(RH48.RH48_NM_TIPO_ESCOLA,'')+rh36.RH36_NM_LOTACAO AS LOTACAO, ")
        strSQL.Append(" RH02.RH02_CD_MATRICULA AS MATRICULA, ")
        strSQL.Append(" RH01.RH01_NU_CPF AS CPF, ")
        strSQL.Append(" RH01.RH01_NM_PESSOA AS SERVIDOR, ")
        strSQL.Append(" RH16.RH16_NM_CARGO AS CARGO, ")
        strSQL.Append(" RH79.RH79_NM_TIPO_CARGA_HORARIA AS CH, ")
        strSQL.Append(" RH05.RH05_NM_TIPO_VINCULO AS VINCULO, ")
        strSQL.Append(" TG06.TG06_NM_TURNO AS TURNO, ")
        strSQL.Append(" (SELECT SUM(RH74.RH74_QT_HORA_ALOCADA) FROM RH74_HABILITACAO RH74 ")
        strSQL.Append(" WHERE RH74.RH80_ID_ALOCACAO_CARGA_HORARIA = rh80.RH80_ID_ALOCACAO_CARGA_HORARIA ")
        strSQL.Append(" AND RH74.RH74_DH_DESATIVACAO IS NULL and rh80.RH80_DH_DESATIVACAO is null) AS QTD_HORAS_ALOCADA, ")
        strSQL.Append(" (SELECT COUNT(*) FROM DBDIARIO..DE13_HORARIO_TURMA DE13 ")
        strSQL.Append(" WHERE rh80.RH80_ID_ALOCACAO_CARGA_HORARIA = DE13.RH80_ID_ALOCACAO_CARGA_HORARIA ")
        strSQL.Append(" AND DE13.DE13_DH_EXCLUSAO IS NULL AND DE13.RH80_ID_ALOCACAO_CARGA_HORARIA IS NOT NULL) AS QTD_HORAS_ENTURMADO, RH01.RH01_ID_PESSOA ")
        strSQL.Append(" from rh01_pessoa rh01 ")

        strSQL.Append(" INNER JOIN RH02_SERVIDOR rh02 on rh01.RH01_ID_PESSOA = rh02.RH01_ID_PESSOA ")
        strSQL.Append(" LEFT JOIN RH05_TIPO_VINCULO RH05 ON RH02.RH05_ID_TIPO_VINCULO = RH05.RH05_ID_TIPO_VINCULO ")
        strSQL.Append(" LEFT JOIN RH16_CARGO RH16 ON rh02.RH16_ID_CARGO = RH16.RH16_ID_CARGO ")
        strSQL.Append(" LEFT JOIN RH04_ORGAO RH04 ON RH16.RH04_ID_ORGAO = RH04.RH04_ID_ORGAO ")
        strSQL.Append(" LEFT JOIN RH14_LOTACAO_SERVIDOR rh14 on rh02.RH02_ID_SERVIDOR = rh14.RH02_ID_SERVIDOR ")
        strSQL.Append(" LEFT JOIN RH36_LOTACAO rh36 on rh14.RH36_ID_LOTACAO = rh36.RH36_ID_LOTACAO ")
        strSQL.Append(" LEFT JOIN dbgeral..TG03_MUNICIPIO tg03 on rh36.TG03_ID_MUNICIPIO = tg03.TG03_ID_MUNICIPIO ")
        strSQL.Append(" LEFT JOIN RH48_TIPO_ESCOLA RH48 ON rh36.RH48_ID_TIPO_ESCOLA = RH48.RH48_ID_TIPO_ESCOLA ")
        strSQL.Append(" LEFT JOIN dbgeral..TG05_REGIONAL tg05 on tg03.TG05_ID_REGIONAL = tg05.TG05_ID_REGIONAL ")


        strSQL.Append(" LEFT JOIN RH80_ALOCACAO_CARGA_HORARIA	rh80 on rh80.RH14_ID_LOTACAO_SERVIDOR = rh14.RH14_ID_LOTACAO_SERVIDOR ")
        strSQL.Append(" LEFT JOIN RH78_SERVIDOR_CARGA_HORARIA	rh78 on rh80.RH78_ID_SERVIDOR_CARGA_HORARIA = rh78.RH78_ID_SERVIDOR_CARGA_HORARIA ")
        strSQL.Append(" LEFT JOIN RH77_CARGA_HORARIA			RH77 ON	RH77.RH77_ID_CARGA_HORARIA = rh78.RH77_ID_CARGA_HORARIA ")
        strSQL.Append(" LEFT JOIN RH79_TIPO_CARGA_HORARIA		RH79 ON RH77.RH79_ID_TIPO_CARGA_HORARIA = RH79.RH79_ID_TIPO_CARGA_HORARIA ")
        strSQL.Append(" LEFT JOIN DBGERAL..TG06_TURNO			TG06 ON rh80.TG06_ID_TURNO = TG06.TG06_ID_TURNO ")
        'strSQL.Append(" LEFT JOIN RH74_HABILITACAO				RH74 ON rh80.RH80_ID_ALOCACAO_CARGA_HORARIA = RH74.RH80_ID_ALOCACAO_CARGA_HORARIA ")

        'strSQL.Append(" LEFT JOIN RH78_SERVIDOR_CARGA_HORARIA rh78 on rh02.RH02_ID_SERVIDOR = rh78.RH02_ID_SERVIDOR ")
        'strSQL.Append(" LEFT JOIN RH77_CARGA_HORARIA RH77 ON RH77.RH77_ID_CARGA_HORARIA = rh78.RH77_ID_CARGA_HORARIA ")
        'strSQL.Append(" LEFT JOIN RH79_TIPO_CARGA_HORARIA RH79 ON RH77.RH79_ID_TIPO_CARGA_HORARIA = RH79.RH79_ID_TIPO_CARGA_HORARIA ")
        'strSQL.Append(" LEFT JOIN RH80_ALOCACAO_CARGA_HORARIA rh80 on rh78.RH78_ID_SERVIDOR_CARGA_HORARIA = rh80.RH78_ID_SERVIDOR_CARGA_HORARIA ")
        'strSQL.Append(" LEFT JOIN DBGERAL..TG06_TURNO TG06 ON rh80.TG06_ID_TURNO = TG06.TG06_ID_TURNO ")

        strSQL.Append(" left join dbdiario..DE13_HORARIO_TURMA de13 on rh80.RH80_ID_ALOCACAO_CARGA_HORARIA = de13.RH80_ID_ALOCACAO_CARGA_HORARIA ")
        strSQL.Append(" WHERE rh01.rh01_id_pessoa > 0 ")

        If Situacao = 6 Or Situacao = 12 Or Situacao = 13 Then


            strSQL.Append(" AND RH02.RH07_ID_SITUACAO_SERVIDOR = " & Situacao)
            'strSQL.Append(" And RH14.RH14_ID_LOTACAO_SERVIDOR Is NULL ")


        ElseIf Situacao = 1 Then


            strSQL.Append(" AND RH02.RH07_ID_SITUACAO_SERVIDOR in (1,10,11) ")
            'strSQL.Append(" AND RH78_DH_DESATIVACAO IS NULL ")
            'strSQL.Append(" AND RH80_DH_DESATIVACAO IS NULL ")
            'strSQL.Append(" AND DE13.DE13_DH_EXCLUSAO is NULL ")

        Else

            strSQL.Append(" AND RH02.RH07_ID_SITUACAO_SERVIDOR = " & Situacao)
            'strSQL.Append(" And RH78_DH_DESATIVACAO Is NULL ")
            'strSQL.Append(" And RH80_DH_DESATIVACAO Is NULL ")
            'strSQL.Append(" And DE13.DE13_DH_EXCLUSAO Is NULL ")

        End If


        If Regional > 0 Then
            strSQL.Append(" And TG05.TG05_ID_REGIONAL = " & Regional)
        End If

        If Cidade > 0 Then
            strSQL.Append(" And TG03.TG03_ID_MUNICIPIO = " & Cidade)
        End If

        If Lotacao > 0 Then
            strSQL.Append(" And rh36.RH36_ID_LOTACAO = " & Lotacao)
        End If

        If Cargo > 0 Then
            strSQL.Append(" And RH16.RH16_ID_CARGO = " & Cargo)
        End If

        If CargaHoraria > 0 Then
            strSQL.Append(" And RH79.RH79_ID_TIPO_CARGA_HORARIA = " & CargaHoraria)
        End If

        If Vinculo > 0 Then
            strSQL.Append(" And RH05.RH05_ID_TIPO_VINCULO = " & Vinculo)
        End If

        If Professor Then
            strSQL.Append(" And rh16.RH85_ID_CATEGORIA_CARGO = 1")
        End If

        If Reducao Then
            strSQL.Append(" And RH02_IN_REDUCAO_CG_HORARIA = 1")
        End If

        If Nome <> "" Then
            strSQL.Append(" And upper(RH01_NM_PESSOA) Like '%" & Nome.ToUpper & "%' collate sql_latin1_general_cp1251_cs_as")
        End If


        If Ampliacao Then
            strSQL.Append(" AND RH02_IN_AMPLIACAO_CG_HORARIA = 1")
        End If

        If Orgao > 0 Then
            strSQL.Append(" AND RH04.RH04_ID_ORGAO =" & Orgao)
        Else
            strSQL.Append(" AND RH04.RH04_ID_ORGAO IN (5,6,7,8) ")
        End If

        Select Case Lotados

            Case 1
                strSQL.Append(" AND RH14.RH14_ID_LOTACAO_SERVIDOR IS NOT NULL ")

            Case 2
                strSQL.Append(" AND RH14.RH14_ID_LOTACAO_SERVIDOR IS NULL ")
        End Select

        Select Case Enturmados

            Case 1
                strSQL.Append(" and de13.RH80_ID_ALOCACAO_CARGA_HORARIA is not null ")
                strSQL.Append(" And RH78_DH_DESATIVACAO Is NULL ")
                strSQL.Append(" And RH80_DH_DESATIVACAO Is NULL ")
                strSQL.Append(" And DE13.DE13_DH_EXCLUSAO Is NULL ")

            Case 2
                strSQL.Append(" and de13.RH80_ID_ALOCACAO_CARGA_HORARIA is null ")
                strSQL.Append(" And RH78_DH_DESATIVACAO Is NULL ")
                strSQL.Append(" And RH80_DH_DESATIVACAO Is NULL ")
        End Select

        If Replace(Replace(Replace(Cpf, ".", ""), "-", ""), "_", "") <> "" Then
            strSQL.Append(" and upper(RH01_NU_CPF) = '" & Replace(Replace(Cpf.ToUpper, ".", ""), "-", "") & "'")
        End If


        strSQL.Append(" ) as tab1 ")

        If RestoCargaHoraria Then
            strSQL.Append(" where isnull(QTD_HORAS_ALOCADA,0)-isnull(QTD_HORAS_ENTURMADO,0) > 0 ")
        End If

        strSQL.Append(" ORDER BY 1,2,3,4,5 ")

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarPessoaRecadastramentoGestao(Optional ByVal Sort As String = "", Optional ByVal MesRecadastramento As Integer = 0 _
                                              , Optional ByVal Regional As Integer = 0, Optional ByVal Cidade As Integer = 0, Optional ByVal Lotacao As Integer = 0 _
                                              , Optional PessoaNome As String = "", Optional Cpf As String = "", Optional ByVal Recadastrados As Boolean = False _
                                              , Optional ByVal Ano As Integer = 0) As DataTable

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder


        strSQL.Append("  select  RH86_ID_RECADASTRAMENTO,TG05.TG05_NM_REGIONAL, TG03.TG03_NM_MUNICIPIO, RH36.RH36_NM_LOTACAO, RH02.RH02_CD_MATRICULA, RH01_NM_PESSOA, RH86.RH86_DH_CADASTRO, RH01.RH01_DT_NASCIMENTO	, RH86_NR_ANO_RECADASTRAMENTO  ")
        strSQL.Append("  	from 	RH01_PESSOA					RH01	  ")
        strSQL.Append("  		left	join	RH86_RECADASTRAMENTO		rh86	on rh01.RH01_ID_PESSOA = rh86.RH01_ID_PESSOA	  ")
        strSQL.Append("  		LEFT	JOIN	RH02_SERVIDOR				RH02	ON	RH01.RH01_ID_PESSOA		=	RH02.RH01_ID_PESSOA	  ")
        strSQL.Append("  		LEFT	JOIN	RH14_LOTACAO_SERVIDOR		RH14	ON	RH02.RH02_ID_SERVIDOR	=	RH14.RH02_ID_SERVIDOR	  ")
        strSQL.Append("         left    join    RH88_PERIODO                RH88    on  RH88.RH88_ID_PERIODO    =   RH14.RH88_ID_PERIODO ")
        strSQL.Append("  		LEFT	JOIN	RH36_LOTACAO				RH36	ON	RH14.RH36_ID_LOTACAO	=	RH36.RH36_ID_LOTACAO		  ")
        strSQL.Append("  		LEFT	JOIN	DBGERAL..TG03_MUNICIPIO		TG03	ON	RH36.TG03_ID_MUNICIPIO	=	TG03.TG03_ID_MUNICIPIO	  ")
        strSQL.Append("  		LEFT	JOIN	DBGERAL..TG05_REGIONAL		TG05	ON	TG03.TG05_ID_REGIONAL	=	TG05.TG05_ID_REGIONAL	  ")
        strSQL.Append("  		WHERE RH01.RH01_ID_PESSOA IS NOT NULL AND RH02.RH07_ID_SITUACAO_SERVIDOR IN (1,10,11) and rh14.RH14_DT_DESLIGAMENTO is null	  ")
        strSQL.Append("  		and RH02.RH04_ID_ORGAO IN (5,6,7,8)	  ")
        ' strSQL.Append("         And RH88.RH88_NM_PERIODO = YEAR(GETDATE()) ")

        If Regional > 0 Then
            strSQL.Append("  and TG05.TG05_ID_REGIONAL =  " & Regional)
        End If

        If Ano > 0 Then
            strSQL.Append("  and RH86_NR_ANO_RECADASTRAMENTO  =  " & Ano)
        Else
            strSQL.Append("  and RH86_NR_ANO_RECADASTRAMENTO  =  " & Date.Now.Year)
        End If

        If Cidade > 0 Then
            strSQL.Append("  and tg03.TG03_ID_MUNICIPIO =  " & Cidade)

        End If

        If Lotacao > 0 Then
            strSQL.Append(" and RH36.RH36_ID_LOTACAO =  " & Lotacao)
        End If

        If MesRecadastramento > 0 Then
            strSQL.Append(" and MONTH(RH01.RH01_DT_NASCIMENTO) =  " & MesRecadastramento)
        Else
            strSQL.Append(" and MONTH(RH01.RH01_DT_NASCIMENTO) =  month(getdate())")
        End If


        If PessoaNome <> "" Then
            strSQL.Append(" And  RH01_NM_PESSOA Like '%" & PessoaNome & "%' collate sql_latin1_general_cp1251_cs_as")
        End If

        If Replace(Replace(Replace(Cpf, ".", ""), "-", ""), "_", "") <> "" Then
            strSQL.Append(" and RH01_NU_CPF = '" & Replace(Replace(Cpf, ".", ""), "-", "") & "'")
        End If

        If Recadastrados Then
            strSQL.Append(" and RH86_ID_RECADASTRAMENTO is not null")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH86_NR_ANO_RECADASTRAMENTO DESC", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)

    End Function

    Public Function PesquisarLotacao(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional CPF As String = "", Optional Nome As String = "", Optional NomeMae As String = "", Optional NomePai As String = "", Optional Sexo As String = "", Optional DDDTelefone As String = "", Optional NumeroTelefone As String = "", Optional DDDCelular As String = "", Optional NumeroCelular As String = "", Optional Email As String = "", Optional DataNascimento As String = "", Optional AtestadoFisicoMental As String = "", Optional DeclaracaoAcumulacaoCargo As String = "", Optional DeclaracaoBens As String = "", Optional DataHotaCadastro As String = "", Optional Usuario As Integer = 0, Optional ContruibuiFunbem As String = "", Optional UsuarioLogin As Integer = 0, Optional Lotacao As Integer = 0, Optional Top As Integer = 0, Optional Perfil As Integer = 0, Optional LotacaoPrincipal As Boolean = False) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select distinct top " & Top.ToString & " RH01.RH01_ID_PESSOA, RH01_NM_PESSOA, RH16_NM_CARGO, RH14.RH36_ID_LOTACAO, rh02.rh02_cd_matricula ")
        strSQL.Append(" From  RH01_PESSOA As RH01 ")
        strSQL.Append(" Left Join RH02_SERVIDOR as RH02 on RH02.RH01_ID_PESSOA = RH01.RH01_ID_PESSOA ")
        strSQL.Append(" Left Join RH16_CARGO as RH16 on RH16.RH16_ID_CARGO = RH02.RH16_ID_CARGO ")
        strSQL.Append(" Left Join RH14_LOTACAO_SERVIDOR as RH14 on RH14.RH02_ID_SERVIDOR = RH02.RH02_ID_SERVIDOR ")
        strSQL.Append(" where RH01.RH01_ID_PESSOA Is Not null ")
        strSQL.Append(" and RH07_ID_SITUACAO_SERVIDOR in (1,10,11)  ")
        strSQL.Append(" And RH14_DT_DESLIGAMENTO Is null ")

        If Codigo > 0 Then
            strSQL.Append(" And RH01_ID_PESSOA = " & Codigo)
        End If

        If CPF <> "" Then
            strSQL.Append(" And Replace(Replace(upper(RH01_NU_CPF),'.',''),'-','') like '%" & CPF.ToUpper.Replace(".", "").Replace("-", "") & "%'")
        End If

        If LotacaoPrincipal Then
            strSQL.Append(" And rh14.RH14_TP_LOTACAO_SERVIDOR = 'P' ")
        End If

        If Lotacao > 0 Then
            strSQL.Append(" And RH36_ID_LOTACAO = " & Lotacao)
        End If

        If Nome <> "" Then
            strSQL.Append(" And upper(RH01_NM_PESSOA) like '%" & Nome.ToUpper & "%' collate sql_latin1_general_cp1251_cs_as")
        End If

        If NomeMae <> "" Then
            strSQL.Append(" and upper(RH01_NM_MAE) like '%" & NomeMae.ToUpper & "%'")
        End If

        If NomePai <> "" Then
            strSQL.Append(" and upper(RH01_NM_PAI) like '%" & NomePai.ToUpper & "%'")
        End If

        If Sexo <> "" Then
            strSQL.Append(" and upper(RH01_TP_SEXO) like '%" & Sexo.ToUpper & "%'")
        End If

        If DDDTelefone <> "" Then
            strSQL.Append(" and upper(RH01_NU_DDD_TELEFONE) like '%" & DDDTelefone.ToUpper & "%'")
        End If

        If NumeroTelefone <> "" Then
            strSQL.Append(" and upper(RH01_NU_TELEFONE) like '%" & NumeroTelefone.ToUpper & "%'")
        End If

        If DDDCelular <> "" Then
            strSQL.Append(" and upper(RH01_NU_DDD_CELULAR) like '%" & DDDCelular.ToUpper & "%'")
        End If

        If NumeroCelular <> "" Then
            strSQL.Append(" and upper(RH01_NU_CELULAR) like '%" & NumeroCelular.ToUpper & "%'")
        End If

        If Email <> "" Then
            strSQL.Append(" and upper(RH01_NM_EMAIL) like '%" & Email.ToUpper & "%'")
        End If

        If DataNascimento <> "" Then
            strSQL.Append(" and upper(RH01_DT_NASCIMENTO) like '%" & DataNascimento.ToUpper & "%'")
        End If

        If AtestadoFisicoMental <> "" Then
            strSQL.Append(" and upper(RH01_IN_ATESTADO_FISICO_MENTAL) like '%" & AtestadoFisicoMental.ToUpper & "%'")
        End If

        If DeclaracaoAcumulacaoCargo <> "" Then
            strSQL.Append(" and upper(RH01_IN_DECLARA_ACUMULACAO_CARGO) like '%" & DeclaracaoAcumulacaoCargo.ToUpper & "%'")
        End If

        If DeclaracaoBens <> "" Then
            strSQL.Append(" and upper(RH01_IN_DECLARACAO_BENS) like '%" & DeclaracaoBens.ToUpper & "%'")
        End If

        If IsDate(DataHotaCadastro) Then
            strSQL.Append(" and RH01_DH_CADASTRO = Convert(DateTime, '" & DataHotaCadastro & "', 103)")
        End If

        If Usuario > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & Usuario)
        End If

        If ContruibuiFunbem <> "" Then
            strSQL.Append(" and upper(RH01_IN_CONTRIBUI_FUNBEM) like '%" & ContruibuiFunbem.ToUpper & "%'")
        End If

        If UsuarioLogin > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO_LOGIN = " & UsuarioLogin)
        End If

        If Perfil > 0 Then
            strSQL.Append(" and RH01.CA04_ID_USUARIO_LOGIN in (select FKCA08CA04_COD_USUARIO from [172.16.2.71].DBCONTROLEACESSO.DBO.CA08_PERFIL_USUARIO ")
            strSQL.Append(" where FKCA08CA01_COD_APLICACAO = 87 and FKCA08CA03_COD_PERFIL = " & Perfil & " ) ")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH01_NM_PESSOA", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarRecadastramentoAnual(Optional ByVal Sort As String = "", Optional PessoaId As Integer = 0, Optional Cpf As String = "", Optional Nome As String = "", Optional Mae As String = "", Optional Pai As String = "", Optional Sexo As String = "", Optional DddTelefone As String = "", Optional Telefone As String = "", Optional DddCelular As String = "", Optional Celular As String = "", Optional Email As String = "", Optional DataNascimento As String = "", Optional AtestadoFisicoMental As String = "", Optional DeclaraAcumulacaoCargo As String = "", Optional DeclaracaoBens As String = "", Optional DataHoraCadastro As String = "", Optional UsuarioId As Integer = 0, Optional ContribuicaoFunbem As String = "", Optional ByVal Regional As Integer = 0, Optional ByVal Cidade As Integer = 0, Optional ByVal Mes As Integer = 0, Optional ByVal Ano As Integer = 0, Optional ByVal NaoRecadastro As Integer = 0) As DataTable

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select top 1 RH86_ID_RECADASTRAMENTO,RH01.RH01_ID_PESSOA, RH01.RH01_NM_PESSOA AS SERVIDOR, TG05.TG05_NM_REGIONAL AS REGIONAL, TG03.TG03_NM_MUNICIPIO AS CIDADE, RH36.RH36_NM_LOTACAO AS LOTACAO,RH86_NR_ANO_RECADASTRAMENTO ")
        strSQL.Append(" from RH01_PESSOA rh01")
        strSQL.Append(" left join RH02_SERVIDOR RH02 on RH01.RH01_ID_PESSOA = RH02.RH01_ID_PESSOA    ")
        strSQL.Append(" left join RH14_LOTACAO_SERVIDOR RH14 on RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR  and rh14.RH14_DT_DESLIGAMENTO is null")
        strSQL.Append(" left join RH88_PERIODO RH88 on RH88.RH88_ID_PERIODO = RH14.RH88_ID_PERIODO ")
        strSQL.Append(" left join RH36_LOTACAO RH36 on RH14.RH36_ID_LOTACAO = RH36.RH36_ID_LOTACAO ")
        strSQL.Append(" left join DBGERAL..TG03_MUNICIPIO TG03 on RH36.TG03_ID_MUNICIPIO = TG03.TG03_ID_MUNICIPIO ")
        strSQL.Append(" left join DBGERAL..TG05_REGIONAL TG05 on TG03.TG05_ID_REGIONAL = TG05.TG05_ID_REGIONAL ")
        strSQL.Append(" inner join RH07_SITUACAO_SERVIDOR rh07 on rh07.RH07_ID_SITUACAO_SERVIDOR = rh02.RH07_ID_SITUACAO_SERVIDOR ")
        strSQL.Append(" left join RH86_RECADASTRAMENTO RH86 on rh01.RH01_ID_PESSOA = RH86.RH01_ID_PESSOA ")
        strSQL.Append(" where rh01.RH01_ID_PESSOA is not null")
        strSQL.Append(" and rh07.RH07_ID_SITUACAO_SERVIDOR in (1,10,11,12,6,13) ")

        If Mes > 0 Then
            strSQL.Append(" and Month(rh01.RH01_DT_NASCIMENTO) =  " & Mes)
        End If

        If Ano > 0 Then

            strSQL.Append(" and not exists (select * from  RH86_RECADASTRAMENTO RH86_NAORECADASTRADOS ")
            strSQL.Append("	 where RH86_NAORECADASTRADOS.RH01_ID_PESSOA = rh01.RH01_ID_PESSOA ")
            strSQL.Append(" And Year(RH86_NAORECADASTRADOS.RH86_DH_CADASTRO) = " & Ano & ")")

        End If


        strSQL.Append(" and RH02.RH04_ID_ORGAO IN (5,6,7,8) ")
        strSQL.Append(" and rh14.RH14_DT_DESLIGAMENTO is null ")
        'strSQL.Append(" And RH88.RH88_NM_PERIODO = YEAR(GETDATE())")

        If PessoaId > 0 Then
            strSQL.Append(" and rh01.RH01_ID_PESSOA = " & PessoaId)
        End If

        If Replace(Replace(Replace(Cpf, ".", ""), "-", ""), "_", "") <> "" Then
            strSQL.Append(" and upper(RH01_NU_CPF) = '" & Replace(Replace(Cpf.ToUpper, ".", ""), "-", "") & "'")
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(RH01_NM_PESSOA) like '%" & Nome.ToUpper & "%' collate sql_latin1_general_cp1251_cs_as")
        End If

        If Mae <> "" Then
            strSQL.Append(" and upper(RH01_NM_MAE) like '%" & Mae.ToUpper & "%'")
        End If

        If Pai <> "" Then
            strSQL.Append(" and upper(RH01_NM_PAI) like '%" & Pai.ToUpper & "%'")
        End If

        If Sexo <> "" Then
            strSQL.Append(" and upper(RH01_TP_SEXO) like '%" & Sexo.ToUpper & "%'")
        End If

        If DddTelefone <> "" Then
            strSQL.Append(" and upper(RH01_NU_DDD_TELEFONE) like '%" & DddTelefone.ToUpper & "%'")
        End If

        If Telefone <> "" Then
            strSQL.Append(" and upper(RH01_NU_TELEFONE) like '%" & Telefone.ToUpper & "%'")
        End If

        If DddCelular <> "" Then
            strSQL.Append(" and upper(RH01_NU_DDD_CELULAR) like '%" & DddCelular.ToUpper & "%'")
        End If

        If Celular <> "" Then
            strSQL.Append(" and upper(RH01_NU_CELULAR) like '%" & Celular.ToUpper & "%'")
        End If

        If Email <> "" Then
            strSQL.Append(" and upper(RH01_NM_EMAIL) like '%" & Email.ToUpper & "%'")
        End If

        If DataNascimento <> "" Then
            strSQL.Append(" and upper(RH01_DT_NASCIMENTO) like '%" & DataNascimento.ToUpper & "%'")
        End If

        If AtestadoFisicoMental <> "" Then
            strSQL.Append(" and upper(RH01_IN_ATESTADO_FISISCO_MENTAL) like '%" & AtestadoFisicoMental.ToUpper & "%'")
        End If

        If DeclaraAcumulacaoCargo <> "" Then
            strSQL.Append(" and upper(RH01_IN_DECLARA_ACUMULACAO_CARGO) like '%" & DeclaraAcumulacaoCargo.ToUpper & "%'")
        End If

        If DeclaracaoBens <> "" Then
            strSQL.Append(" and upper(RH01_IN_DECLARACAO_BENS) like '%" & DeclaracaoBens.ToUpper & "%'")
        End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" and RH01_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
        End If

        If UsuarioId > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & UsuarioId)
        End If

        If Regional > 0 Then
            strSQL.Append(" and TG05.TG05_ID_REGIONAL = " & Regional)
        End If

        If Cidade > 0 Then
            strSQL.Append(" and TG03.TG03_ID_MUNICIPIO = " & Cidade)
        End If

        If ContribuicaoFunbem <> "" Then
            strSQL.Append(" and upper(RH01_IN_CONTRIBUI_FUNBEM) like '%" & ContribuicaoFunbem.ToUpper & "%'")
        End If


        strSQL.Append(" Order By " & IIf(Sort = "", "RH86_NR_ANO_RECADASTRAMENTO desc", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarRecadastramento(Optional ByVal Sort As String = "", Optional PessoaId As Integer = 0, Optional Cpf As String = "", Optional Nome As String = "", Optional Mae As String = "", Optional Pai As String = "", Optional Sexo As String = "", Optional DddTelefone As String = "", Optional Telefone As String = "", Optional DddCelular As String = "", Optional Celular As String = "", Optional Email As String = "", Optional DataNascimento As String = "", Optional AtestadoFisicoMental As String = "", Optional DeclaraAcumulacaoCargo As String = "", Optional DeclaracaoBens As String = "", Optional DataHoraCadastro As String = "", Optional UsuarioId As Integer = 0, Optional ContribuicaoFunbem As String = "", Optional ByVal Regional As Integer = 0, Optional ByVal Cidade As Integer = 0, Optional ByVal Mes As Integer = 0, Optional ByVal Ano As Integer = 0, Optional ByVal NaoRecadastro As Integer = 0) As DataTable

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select top 1 RH86_ID_RECADASTRAMENTO,RH01.RH01_ID_PESSOA, RH01.RH01_NM_PESSOA AS SERVIDOR, TG05.TG05_NM_REGIONAL AS REGIONAL, TG03.TG03_NM_MUNICIPIO AS CIDADE, RH36.RH36_NM_LOTACAO AS LOTACAO, RH86_NR_ANO_RECADASTRAMENTO ")
        strSQL.Append(" from RH01_PESSOA rh01")
        strSQL.Append(" left join RH02_SERVIDOR RH02 on RH01.RH01_ID_PESSOA = RH02.RH01_ID_PESSOA    ")
        strSQL.Append(" left join RH14_LOTACAO_SERVIDOR RH14 on RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR  and rh14.RH14_DT_DESLIGAMENTO is null   ")
        strSQL.Append(" left join RH88_PERIODO RH88 on RH88.RH88_ID_PERIODO = RH14.RH88_ID_PERIODO ")
        strSQL.Append(" left join RH36_LOTACAO RH36 on RH14.RH36_ID_LOTACAO = RH36.RH36_ID_LOTACAO ")
        strSQL.Append(" left join DBGERAL..TG03_MUNICIPIO TG03 on RH36.TG03_ID_MUNICIPIO = TG03.TG03_ID_MUNICIPIO ")
        strSQL.Append(" left join DBGERAL..TG05_REGIONAL TG05 on TG03.TG05_ID_REGIONAL = TG05.TG05_ID_REGIONAL ")
        strSQL.Append(" inner join RH07_SITUACAO_SERVIDOR rh07 on rh07.RH07_ID_SITUACAO_SERVIDOR = rh02.RH07_ID_SITUACAO_SERVIDOR ")
        strSQL.Append(" left join RH86_RECADASTRAMENTO RH86 on rh01.RH01_ID_PESSOA = RH86.RH01_ID_PESSOA ")
        strSQL.Append(" where rh01.RH01_ID_PESSOA is not null ")
        strSQL.Append(" and month(RH01_DT_NASCIMENTO) = Month(getdate())")
        strSQL.Append(" And rh07.RH07_ID_SITUACAO_SERVIDOR In (1,10,11,12,6,13) ")
        'strSQL.Append(" And RH88.RH88_NM_PERIODO = YEAR(GETDATE())")

        '||||||| .r11889
        '        strSQL.Append(" And month(RH01_DT_NASCIMENTO) = month(getdate())")
        '        strSQL.Append(" And rh07.RH07_IN_SERVIDOR_ATIVO = 1 ")
        '=======
        '        If Mes = 0 And Ano = 0 Then
        '            strSQL.Append(" And month(RH01_DT_NASCIMENTO) = month(getdate())")
        '        End If
        '        strSQL.Append(" And rh07.RH07_IN_SERVIDOR_ATIVO = 1 ")
        '>>>>>>> .r11930


        strSQL.Append(" And RH02.RH04_ID_ORGAO In (5,6,7,8) ")
        strSQL.Append(" And rh14.RH14_DT_DESLIGAMENTO Is null ")

        If PessoaId > 0 Then
            strSQL.Append(" And rh01.RH01_ID_PESSOA = " & PessoaId)
        End If

        If Replace(Replace(Replace(Cpf, ".", ""), "-", ""), "_", "") <> "" Then
            'strSQL.Append(" And upper(RH01_NU_CPF) Like '%" & replace(replace(Cpf.toUpper,".",""),"-","") & "%'")
            strSQL.Append(" and upper(RH01_NU_CPF) = '" & Replace(Replace(Cpf.ToUpper, ".", ""), "-", "") & "'")
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(RH01_NM_PESSOA) like '%" & Nome.ToUpper & "%' collate sql_latin1_general_cp1251_cs_as")
        End If

        If Mae <> "" Then
            strSQL.Append(" and upper(RH01_NM_MAE) like '%" & Mae.ToUpper & "%'")
        End If

        If Pai <> "" Then
            strSQL.Append(" and upper(RH01_NM_PAI) like '%" & Pai.ToUpper & "%'")
        End If

        If Sexo <> "" Then
            strSQL.Append(" and upper(RH01_TP_SEXO) like '%" & Sexo.ToUpper & "%'")
        End If

        If DddTelefone <> "" Then
            strSQL.Append(" and upper(RH01_NU_DDD_TELEFONE) like '%" & DddTelefone.ToUpper & "%'")
        End If

        If Telefone <> "" Then
            strSQL.Append(" and upper(RH01_NU_TELEFONE) like '%" & Telefone.ToUpper & "%'")
        End If

        If DddCelular <> "" Then
            strSQL.Append(" and upper(RH01_NU_DDD_CELULAR) like '%" & DddCelular.ToUpper & "%'")
        End If

        If Celular <> "" Then
            strSQL.Append(" and upper(RH01_NU_CELULAR) like '%" & Celular.ToUpper & "%'")
        End If

        If Email <> "" Then
            strSQL.Append(" and upper(RH01_NM_EMAIL) like '%" & Email.ToUpper & "%'")
        End If

        If DataNascimento <> "" Then
            strSQL.Append(" and upper(RH01_DT_NASCIMENTO) like '%" & DataNascimento.ToUpper & "%'")
        End If

        If AtestadoFisicoMental <> "" Then
            strSQL.Append(" and upper(RH01_IN_ATESTADO_FISISCO_MENTAL) like '%" & AtestadoFisicoMental.ToUpper & "%'")
        End If

        If DeclaraAcumulacaoCargo <> "" Then
            strSQL.Append(" and upper(RH01_IN_DECLARA_ACUMULACAO_CARGO) like '%" & DeclaraAcumulacaoCargo.ToUpper & "%'")
        End If

        If DeclaracaoBens <> "" Then
            strSQL.Append(" and upper(RH01_IN_DECLARACAO_BENS) like '%" & DeclaracaoBens.ToUpper & "%'")
        End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" and RH01_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
        End If

        If UsuarioId > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & UsuarioId)
        End If

        If Regional > 0 Then
            strSQL.Append(" and TG05.TG05_ID_REGIONAL = " & Regional)
        End If

        If Cidade > 0 Then
            strSQL.Append(" and TG03.TG03_ID_MUNICIPIO = " & Cidade)
        End If

        If ContribuicaoFunbem <> "" Then
            strSQL.Append(" and upper(RH01_IN_CONTRIBUI_FUNBEM) like '%" & ContribuicaoFunbem.ToUpper & "%'")
        End If

        '<<<<<<< .mine
        '        If Mes > 0 Then
        '            strSQL.Append(" and Month(rh01.RH01_DT_NASCIMENTO) =  " & Mes)
        '        End If

        '        If Ano > 0 Then
        '            strSQL.Append(" and Year(RH86.RH86_DH_CADASTRO) =  " & Ano)
        '        End If

        '||||||| .r11889

        'If Mes > 0 Then
        '    strSQL.Append(" and Month(rh01.RH01_DT_NASCIMENTO) =  " & Mes)
        'End If

        'If Ano > 0 And NaoRecadastro > 0 Then
        '    strSQL.Append("and (Year(RH86.RH86_DH_CADASTRO)  <> (" & Ano & ")  or RH86.RH86_ID_RECADASTRAMENTO is null )")
        'ElseIf Ano = 0 And NaoRecadastro > 0 Then
        '    strSQL.Append(" and RH86.RH86_ID_RECADASTRAMENTO is null ")
        'ElseIf Ano > 0 And NaoRecadastro = 0 Then
        '    strSQL.Append("and Year(RH86.RH86_DH_CADASTRO) = " & Ano)
        'End If


        strSQL.Append(" Order By " & IIf(Sort = "", "RH86_NR_ANO_RECADASTRAMENTO desc", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional PessoaId As Integer = 0, Optional Cpf As String = "", Optional Nome As String = "", Optional Mae As String = "", Optional Pai As String = "", Optional Sexo As String = "", Optional DddTelefone As String = "", Optional Telefone As String = "", Optional DddCelular As String = "", Optional Celular As String = "", Optional Email As String = "", Optional DataNascimento As String = "", Optional AtestadoFisicoMental As String = "", Optional DeclaraAcumulacaoCargo As String = "", Optional DeclaracaoBens As String = "", Optional DataHoraCadastro As String = "", Optional UsuarioId As Integer = 0, Optional ContribuicaoFunbem As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH01_PESSOA")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where RH01_ID_PESSOA is not null")

        If PessoaId > 0 Then
            strSQL.Append(" and RH01_ID_PESSOA = " & PessoaId)
        End If

        If Replace(Replace(Replace(Cpf, ".", ""), "-", ""), "_", "") <> "" Then
            'strSQL.Append(" and upper(RH01_NU_CPF) like '%" & replace(replace(Cpf.toUpper,".",""),"-","") & "%'")
            strSQL.Append(" and upper(RH01_NU_CPF) = '" & Replace(Replace(Cpf.ToUpper, ".", ""), "-", "") & "'")
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(RH01_NM_PESSOA) like '%" & Nome.ToUpper & "%' collate sql_latin1_general_cp1251_cs_as")
        End If

        If Mae <> "" Then
            strSQL.Append(" and upper(RH01_NM_MAE) like '%" & Mae.ToUpper & "%'")
        End If

        If Pai <> "" Then
            strSQL.Append(" and upper(RH01_NM_PAI) like '%" & Pai.ToUpper & "%'")
        End If

        If Sexo <> "" Then
            strSQL.Append(" and upper(RH01_TP_SEXO) like '%" & Sexo.ToUpper & "%'")
        End If

        If DddTelefone <> "" Then
            strSQL.Append(" and upper(RH01_NU_DDD_TELEFONE) like '%" & DddTelefone.ToUpper & "%'")
        End If

        If Telefone <> "" Then
            strSQL.Append(" and upper(RH01_NU_TELEFONE) like '%" & Telefone.ToUpper & "%'")
        End If

        If DddCelular <> "" Then
            strSQL.Append(" and upper(RH01_NU_DDD_CELULAR) like '%" & DddCelular.ToUpper & "%'")
        End If

        If Celular <> "" Then
            strSQL.Append(" and upper(RH01_NU_CELULAR) like '%" & Celular.ToUpper & "%'")
        End If

        If Email <> "" Then
            strSQL.Append(" and upper(RH01_NM_EMAIL) like '%" & Email.ToUpper & "%'")
        End If

        If DataNascimento <> "" Then
            strSQL.Append(" and upper(RH01_DT_NASCIMENTO) like '%" & DataNascimento.ToUpper & "%'")
        End If

        If AtestadoFisicoMental <> "" Then
            strSQL.Append(" and upper(RH01_IN_ATESTADO_FISISCO_MENTAL) like '%" & AtestadoFisicoMental.ToUpper & "%'")
        End If

        If DeclaraAcumulacaoCargo <> "" Then
            strSQL.Append(" and upper(RH01_IN_DECLARA_ACUMULACAO_CARGO) like '%" & DeclaraAcumulacaoCargo.ToUpper & "%'")
        End If

        If DeclaracaoBens <> "" Then
            strSQL.Append(" and upper(RH01_IN_DECLARACAO_BENS) like '%" & DeclaracaoBens.ToUpper & "%'")
        End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" and RH01_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
        End If

        If UsuarioId > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & UsuarioId)
        End If

        If ContribuicaoFunbem <> "" Then
            strSQL.Append(" and upper(RH01_IN_CONTRIBUI_FUNBEM) like '%" & ContribuicaoFunbem.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH01_ID_PESSOA", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarHabilidade(ByVal Codigo As Integer) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select RH72.RH72_ID_HABILIDADE, DE09.DE09_NM_DISCIPLINA, RH02.RH02_CD_MATRICULA  ")
        strSQL.Append(", CASE  RH72.RH72_TP_HABILIDADE WHEN 1 THEN 'COMUM'  WHEN 2 THEN 'DESVIO' END AS TP_HABILIDADE ")
        strSQL.Append(" From  RH72_HABILIDADE As RH72 ")
        strSQL.Append(" Left Join DBDIARIO..DE09_DISCIPLINA AS DE09 ON DE09.DE09_ID_DISCIPLINA = RH72.DE09_ID_DISCIPLINA  ")
        strSQL.Append(" Left Join RH02_SERVIDOR AS RH02 ON RH02.RH02_ID_SERVIDOR = RH72.RH02_ID_SERVIDOR  ")
        strSQL.Append(" WHERE RH02.RH01_ID_PESSOA = " & Codigo)
        strSQL.Append(" And RH72.RH72_DH_DESATIVACAO Is null ")
        strSQL.Append(" and RH07_ID_SITUACAO_SERVIDOR in (1,10,11)  ")


        Return cnn.AbrirDataTable(strSQL.ToString)

        cnn = Nothing
        strSQL.Length = 0
    End Function

    Public Function PesquisarHabilitacoes(ByVal Codigo As Integer) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select RH48.RH48_NM_TIPO_ESCOLA + ' ' + RH36.RH36_NM_LOTACAO as RH36_NM_LOTACAO , RH72.RH72_ID_HABILIDADE, DE09.DE09_NM_DISCIPLINA ")
        strSQL.Append("	, RH74.RH74_QT_HORA_ALOCADA ")
        strSQL.Append(" From  RH74_HABILITACAO As RH74 ")
        strSQL.Append(" Left Join RH72_HABILIDADE As RH72 ON RH72.RH72_ID_HABILIDADE = RH74.RH72_ID_HABILIDADE ")
        strSQL.Append(" Left Join DBDIARIO..DE09_DISCIPLINA AS DE09 ON DE09.DE09_ID_DISCIPLINA = RH72.DE09_ID_DISCIPLINA  ")
        strSQL.Append(" Left Join RH02_SERVIDOR AS RH02 ON RH02.RH02_ID_SERVIDOR = RH72.RH02_ID_SERVIDOR ")
        strSQL.Append(" Left Join RH80_ALOCACAO_CARGA_HORARIA as RH80 on RH80.RH80_ID_ALOCACAO_CARGA_HORARIA = RH74.RH80_ID_ALOCACAO_CARGA_HORARIA ")
        strSQL.Append(" Left Join RH14_LOTACAO_SERVIDOR AS RH14 ON RH14.RH14_ID_LOTACAO_SERVIDOR = RH80.RH14_ID_LOTACAO_SERVIDOR ")
        strSQL.Append(" Left Join RH36_LOTACAO  AS RH36 ON RH36.RH36_ID_LOTACAO = RH14.RH36_ID_LOTACAO ")
        strSQL.Append(" Left Join RH48_TIPO_ESCOLA as RH48 on RH48.RH48_ID_TIPO_ESCOLA = RH36.RH48_ID_TIPO_ESCOLA ")
        strSQL.Append(" WHERE RH02.RH01_ID_PESSOA = " & Codigo)
        strSQL.Append(" And RH72.RH72_DH_DESATIVACAO Is null ")
        strSQL.Append(" And RH74.RH74_DH_DESATIVACAO Is NULL ")

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarLotacoesMatricula(ByVal Codigo As Integer) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select isnull(RH48.RH48_NM_TIPO_ESCOLA,'') + ' ' + RH36.RH36_NM_LOTACAO as RH36_NM_LOTACAO, RH36_CD_INEP_LOTACAO, RH02_CD_MATRICULA,  ")
        strSQL.Append(" TG05.TG05_NM_REGIONAL,TG03.TG03_NM_MUNICIPIO ")
        strSQL.Append(" From  RH14_LOTACAO_SERVIDOR As RH14 ")
        strSQL.Append("  inner join RH88_PERIODO rh88 on rh14.RH88_ID_PERIODO = rh88.RH88_ID_PERIODO ")
        strSQL.Append(" Left Join  RH02_SERVIDOR AS RH02 ON RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR ")
        strSQL.Append(" Left Join  RH36_LOTACAO  AS RH36 ON RH36.RH36_ID_LOTACAO =  RH14.RH36_ID_LOTACAO ")
        strSQL.Append(" left join dbgeral..TG03_MUNICIPIO tg03 on rh36.TG03_ID_MUNICIPIO = tg03.TG03_ID_MUNICIPIO ")
        strSQL.Append(" left join dbgeral..TG05_REGIONAL tg05 on tg03.TG05_ID_REGIONAL = tg05.TG05_ID_REGIONAL ")
        strSQL.Append(" Left Join  RH48_TIPO_ESCOLA as RH48 on RH48.RH48_ID_TIPO_ESCOLA = RH36.RH48_ID_TIPO_ESCOLA ")
        strSQL.Append(" WHERE RH02.RH01_ID_PESSOA = " & Codigo)
        strSQL.Append(" And RH07_ID_SITUACAO_SERVIDOR in (1,11,10) ")
        strSQL.Append(" and RH88_NM_PERIODO = year(getdate()) ")
        strSQL.Append(" And RH14_DT_DESLIGAMENTO Is NULL ")

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarPerfil(ByVal Codigo As Integer) As DataTable
	    Dim cnn As New Conexao
	    Dim strSQL As New StringBuilder

	    strSQL.Append(" Select RH01.RH01_ID_PESSOA, RH01.RH01_NU_CPF, RH01.RH01_NM_PESSOA ")
	    strSQL.Append(", Left(RH01.RH01_NU_CPF, 3) + '.' +  substring(RH01.RH01_NU_CPF,4,3) + '.' +  substring(RH01.RH01_NU_CPF,7,3) + '-' + RIGHT(RH01.RH01_NU_CPF, 2) AS CPF ")
	    strSQL.Append(" , RH02.RH02_ID_SERVIDOR, isnull(RH02.RH02_CD_MATRICULA,'') as RH02_CD_MATRICULA ")
        strSQL.Append(", RH67.RH67_ID_DOCUMENTO_PESSOA, RH67.RH67_SG_EXTENSAO_ARQ_IMAGEM, RH67.RH67_IM_PESSOA, RH07_NM_SITUACAO_SERVIDOR ")
        strSQL.Append(" , RH16.RH16_ID_CARGO, RH16.RH16_NM_CARGO ")
        strSQL.Append(", RH79.RH79_NM_TIPO_CARGA_HORARIA ")
        strSQL.Append(", rh07.RH07_NM_SITUACAO_SERVIDOR ")
        strSQL.Append(", rh05.RH05_NM_tipo_vinculo ")
        strSQL.Append("  From  RH01_PESSOA AS RH01  ")
        strSQL.Append("  Left Join  RH02_SERVIDOR AS RH02 ON RH02.RH01_ID_PESSOA = RH01.RH01_ID_PESSOA ")
        strSQL.Append("  Left Join  rh05_tipo_vinculo AS rh05 ON RH02.rh05_id_tipo_vinculo = rh05.rh05_id_tipo_vinculo ")
        strSQL.Append("  Left Join  RH67_DOCUMENTO_PESSOA AS RH67 ON RH67.RH01_ID_PESSOA = RH01.RH01_ID_PESSOA ")
	    strSQL.Append("  Left Join  RH16_CARGO as RH16 on RH16.RH16_ID_CARGO = RH02.RH16_ID_CARGO  ")
	    strSQL.Append("  Left Join  RH78_SERVIDOR_CARGA_HORARIA as RH78 ON RH78.RH02_ID_SERVIDOR = RH02.RH02_ID_SERVIDOR ")
	    strSQL.Append("  Left Join  RH77_CARGA_HORARIA  as RH77 on RH77.RH77_ID_CARGA_HORARIA = RH78.RH77_ID_CARGA_HORARIA ")
        strSQL.Append("  Left Join  RH79_TIPO_CARGA_HORARIA as RH79 on RH79.RH79_ID_TIPO_CARGA_HORARIA = RH77.RH79_ID_TIPO_CARGA_HORARIA ")
        strSQL.Append("   left join RH07_SITUACAO_SERVIDOR rh07 on rh02.RH07_ID_SITUACAO_SERVIDOR = rh07.RH07_ID_SITUACAO_SERVIDOR ")
        strSQL.Append("  where RH01.RH01_ID_PESSOA = " & Codigo)
        strSQL.Append("	   And rh07.RH07_ID_SITUACAO_SERVIDOR in (1,10,11) ")

        Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select RH01_ID_PESSOA as CODIGO, RH01_NU_CPF as DESCRICAO")
        strSQL.Append(" from RH01_PESSOA")
        strSQL.Append(" order by 2 ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return dt
    End Function

    Public Function PesquisarAlocacaoCargaHoraria(ByVal Pessoa As Integer) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append("  Select  RH80.RH80_ID_ALOCACAO_CARGA_HORARIA, isnull(RH48.RH48_NM_TIPO_ESCOLA,'') + ' ' + RH36.RH36_NM_LOTACAO as RH36_NM_LOTACAO, convert(varchar,RH80_QT_HORA_ALOCADA) + ' HORAS DE ' + RH79_NM_TIPO_CARGA_HORARIA  as CARGA_HORARIA " & Chr(13))
        strSQL.Append("  , RH02.RH02_CD_MATRICULA, TG06.TG06_NM_TURNO " & Chr(13))
        strSQL.Append(" From  RH80_ALOCACAO_CARGA_HORARIA As RH80 " & Chr(13))
        strSQL.Append(" Left Join RH78_SERVIDOR_CARGA_HORARIA as RH78 on RH78.RH78_ID_SERVIDOR_CARGA_HORARIA = RH80.RH78_ID_SERVIDOR_CARGA_HORARIA " & Chr(13))
        strSQL.Append(" Left Join RH77_CARGA_HORARIA as RH77 on RH77.RH77_ID_CARGA_HORARIA = RH78.RH77_ID_CARGA_HORARIA " & Chr(13))
        strSQL.Append(" Left Join RH79_TIPO_CARGA_HORARIA as RH79 on RH79.RH79_ID_TIPO_CARGA_HORARIA = RH77.RH79_ID_TIPO_CARGA_HORARIA " & Chr(13))
        strSQL.Append(" Left Join RH14_LOTACAO_SERVIDOR As RH14 ON RH14.RH14_ID_LOTACAO_SERVIDOR = RH80.RH14_ID_LOTACAO_SERVIDOR " & Chr(13))
        strSQL.Append(" Left Join RH02_SERVIDOR AS RH02 ON RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR  " & Chr(13))
        strSQL.Append(" Left Join RH36_LOTACAO  AS RH36 ON RH36.RH36_ID_LOTACAO =  RH14.RH36_ID_LOTACAO " & Chr(13))
        strSQL.Append(" Left Join RH48_TIPO_ESCOLA as RH48 on RH48.RH48_ID_TIPO_ESCOLA = RH36.RH48_ID_TIPO_ESCOLA " & Chr(13))
        strSQL.Append(" Left Join DBGERAL.DBO.TG06_TURNO AS TG06 ON TG06.TG06_ID_TURNO = RH80.TG06_ID_TURNO " & Chr(13))
        strSQL.Append(" where RH80.RH80_ID_ALOCACAO_CARGA_HORARIA Is Not null " & Chr(13))
        strSQL.Append(" And RH80.RH80_DH_DESATIVACAO Is null " & Chr(13))
        strSQL.Append(" And RH02.RH01_ID_PESSOA =  " & Pessoa & Chr(13))
        strSQL.Append(" And (RH78_DT_TERMINO_VIGENCIA Is null  " & Chr(13))
        strSQL.Append(" Or RH78_DT_TERMINO_VIGENCIA > getdate()) " & Chr(13))
        strSQL.Append(" And RH07_ID_SITUACAO_SERVIDOR In (1, 10, 11)  " & Chr(13))

        strSQL.Append(" ORDER BY isnull(RH48.RH48_NM_TIPO_ESCOLA,'') + ' ' + RH36.RH36_NM_LOTACAO , TG06.TG06_NM_TURNO, RH79_NM_TIPO_CARGA_HORARIA")
        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function


    Public Function ObterUltimo() As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(RH01_ID_PESSOA) from RH01_PESSOA")

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

    Public Function Excluir(ByVal PessoaId As String) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from RH01_PESSOA")
        strSQL.Append(" where RH01_ID_PESSOA = " & PessoaId)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

    Public Function ObterLotacaoServidor(ByVal CodigoPessoa As Integer, Optional Ano As Integer = 2022) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoLotacao As Integer


        If CodigoPessoa > 0 Then

            strSQL.Append(" SELECT TOP 1 RH36.RH36_ID_LOTACAO" & vbCrLf)
            strSQL.Append("   FROM RH02_SERVIDOR         RH02 " & vbCrLf)
            strSQL.Append("   JOIN RH01_PESSOA           RH01 ON RH01.RH01_ID_PESSOA   = RH02.RH01_ID_PESSOA" & vbCrLf)
            strSQL.Append("   JOIN RH14_LOTACAO_SERVIDOR RH14 ON RH14.RH02_ID_SERVIDOR = RH02.RH02_ID_SERVIDOR" & vbCrLf)
            strSQL.Append("   JOIN RH36_LOTACAO          RH36 ON RH36.RH36_ID_LOTACAO  = RH14.RH36_ID_LOTACAO" & vbCrLf)
            strSQL.Append("   JOIN RH88_PERIODO          RH88 ON RH88.RH88_ID_PERIODO  = RH14.RH88_ID_PERIODO" & vbCrLf)
            strSQL.Append("  WHERE RH36.RH36_ID_LOTACAO IS NOT NULL" & vbCrLf)
            strSQL.Append("    AND RH14.RH14_DT_DESLIGAMENTO IS NULL" & vbCrLf)
            strSQL.Append("    AND LEFT(RH88.RH88_NM_PERIODO,4) = " & Ano & vbCrLf) 'YEAR(GETDATE())
            strSQL.Append("    AND RH02.RH01_ID_PESSOA = " & CodigoPessoa & vbCrLf)

            With cnn.AbrirDataTable(strSQL.ToString)
                If Not IsDBNull(.Rows(0)(0)) Then
                    CodigoLotacao = .Rows(0)(0)
                Else
                    CodigoLotacao = 0
                End If
            End With

            cnn.FecharBanco()
            cnn = Nothing

        Else
            CodigoLotacao = Nothing
        End If

        Return CodigoLotacao

    End Function

    Public Function LotacaoServidorValidaProtocolo(ByVal CodigoPessoa As Integer, Optional Ano As Integer = 2022) As Boolean
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim retorno As Boolean

        strSQL.Append(" SELECT TOP 1 RH02.RH02_CD_MATRICULA, RH02.RH02_ID_SERVIDOR, RH01.RH01_NM_PESSOA, RH01.CA04_ID_USUARIO " & vbCrLf)
        strSQL.Append("   FROM RH02_SERVIDOR         RH02 " & vbCrLf)
        strSQL.Append("   JOIN RH01_PESSOA           RH01 ON RH01.RH01_ID_PESSOA   = RH02.RH01_ID_PESSOA" & vbCrLf)
        strSQL.Append("   JOIN RH14_LOTACAO_SERVIDOR RH14 ON RH14.RH02_ID_SERVIDOR = RH02.RH02_ID_SERVIDOR" & vbCrLf)
        strSQL.Append("   JOIN RH36_LOTACAO          RH36 ON RH36.RH36_ID_LOTACAO  = RH14.RH36_ID_LOTACAO" & vbCrLf)
        strSQL.Append("   JOIN RH88_PERIODO          RH88 ON RH88.RH88_ID_PERIODO  = RH14.RH88_ID_PERIODO" & vbCrLf)
        strSQL.Append("  WHERE RH36.RH36_ID_LOTACAO IS NOT NULL" & vbCrLf)
        strSQL.Append("    AND RH14.RH14_DT_DESLIGAMENTO IS NULL" & vbCrLf)
        strSQL.Append("    AND LEFT(RH88.RH88_NM_PERIODO,4) = " & Ano & vbCrLf) 'YEAR(GETDATE())
        strSQL.Append("    AND RH02.RH01_ID_PESSOA = " & CodigoPessoa & vbCrLf)
        strSQL.Append("    AND RH36.RH36_ID_LOTACAO = 950" & vbCrLf) 'SUPERVISÃO DE PROTOCOLO E ARQUIVO

        Try

            With cnn.AbrirDataTable(strSQL.ToString)
                If .Rows.Count > 0 Then
                    retorno = True
                Else
                    retorno = False
                End If
            End With

        Catch ex As Exception
            Dim erro As String = ex.ToString
            retorno = False
        End Try

        cnn = Nothing
        Return retorno
    End Function

    Public Function LotacaoServidorValidaJuridico(ByVal CodigoPessoa As Integer, Optional Ano As Integer = 2022) As Boolean
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim retorno As Boolean

        strSQL.Append(" SELECT TOP 1 RH02.RH02_CD_MATRICULA, RH02.RH02_ID_SERVIDOR, RH01.RH01_NM_PESSOA, RH01.CA04_ID_USUARIO " & vbCrLf)
        strSQL.Append("   FROM RH02_SERVIDOR         RH02 " & vbCrLf)
        strSQL.Append("   JOIN RH01_PESSOA           RH01 ON RH01.RH01_ID_PESSOA   = RH02.RH01_ID_PESSOA" & vbCrLf)
        strSQL.Append("   JOIN RH14_LOTACAO_SERVIDOR RH14 ON RH14.RH02_ID_SERVIDOR = RH02.RH02_ID_SERVIDOR" & vbCrLf)
        strSQL.Append("   JOIN RH36_LOTACAO          RH36 ON RH36.RH36_ID_LOTACAO  = RH14.RH36_ID_LOTACAO" & vbCrLf)
        strSQL.Append("   JOIN RH88_PERIODO          RH88 ON RH88.RH88_ID_PERIODO  = RH14.RH88_ID_PERIODO" & vbCrLf)
        strSQL.Append("  WHERE RH36.RH36_ID_LOTACAO IS NOT NULL" & vbCrLf)
        strSQL.Append("    AND RH14.RH14_DT_DESLIGAMENTO IS NULL" & vbCrLf)
        strSQL.Append("    AND LEFT(RH88.RH88_NM_PERIODO,4) = " & Ano & vbCrLf) 'YEAR(GETDATE())
        strSQL.Append("    AND RH02.RH01_ID_PESSOA = " & CodigoPessoa & vbCrLf)
        strSQL.Append("    AND RH36.RH36_ID_LOTACAO = 937" & vbCrLf) 'ASSESSORIA JURIDICA

        Try

            With cnn.AbrirDataTable(strSQL.ToString)
                If .Rows.Count > 0 Then
                    retorno = True
                Else
                    retorno = False
                End If
            End With

        Catch ex As Exception
            Dim erro As String = ex.ToString
            retorno = False
        End Try

        cnn = Nothing
        Return retorno
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

