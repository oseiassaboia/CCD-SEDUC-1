Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Runtime.InteropServices.WindowsRuntime

Public Class Lotacao
	Private RH36_ID_LOTACAO as Integer
	Private RH36_ID_LOTACAO_1 as String
	Private RH32_ID_TIPO_LOTACAO as String
	Private RH48_ID_TIPO_ESCOLA as String
	Private RH47_ID_TIPO_UNIDADE as String
	Private RH61_ID_TIPO_DEPEND_ADM as String
	Private TG03_ID_MUNICIPIO as String
	Private RH36_DH_SIT_FUNCIONAMENTO  as String
	Private CA04_ID_USUARIO as Integer
	Private RH36_NM_LOTACAO as String
	Private RH36_SG_LOTACAO as String
	Private RH36_NU_TELEFONE as String
	Private RH36_NM_EMAIL as String
	Private RH36_CD_INEP_LOTACAO as String
	Private RH36_DH_CADASTRO as String
    Private RH60_ID_SIT_FUNCIONAMENTO As Integer
    Private RH36_NM_APELIDO_LOTACAO As String
    Private RH36_IN_UTILIZA_SIAEP As Integer
    Private RH36_IN_ESCOLA_DIGNA As Integer

    Public Property LotacaoId() as Integer
		Get
			Return RH36_ID_LOTACAO
		End Get
		Set(ByVal Value As Integer)
			RH36_ID_LOTACAO = Value
		End Set
	End Property
	Public Property LotacaoMae() as integer
		Get
			Return RH36_ID_LOTACAO_1
		End Get
		Set(ByVal Value As Integer)
			RH36_ID_LOTACAO_1 = Value
		End Set
	End Property
	Public Property TipoLotacaoId() as String
		Get
			Return RH32_ID_TIPO_LOTACAO
		End Get
		Set(ByVal Value As String)
			RH32_ID_TIPO_LOTACAO = Value
		End Set
	End Property
	Public Property TipoEscolaId() as String
		Get
			Return RH48_ID_TIPO_ESCOLA
		End Get
		Set(ByVal Value As String)
			RH48_ID_TIPO_ESCOLA = Value
		End Set
	End Property
	Public Property TipoUnidadeId() as String
		Get
			Return RH47_ID_TIPO_UNIDADE
		End Get
		Set(ByVal Value As String)
			RH47_ID_TIPO_UNIDADE = Value
		End Set
	End Property
	Public Property DependenciaAdmId() as String
		Get
			Return RH61_ID_TIPO_DEPEND_ADM
		End Get
		Set(ByVal Value As String)
			RH61_ID_TIPO_DEPEND_ADM = Value
		End Set
	End Property
	Public Property MunicipioId() as String
		Get
			Return TG03_ID_MUNICIPIO
		End Get
		Set(ByVal Value As String)
			TG03_ID_MUNICIPIO = Value
		End Set
	End Property
    Public Property SituacaoFuncionamentoId() As Integer
        Get
            Return RH60_ID_SIT_FUNCIONAMENTO
        End Get
        Set(ByVal Value As Integer)
            RH60_ID_SIT_FUNCIONAMENTO = Value
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
	Public Property DescricaoLotacao() as String
		Get
			Return RH36_NM_LOTACAO
		End Get
		Set(ByVal Value As String)
			RH36_NM_LOTACAO = Value
		End Set
	End Property
	Public Property SiglaLotacao() as String
		Get
			Return RH36_SG_LOTACAO
		End Get
		Set(ByVal Value As String)
			RH36_SG_LOTACAO = Value
		End Set
	End Property
	Public Property Contato() as String
		Get
			Return RH36_NU_TELEFONE
		End Get
		Set(ByVal Value As String)
			RH36_NU_TELEFONE = Value
		End Set
	End Property
	Public Property Email() as String
		Get
			Return RH36_NM_EMAIL
		End Get
		Set(ByVal Value As String)
			RH36_NM_EMAIL = Value
		End Set
	End Property
	Public Property InepLotacao() as String
		Get
			Return RH36_CD_INEP_LOTACAO
		End Get
		Set(ByVal Value As String)
			RH36_CD_INEP_LOTACAO = Value
		End Set
	End Property
	Public Property DataHoraCadastro() as String
		Get
			Return RH36_DH_CADASTRO
		End Get
		Set(ByVal Value As String)
			RH36_DH_CADASTRO = Value
		End Set
	End Property
	Public Property DataHoraSituacaoFuncionamento() as String
		Get
            Return RH36_DH_SIT_FUNCIONAMENTO
        End Get
		Set(ByVal Value As String)
            RH36_DH_SIT_FUNCIONAMENTO = Value
        End Set
	End Property

    Public Property Apelido() As String
        Get
            Return RH36_NM_APELIDO_LOTACAO
        End Get
        Set(value As String)
            RH36_NM_APELIDO_LOTACAO = Value
        End Set
    End Property

    Public Property UtilizaSiaep As Integer
        Get
            Return RH36_IN_UTILIZA_SIAEP
        End Get
        Set(value As Integer)
            RH36_IN_UTILIZA_SIAEP = value
        End Set
    End Property
    Public Property EscolaDigna As Integer
        Get
            Return RH36_IN_ESCOLA_DIGNA
        End Get
        Set(value As Integer)
            RH36_IN_ESCOLA_DIGNA = value
        End Set
    End Property


    Public Sub New(Optional ByVal LotacaoId as integer = 0)
		If LotacaoId > 0 Then
			Obter(LotacaoId)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH36_LOTACAO")
		strSQL.Append(" where RH36_ID_LOTACAO = " & LotacaoId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

        dr("RH36_ID_LOTACAO_1") = ProBanco(RH36_ID_LOTACAO_1, eTipoValor.CHAVE)
        dr("RH32_ID_TIPO_LOTACAO") = ProBanco(RH32_ID_TIPO_LOTACAO, eTipoValor.CHAVE)
        dr("RH48_ID_TIPO_ESCOLA") = ProBanco(RH48_ID_TIPO_ESCOLA, eTipoValor.CHAVE)
        dr("RH47_ID_TIPO_UNIDADE") = ProBanco(RH47_ID_TIPO_UNIDADE, eTipoValor.CHAVE)
        dr("RH61_ID_TIPO_DEPEND_ADM") = ProBanco(RH61_ID_TIPO_DEPEND_ADM, eTipoValor.CHAVE)
        dr("TG03_ID_MUNICIPIO") = ProBanco(TG03_ID_MUNICIPIO, eTipoValor.CHAVE)
        dr("RH60_ID_SIT_FUNCIONAMENTO") = ProBanco(RH60_ID_SIT_FUNCIONAMENTO, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("RH36_NM_LOTACAO") = ProBanco(RH36_NM_LOTACAO, eTipoValor.TEXTO)
		dr("RH36_SG_LOTACAO") = ProBanco(RH36_SG_LOTACAO, eTipoValor.TEXTO)
		dr("RH36_NU_TELEFONE") = ProBanco(RH36_NU_TELEFONE, eTipoValor.TEXTO)
		dr("RH36_NM_EMAIL") = ProBanco(RH36_NM_EMAIL, eTipoValor.TEXTO)
		dr("RH36_CD_INEP_LOTACAO") = ProBanco(RH36_CD_INEP_LOTACAO, eTipoValor.TEXTO)
		dr("RH36_DH_CADASTRO") = ProBanco(RH36_DH_CADASTRO, eTipoValor.DATA)
        dr("RH36_DH_SIT_FUNCIONAMENTO") = ProBanco(RH36_DH_SIT_FUNCIONAMENTO, eTipoValor.DATA)
        dr("RH36_NM_APELIDO_LOTACAO") = ProBanco(RH36_NM_APELIDO_LOTACAO, eTipoValor.TEXTO)
        dr("RH36_IN_UTILIZA_SIAEP") = ProBanco(RH36_IN_UTILIZA_SIAEP, eTipoValor.BOOLEANO)
        dr("RH36_IN_ESCOLA_DIGNA") = ProBanco(RH36_IN_ESCOLA_DIGNA, eTipoValor.BOOLEANO)

        cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal LotacaoId as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH36_LOTACAO")
		strSQL.Append(" where RH36_ID_LOTACAO = " & LotacaoId)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH36_ID_LOTACAO = DoBanco(dr("RH36_ID_LOTACAO"), eTipoValor.CHAVE)
            RH36_ID_LOTACAO_1 = DoBanco(dr("RH36_ID_LOTACAO_1"), eTipoValor.CHAVE)
            RH32_ID_TIPO_LOTACAO = DoBanco(dr("RH32_ID_TIPO_LOTACAO"), eTipoValor.CHAVE)
            RH48_ID_TIPO_ESCOLA = DoBanco(dr("RH48_ID_TIPO_ESCOLA"), eTipoValor.CHAVE)
            RH47_ID_TIPO_UNIDADE = DoBanco(dr("RH47_ID_TIPO_UNIDADE"), eTipoValor.CHAVE)
            RH61_ID_TIPO_DEPEND_ADM = DoBanco(dr("RH61_ID_TIPO_DEPEND_ADM"), eTipoValor.CHAVE)
            TG03_ID_MUNICIPIO = DoBanco(dr("TG03_ID_MUNICIPIO"), eTipoValor.CHAVE)
            RH60_ID_SIT_FUNCIONAMENTO = DoBanco(dr("RH60_ID_SIT_FUNCIONAMENTO"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
			RH36_NM_LOTACAO = DoBanco(dr("RH36_NM_LOTACAO"), eTipoValor.TEXTO)
			RH36_SG_LOTACAO = DoBanco(dr("RH36_SG_LOTACAO"), eTipoValor.TEXTO)
			RH36_NU_TELEFONE = DoBanco(dr("RH36_NU_TELEFONE"), eTipoValor.TEXTO)
			RH36_NM_EMAIL = DoBanco(dr("RH36_NM_EMAIL"), eTipoValor.TEXTO)
			RH36_CD_INEP_LOTACAO = DoBanco(dr("RH36_CD_INEP_LOTACAO"), eTipoValor.TEXTO)
			RH36_DH_CADASTRO = DoBanco(dr("RH36_DH_CADASTRO"), eTipoValor.DATA)
            RH36_DH_SIT_FUNCIONAMENTO = DoBanco(dr("RH36_DH_SIT_FUNCIONAMENTO"), eTipoValor.DATA)
            RH36_NM_APELIDO_LOTACAO = DoBanco(dr("RH36_NM_APELIDO_LOTACAO"),eTipoValor.TEXTO)
            RH36_IN_UTILIZA_SIAEP = DoBanco(dr("RH36_IN_UTILIZA_SIAEP"), eTipoValor.BOOLEANO)
            RH36_IN_ESCOLA_DIGNA = DoBanco(dr("RH36_IN_ESCOLA_DIGNA"), eTipoValor.BOOLEANO)
        End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional LotacaoId As Integer = 0, Optional LotacaoMae As String = "", Optional TipoLotacaoId As Integer = 0, Optional TipoEscolaId As Integer = 0, Optional TipoUnidadeId As Integer = 0, Optional DependenciaAdmId As Integer = 0, Optional MunicipioId As Integer = 0, Optional SituacaoFuncionamentoId As Integer = 0, Optional UsuarioId As Integer = 0, Optional DescricaoLotacao As String = "", Optional SiglaLotacao As String = "", Optional Contato As String = "", Optional Email As String = "", Optional InepLotacao As String = "", Optional DataHoraCadastro As String = "", Optional DataHoraSituacaoFuncionamento As String = "", Optional ByVal Municipio As Integer = 0, Optional Regional As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select *,RH36_ID_LOTACAO as CODIGO,RH36_NM_LOTACAO as DESCRICAO,isnull(RH48.RH48_NM_TIPO_ESCOLA,'') + ' ' + lotacao.RH36_NM_LOTACAO as LOTACAO ")
        strSQL.Append(" from RH36_LOTACAO lotacao")
        strSQL.Append(" inner join DBGERAL..TG03_MUNICIPIO Municipio on Municipio.TG03_ID_MUNICIPIO  = lotacao.TG03_ID_MUNICIPIO ")
        strSQL.Append(" inner join DBGERAL..TG05_REGIONAL TG05 ON   TG05.TG05_ID_REGIONAL  = Municipio.TG05_ID_REGIONAL ")
        strSQL.Append(" Left Join  RH48_TIPO_ESCOLA as RH48 on RH48.RH48_ID_TIPO_ESCOLA = lotacao.RH48_ID_TIPO_ESCOLA ")
        strSQL.Append(" inner Join  RH61_TIPO_DEPEND_ADM as RH61 on RH61.RH61_ID_TIPO_DEPEND_ADM = lotacao.RH61_ID_TIPO_DEPEND_ADM ")
        strSQL.Append(" inner Join  RH60_SIT_FUNCIONAMENTO as RH60 on RH60.RH60_ID_SIT_FUNCIONAMENTO = lotacao.RH60_ID_SIT_FUNCIONAMENTO ")
        strSQL.Append(" where RH36_ID_LOTACAO is not null")

        If LotacaoId > 0 Then
            strSQL.Append(" and RH36_ID_LOTACAO = " & LotacaoId)
        End If

        If LotacaoMae.ToUpper() <> "" Then
            strSQL.Append(" and RH36_ID_LOTACAO_1 = '%" & LotacaoMae & "%'")
        End If

        If TipoLotacaoId > 0 Then
            strSQL.Append(" and RH32_ID_TIPO_LOTACAO = " & TipoLotacaoId)
        End If

        If TipoEscolaId > 0 Then
            strSQL.Append(" and RH48.RH48_ID_TIPO_ESCOLA = " & TipoEscolaId)
        End If

        If TipoUnidadeId > 0 Then
            strSQL.Append(" and RH47_ID_TIPO_UNIDADE = " & TipoUnidadeId)
        End If

        If DependenciaAdmId > 0 Then
            strSQL.Append(" and RH61.RH61_ID_TIPO_DEPEND_ADM = " & DependenciaAdmId)
        End If

        If MunicipioId > 0 Then
            strSQL.Append(" and Municipio.TG03_ID_MUNICIPIO = " & MunicipioId)
        End If

        If SituacaoFuncionamentoId > 0 Then
            strSQL.Append(" and RH60.RH60_ID_SIT_FUNCIONAMENTO  = " & SituacaoFuncionamentoId)
        End If

        If UsuarioId > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & UsuarioId)
        End If

        If DescricaoLotacao.ToUpper() <> "" Then
            strSQL.Append(" and upper(RH36_NM_LOTACAO) like '%" & DescricaoLotacao.ToUpper & "%'")
        End If

        If SiglaLotacao.ToUpper() <> "" Then
            strSQL.Append(" and upper(RH36_SG_LOTACAO) like '%" & SiglaLotacao.ToUpper & "%'")
        End If

        If Contato.ToUpper() <> "" Then
            strSQL.Append(" and upper(RH36_NU_TELEFONE) like '%" & Contato.ToUpper & "%'")
        End If

        If Email.ToUpper() <> "" Then
            strSQL.Append(" and upper(RH36_NM_EMAIL) like '%" & Email.ToUpper & "%'")
        End If

        If InepLotacao <> "" Then
            strSQL.Append(" and upper(RH36_CD_INEP_LOTACAO) like '%" & InepLotacao.ToUpper & "%'")
        End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" and RH36_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
        End If

        If IsDate(DataHoraSituacaoFuncionamento) Then
            strSQL.Append(" and RH60_DH_SIT_FUNCIONAMENTO = Convert(DateTime, '" & DataHoraSituacaoFuncionamento & "', 103)")
        End If

        If Municipio > 0 Then
            strSQL.Append(" and Municipio.TG03_ID_MUNICIPIO = " & Municipio)
        End If

        If Regional > 0 Then
            strSQL.Append(" and TG05.TG05_ID_REGIONAL = " & Regional)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH36_NM_LOTACAO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function
	Public Function PesquisarMapeamento(Optional ByVal Sort As String = "", Optional LotacaoId As Integer = 0, Optional LotacaoMae As String = "", Optional TipoLotacaoId As Integer = 0, Optional TipoEscolaId As Integer = 0, Optional TipoUnidadeId As Integer = 0, Optional DependenciaAdmId As Integer = 0, Optional MunicipioId As Integer = 0, Optional SituacaoFuncionamentoId As Integer = 0, Optional UsuarioId As Integer = 0, Optional DescricaoLotacao As String = "", Optional SiglaLotacao As String = "", Optional Contato As String = "", Optional Email As String = "", Optional InepLotacao As String = "", Optional DataHoraCadastro As String = "", Optional DataHoraSituacaoFuncionamento As String = "", Optional ByVal Municipio As Integer = 0, Optional Regional As Integer = 0 _
										, Optional ByVal PessoaId As Integer = 0, Optional MapeamentoEncerrado As Integer = 0, Optional ByVal TipoEscola As String = "") As DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append(" select distinct lotacao.*,lotacao.RH36_ID_LOTACAO as CODIGO,RH36_NM_LOTACAO as DESCRICAO,isnull(RH48.RH48_NM_TIPO_ESCOLA,'') + ' ' + lotacao.RH36_NM_LOTACAO as LOTACAO,")
		'strSQL.Append("  isnull(RH89.RH89_ID_PERIODO_MAPEAMENTO,0) as ENCERRADO ,    ")
		strSQL.Append(" isnull((select top 1 rh89_1.RH89_ID_PERIODO_MAPEAMENTO ")
		strSQL.Append(" 	from RH89_PERIODO_MAPEAMENTO rh89_1 ")
		strSQL.Append(" 	where rh89_1.RH36_ID_LOTACAO = lotacao.RH36_ID_LOTACAO ")
		strSQL.Append(" 	and year(rh89.RH89_DH_CADASTRO)  = year(getdate()) ),0)  as ENCERRADO, ")
		strSQL.Append("(select count(RH02_ID_SERVIDOR) from RH14_LOTACAO_SERVIDOR rh14_1 ")
		strSQL.Append(" 		inner join RH88_PERIODO rh88_1 On rh14_1.RH88_ID_PERIODO = rh88_1.RH88_ID_PERIODO ")
		strSQL.Append(" 	where rh14_1.RH36_ID_LOTACAO = lotacao.RH36_ID_LOTACAO And rh14_1.RH14_DT_DESLIGAMENTO Is null And rh88_1.RH88_NM_PERIODO = year(getdate()))As Qtd_Servidor, ")
		strSQL.Append(" (Select count(RH02_ID_SERVIDOR) from RH14_LOTACAO_SERVIDOR rh14_2 ")
		strSQL.Append(" 		inner join RH88_PERIODO rh88_2 On rh14_2.RH88_ID_PERIODO = rh88_2.RH88_ID_PERIODO ")
		strSQL.Append(" 		inner join RH53_FUNCAO_LOTADO rh53_2 On rh14_2.RH14_ID_LOTACAO_SERVIDOR = rh53_2.RH14_ID_LOTACAO_SERVIDOR ")
		strSQL.Append(" 	where rh14_2.RH36_ID_LOTACAO = lotacao.RH36_ID_LOTACAO And rh14_2.RH14_DT_DESLIGAMENTO Is null And rh88_2.RH88_NM_PERIODO = year(getdate()))As Qtd_Servidor_Funcao, ")
		strSQL.Append(" (Select count(RH02_ID_SERVIDOR) from RH14_LOTACAO_SERVIDOR rh14_3 ")
		strSQL.Append(" 		inner join RH88_PERIODO rh88_3 On rh14_3.RH88_ID_PERIODO = rh88_3.RH88_ID_PERIODO ")
		strSQL.Append(" 		inner join RH53_FUNCAO_LOTADO rh53_3 On rh14_3.RH14_ID_LOTACAO_SERVIDOR = rh53_3.RH14_ID_LOTACAO_SERVIDOR ")
		strSQL.Append(" 	where rh14_3.RH36_ID_LOTACAO = lotacao.RH36_ID_LOTACAO And rh53_3.RH52_ID_FUNCAO_MAPEAMENTO=18 And rh14_3.RH14_DT_DESLIGAMENTO Is null And rh88_3.RH88_NM_PERIODO = year(getdate()))As Qtd_Servidor_Professor, ")
		strSQL.Append(" (Select count(RH02_ID_SERVIDOR) from RH14_LOTACAO_SERVIDOR rh14_4 ")
		strSQL.Append(" 		inner join RH88_PERIODO rh88_4 On rh14_4.RH88_ID_PERIODO = rh88_4.RH88_ID_PERIODO ")
		strSQL.Append(" 		inner join RH80_ALOCACAO_CARGA_HORARIA rh80_4 On rh14_4.RH14_ID_LOTACAO_SERVIDOR = rh80_4.RH14_ID_LOTACAO_SERVIDOR ")
		strSQL.Append(" 		inner join RH74_HABILITACAO rh74_4 On rh80_4.RH80_ID_ALOCACAO_CARGA_HORARIA = rh74_4.RH80_ID_ALOCACAO_CARGA_HORARIA ")
		strSQL.Append(" 	where rh14_4.RH36_ID_LOTACAO = lotacao.RH36_ID_LOTACAO And rh14_4.RH14_DT_DESLIGAMENTO Is null And rh88_4.RH88_NM_PERIODO = year(getdate()) ")
		strSQL.Append(" 		And rh80_4.RH80_DH_DESATIVACAO Is null And rh74_4.RH74_DH_DESATIVACAO Is null )As Qtd_CH_Distribuida ")
		strSQL.Append(" from RH36_LOTACAO lotacao")
		strSQL.Append(" inner join DBGERAL..TG03_MUNICIPIO Municipio On Municipio.TG03_ID_MUNICIPIO  = lotacao.TG03_ID_MUNICIPIO ")
		strSQL.Append(" inner join DBGERAL..TG05_REGIONAL TG05 On   TG05.TG05_ID_REGIONAL  = Municipio.TG05_ID_REGIONAL ")
		strSQL.Append(" Left Join  RH48_TIPO_ESCOLA As RH48 On RH48.RH48_ID_TIPO_ESCOLA = lotacao.RH48_ID_TIPO_ESCOLA ")
		strSQL.Append(" inner Join  RH61_TIPO_DEPEND_ADM As RH61 On RH61.RH61_ID_TIPO_DEPEND_ADM = lotacao.RH61_ID_TIPO_DEPEND_ADM ")
		strSQL.Append(" inner Join  RH60_SIT_FUNCIONAMENTO As RH60 On RH60.RH60_ID_SIT_FUNCIONAMENTO = lotacao.RH60_ID_SIT_FUNCIONAMENTO ")
		strSQL.Append(" left join RH14_LOTACAO_SERVIDOR rh14 On lotacao.RH36_ID_LOTACAO = rh14.RH36_ID_LOTACAO  And RH14_DT_DESLIGAMENTO Is null ")
		strSQL.Append("  left join RH88_PERIODO rh88 On rh14.RH88_ID_PERIODO = rh88.RH88_ID_PERIODO  ")
		strSQL.Append(" left join RH02_SERVIDOR rh02 On rh02.RH02_ID_SERVIDOR = rh14.RH02_ID_SERVIDOR ")
		strSQL.Append(" left join RH89_PERIODO_MAPEAMENTO rh89 On rh89.rh36_id_lotacao = lotacao.rh36_id_lotacao And year(rh89.RH89_DH_CADASTRO)  = year(getdate()) ")
		strSQL.Append(" where lotacao.RH36_ID_LOTACAO Is Not null")
		strSQL.Append(" And RH32_ID_TIPO_LOTACAO in ( 2  , 4) ")
		strSQL.Append("  And rh88.RH88_NM_PERIODO = year(getdate()) ")
		strSQL.Append(" And RH47_ID_TIPO_UNIDADE In (1,2,3)  ")
		strSQL.Append(" And RH61.RH61_ID_TIPO_DEPEND_ADM = 1  ")
		strSQL.Append(" And RH60.RH60_ID_SIT_FUNCIONAMENTO  = 1 ")


		If TipoEscola <> "" Then
			strSQL.Append(" And RH48.RH48_ID_TIPO_ESCOLA In (" & TipoEscola & ") ")

		End If


		If MapeamentoEncerrado > 0 Then
			strSQL.Append(" And rh89.RH89_ID_PERIODO_MAPEAMENTO Is Not null ")
		End If

		If PessoaId > 0 Then
			strSQL.Append(" And rh02.RH01_ID_PESSOA = " & PessoaId)
		End If

		If LotacaoId > 0 Then
			strSQL.Append(" And RH36_ID_LOTACAO = " & LotacaoId)
		End If

		If LotacaoMae.ToUpper() <> "" Then
			strSQL.Append(" And RH36_ID_LOTACAO_1 = '%" & LotacaoMae & "%'")
		End If

		If TipoLotacaoId > 0 Then
			strSQL.Append(" and RH32_ID_TIPO_LOTACAO = " & TipoLotacaoId)
		End If

		If TipoEscolaId > 0 Then
			strSQL.Append(" and RH48.RH48_ID_TIPO_ESCOLA = " & TipoEscolaId)
		End If

		If TipoUnidadeId > 0 Then
			strSQL.Append(" and RH47_ID_TIPO_UNIDADE = " & TipoUnidadeId)
		End If

		If DependenciaAdmId > 0 Then
			strSQL.Append(" and RH61.RH61_ID_TIPO_DEPEND_ADM = " & DependenciaAdmId)
		End If

		If MunicipioId > 0 Then
			strSQL.Append(" and Municipio.TG03_ID_MUNICIPIO = " & MunicipioId)
		End If

		If SituacaoFuncionamentoId > 0 Then
			strSQL.Append(" and RH60.RH60_ID_SIT_FUNCIONAMENTO  = " & SituacaoFuncionamentoId)
		End If

		If UsuarioId > 0 Then
			strSQL.Append(" and CA04_ID_USUARIO = " & UsuarioId)
		End If

		If DescricaoLotacao.ToUpper() <> "" Then
			strSQL.Append(" and upper(RH36_NM_LOTACAO) like '%" & DescricaoLotacao.ToUpper & "%'")
		End If

		If SiglaLotacao.ToUpper() <> "" Then
			strSQL.Append(" and upper(RH36_SG_LOTACAO) like '%" & SiglaLotacao.ToUpper & "%'")
		End If

		If Contato.ToUpper() <> "" Then
			strSQL.Append(" and upper(RH36_NU_TELEFONE) like '%" & Contato.ToUpper & "%'")
		End If

		If Email.ToUpper() <> "" Then
			strSQL.Append(" and upper(RH36_NM_EMAIL) like '%" & Email.ToUpper & "%'")
		End If

		If InepLotacao <> "" Then
			strSQL.Append(" and upper(RH36_CD_INEP_LOTACAO) like '%" & InepLotacao.ToUpper & "%'")
		End If

		If IsDate(DataHoraCadastro) Then
			strSQL.Append(" and RH36_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
		End If

		If IsDate(DataHoraSituacaoFuncionamento) Then
			strSQL.Append(" and RH60_DH_SIT_FUNCIONAMENTO = Convert(DateTime, '" & DataHoraSituacaoFuncionamento & "', 103)")
		End If

		If Municipio > 0 Then
			strSQL.Append(" and Municipio.TG03_ID_MUNICIPIO = " & Municipio)
		End If

		If Regional > 0 Then
			strSQL.Append(" and TG05.TG05_ID_REGIONAL = " & Regional)
		End If

		strSQL.Append(" Order By " & IIf(Sort = "", "RH36_NM_LOTACAO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function
	Public Function ObterTabela(Optional Municipio As Integer = 0) as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH36_ID_LOTACAO as CODIGO, isnull(RH48.RH48_NM_TIPO_ESCOLA,'') + ' ' + lotacao.RH36_NM_LOTACAO as DESCRICAO")
		strSQL.Append(" from RH36_LOTACAO Lotacao")
        strSQL.Append(" Left Join  RH48_TIPO_ESCOLA as RH48 on RH48.RH48_ID_TIPO_ESCOLA = lotacao.RH48_ID_TIPO_ESCOLA ")
	    strSQL.Append(" WHERE RH36_ID_LOTACAO Is Not null ")

	    If Municipio > 0 Then
	        strSQL.Append(" And TG03_ID_MUNICIPIO = " & Municipio)
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
		
		strSQL.Append(" select max(RH36_ID_LOTACAO) from RH36_LOTACAO")

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
	Public Function Excluir(ByVal LotacaoId as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH36_LOTACAO")
		strSQL.Append(" where RH36_ID_LOTACAO = " & LotacaoId)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

    Public Function ObterLotacoesPessoa(Optional Pessoa As Integer = 0, Optional Ano As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select RH36_ID_LOTACAO As CODIGO, TG05_NM_REGIONAL + ' - ' + TG03_NM_MUNICIPIO + ' - [' + convert(varchar,isnull(RH36_CD_INEP_LOTACAO,'')) + '] ' + isnull(RH48.RH48_NM_TIPO_ESCOLA,'') + ' ' + RH36_NM_LOTACAO as DESCRICAO " & Chr(13))
        strSQL.Append(" From RH36_LOTACAO As RH36 " & Chr(13))
        strSQL.Append(" Left Join RH48_TIPO_ESCOLA as RH48 on RH48.RH48_ID_TIPO_ESCOLA = RH36.RH48_ID_TIPO_ESCOLA " & Chr(13))
        strSQL.Append(" left join DBGERAL.DBO.TG03_MUNICIPIO AS TG03 on TG03.TG03_ID_MUNICIPIO = RH36.TG03_ID_MUNICIPIO " & Chr(13))
        strSQL.Append(" left join DBGERAL.DBO.TG05_REGIONAL AS TG05 on TG05.TG05_ID_REGIONAL = TG03.TG05_ID_REGIONAL " & Chr(13))
        strSQL.Append(" WHERE RH36_ID_LOTACAO Is Not null " & Chr(13))
        strSQL.Append(" AND RH32_ID_TIPO_LOTACAO = 2 " & Chr(13))
        strSQL.Append(" And RH36.RH60_ID_SIT_FUNCIONAMENTO = 1 " & Chr(13))
        'strSQL.Append(" And RH61_ID_TIPO_DEPEND_ADM = 1 " & Chr(13))
        strSQL.Append(" And RH36_IN_UTILIZA_SIAEP = 1 " & Chr(13))

        If Pessoa > 0 Then
            strSQL.Append(" and RH36_ID_LOTACAO in " & Chr(13))
            strSQL.Append(" ( " & Chr(13))
            strSQL.Append("     Select RH36_ID_LOTACAO " & Chr(13))
            strSQL.Append("     From RH14_LOTACAO_SERVIDOR As RH14 " & Chr(13))
            strSQL.Append("     Left Join RH02_SERVIDOR as RH02 on RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR " & Chr(13))
            strSQL.Append("     Left Join RH88_PERIODO as RH88 on RH88.RH88_ID_PERIODO = RH14.RH88_ID_PERIODO " & Chr(13))
            strSQL.Append("     where RH07_ID_SITUACAO_SERVIDOR in (1,10,11)  " & Chr(13))
            strSQL.Append("     And RH14_DT_DESLIGAMENTO Is null " & Chr(13))
            strSQL.Append("     And RH88.RH88_NM_PERIODO = '" & Ano & "'" & Chr(13))
            strSQL.Append("     And RH01_ID_PESSOA = " & Pessoa & Chr(13))

            If Ano > 0 Then
                strSQL.Append("     And RH88.RH88_NM_PERIODO = '" & Ano & "'" & Chr(13))
            Else
                strSQL.Append("     And RH88.RH88_NM_PERIODO = '" & Date.Now.Year & "'" & Chr(13))
            End If

            strSQL.Append(" ) ")
        End If

        strSQL.Append(" order by TG05_NM_REGIONAL + ' - ' + TG03_NM_MUNICIPIO + ' [' + convert(varchar,RH36_CD_INEP_LOTACAO) + '] ' + RH48_NM_TIPO_ESCOLA + ' ' + RH36_NM_LOTACAO " & Chr(13))

        dt = cnn.AbrirDataTable(strSQL.ToString)


        cnn = Nothing

        Return dt
    End Function

End Class
