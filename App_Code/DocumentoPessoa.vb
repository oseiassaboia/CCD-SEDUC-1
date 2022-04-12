Imports Microsoft.VisualBasic
Imports System.Data

Public Class DocumentoPessoa
    Implements IDisposable
	Private RH67_ID_DOCUMENTO_PESSOA as Integer
	Private RH01_ID_PESSOA as String
	Private TG12_ID_ESTADO_CIVIL as String
	Private TG14_ID_GRAU_INSTRUCAO as String
	Private TG03_ID_MUNICIPIO_NASCIMENTO as String
	Private TG43_ID_ORGAO_EMISSOR_RG as String
	Private RH67_SG_UF_EMISSOR_RG as String
	Private RH67_NU_RG as String
	Private RH67_NU_DIGITO_RG as String
	Private RH67_DT_EMISSAO_RG as String
	Private RH67_NU_PIS_PASEP as String
	Private RH67_NU_TITULO_ELEITOR as String
	Private RH67_CD_SECAO_ELEITOR as String
	Private RH67_CD_ZONA_ELEITOR as String
	Private RH67_NU_CERTIFICADO_RESERVISTA as String
	Private RH67_NM_REGIAO_RESERVISTA as String
	Private RH67_NU_CONSELHO as String
    Private RH67_DT_CONSELHO As String
    Private RH67_DH_ULTIMA_CARGA As String
    Private RH67_SG_EXTENSAO_ARQ_IMAGEM as String
    Private RH67_IM_PESSOA As String
    Private RH67_IM_DOC_OFICIAL_IDENTIFIC As String
    Private RH67_IM_CPF As String
    Private RH67_IM_PIS_PASEP As String
    Private RH67_IM_COMPROVANTE_RESIDENCIA As String
    Private RH67_SG_EXTENSAO_IMG_IDENTIFIC As String
    Private RH67_SG_EXTENSAO_IMG_CPF As String
    Private RH67_SG_EXTENSAO_IMG_PIS_PASEP As String
    Private RH67_SG_EXTENSAO_IMG_COMP_RES As String

    Public Property DocumentoId() as Integer
		Get
			Return RH67_ID_DOCUMENTO_PESSOA
		End Get
		Set(ByVal Value As Integer)
			RH67_ID_DOCUMENTO_PESSOA = Value
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
	Public Property EstadoCivilId() as String
		Get
			Return TG12_ID_ESTADO_CIVIL
		End Get
		Set(ByVal Value As String)
			TG12_ID_ESTADO_CIVIL = Value
		End Set
	End Property
	Public Property GrauInstrucaoId() as String
		Get
			Return TG14_ID_GRAU_INSTRUCAO
		End Get
		Set(ByVal Value As String)
			TG14_ID_GRAU_INSTRUCAO = Value
		End Set
	End Property
	Public Property MunicipioNascimentoId() as String
		Get
			Return TG03_ID_MUNICIPIO_NASCIMENTO
		End Get
		Set(ByVal Value As String)
			TG03_ID_MUNICIPIO_NASCIMENTO = Value
		End Set
	End Property
	Public Property OrgaoEmissorId() as String
		Get
			Return TG43_ID_ORGAO_EMISSOR_RG
		End Get
		Set(ByVal Value As String)
			TG43_ID_ORGAO_EMISSOR_RG = Value
		End Set
	End Property
	Public Property SiglaUfEmissorId() as String
		Get
			Return RH67_SG_UF_EMISSOR_RG
		End Get
		Set(ByVal Value As String)
			RH67_SG_UF_EMISSOR_RG = Value
		End Set
	End Property
	Public Property NumeroRg() as String
		Get
			Return RH67_NU_RG
		End Get
		Set(ByVal Value As String)
			RH67_NU_RG = Value
		End Set
	End Property
	Public Property RgDigito() as String
		Get
			Return RH67_NU_DIGITO_RG
		End Get
		Set(ByVal Value As String)
			RH67_NU_DIGITO_RG = Value
		End Set
	End Property
	Public Property RgDataEmissao() as String
		Get
			Return RH67_DT_EMISSAO_RG
		End Get
		Set(ByVal Value As String)
			RH67_DT_EMISSAO_RG = Value
		End Set
	End Property
	Public Property PisPasep() as String
		Get
			Return RH67_NU_PIS_PASEP
		End Get
		Set(ByVal Value As String)
			RH67_NU_PIS_PASEP = Value
		End Set
	End Property
	Public Property TituloEleitor() as String
		Get
			Return RH67_NU_TITULO_ELEITOR
		End Get
		Set(ByVal Value As String)
			RH67_NU_TITULO_ELEITOR = Value
		End Set
	End Property
	Public Property SecaoEleitor() as String
		Get
			Return RH67_CD_SECAO_ELEITOR
		End Get
		Set(ByVal Value As String)
			RH67_CD_SECAO_ELEITOR = Value
		End Set
	End Property
	Public Property ZonaEleitor() as String
		Get
			Return RH67_CD_ZONA_ELEITOR
		End Get
		Set(ByVal Value As String)
			RH67_CD_ZONA_ELEITOR = Value
		End Set
	End Property
	Public Property CertificadoReservista() as String
		Get
			Return RH67_NU_CERTIFICADO_RESERVISTA
		End Get
		Set(ByVal Value As String)
			RH67_NU_CERTIFICADO_RESERVISTA = Value
		End Set
	End Property
	Public Property RegiaoReservista() as String
		Get
			Return RH67_NM_REGIAO_RESERVISTA
		End Get
		Set(ByVal Value As String)
			RH67_NM_REGIAO_RESERVISTA = Value
		End Set
	End Property
	Public Property NumeroConselho() as String
		Get
			Return RH67_NU_CONSELHO
		End Get
		Set(ByVal Value As String)
			RH67_NU_CONSELHO = Value
		End Set
	End Property
	Public Property DataConselho() as String
		Get
			Return RH67_DT_CONSELHO
		End Get
		Set(ByVal Value As String)
			RH67_DT_CONSELHO = Value
		End Set
	End Property
	Public Property SiglaExtensaoArquivoImagem() as String
		Get
			Return RH67_SG_EXTENSAO_ARQ_IMAGEM
		End Get
		Set(ByVal Value As String)
			RH67_SG_EXTENSAO_ARQ_IMAGEM = Value
		End Set
	End Property
	Public Property ImagemPessoa() as String
		Get
			Return RH67_IM_PESSOA
		End Get
		Set(ByVal Value As String)
			RH67_IM_PESSOA = Value
		End Set
	End Property

    Public Property ImagemDocumentoOficialIdentificacao() As String
        Get
            Return RH67_IM_DOC_OFICIAL_IDENTIFIC
        End Get
        Set(ByVal Value As String)
            RH67_IM_DOC_OFICIAL_IDENTIFIC = Value
        End Set
    End Property

    Public Property ImagemCpf() As String
        Get
            Return RH67_IM_CPF
        End Get
        Set(ByVal Value As String)
            RH67_IM_CPF = Value
        End Set
    End Property

    Public Property ImagemPisPasep() As String
        Get
            Return RH67_IM_PIS_PASEP
        End Get
        Set(ByVal Value As String)
            RH67_IM_PIS_PASEP = Value
        End Set
    End Property

    Public Property ImagemComprovanteResidencia() As String
        Get
            Return RH67_IM_COMPROVANTE_RESIDENCIA
        End Get
        Set(ByVal Value As String)
            RH67_IM_COMPROVANTE_RESIDENCIA = Value
        End Set
    End Property


    Public Property ExtensaoImgDocumentoOficialIdentificacao() As String
        Get
            Return RH67_SG_EXTENSAO_IMG_IDENTIFIC
        End Get
        Set(ByVal Value As String)
            RH67_SG_EXTENSAO_IMG_IDENTIFIC = Value
        End Set
    End Property


    Public Property ExtensaoImgCpf() As String
        Get
            Return RH67_SG_EXTENSAO_IMG_CPF
        End Get
        Set(ByVal Value As String)
            RH67_SG_EXTENSAO_IMG_CPF = Value
        End Set
    End Property

    Public Property ExtensaoImgPisPasep() As String
        Get
            Return RH67_SG_EXTENSAO_IMG_PIS_PASEP
        End Get
        Set(ByVal Value As String)
            RH67_SG_EXTENSAO_IMG_PIS_PASEP = Value
        End Set
    End Property

    Public Property ExtensaoImgComprovanteResidencia() As String
        Get
            Return RH67_SG_EXTENSAO_IMG_COMP_RES
        End Get
        Set(ByVal Value As String)
            RH67_SG_EXTENSAO_IMG_COMP_RES = Value
        End Set
    End Property

    Public Sub New(Optional ByVal DocumentoId As Integer = 0)
        If DocumentoId > 0 Then
            Obter(DocumentoId)
        End If
    End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH67_DOCUMENTO_PESSOA")
		strSQL.Append(" where RH67_ID_DOCUMENTO_PESSOA = " & DocumentoId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

        dr("RH01_ID_PESSOA") = ProBanco(RH01_ID_PESSOA, eTipoValor.CHAVE)
        dr("TG12_ID_ESTADO_CIVIL") = ProBanco(TG12_ID_ESTADO_CIVIL, eTipoValor.CHAVE)
        dr("TG14_ID_GRAU_INSTRUCAO") = ProBanco(TG14_ID_GRAU_INSTRUCAO, eTipoValor.CHAVE)
        dr("TG03_ID_MUNICIPIO_NASCIMENTO") = ProBanco(TG03_ID_MUNICIPIO_NASCIMENTO, eTipoValor.CHAVE)
        dr("TG43_ID_ORGAO_EMISSOR_RG") = ProBanco(TG43_ID_ORGAO_EMISSOR_RG, eTipoValor.CHAVE)
        dr("RH67_SG_UF_EMISSOR_RG") = ProBanco(RH67_SG_UF_EMISSOR_RG, eTipoValor.TEXTO)
		dr("RH67_NU_RG") = ProBanco(RH67_NU_RG, eTipoValor.TEXTO)
		dr("RH67_NU_DIGITO_RG") = ProBanco(RH67_NU_DIGITO_RG, eTipoValor.TEXTO)
        dr("RH67_DT_EMISSAO_RG") = ProBanco(RH67_DT_EMISSAO_RG, eTipoValor.DATA)
        dr("RH67_NU_PIS_PASEP") = ProBanco(RH67_NU_PIS_PASEP, eTipoValor.TEXTO)
		dr("RH67_NU_TITULO_ELEITOR") = ProBanco(RH67_NU_TITULO_ELEITOR, eTipoValor.TEXTO)
		dr("RH67_CD_SECAO_ELEITOR") = ProBanco(RH67_CD_SECAO_ELEITOR, eTipoValor.TEXTO)
		dr("RH67_CD_ZONA_ELEITOR") = ProBanco(RH67_CD_ZONA_ELEITOR, eTipoValor.TEXTO)
		dr("RH67_NU_CERTIFICADO_RESERVISTA") = ProBanco(RH67_NU_CERTIFICADO_RESERVISTA, eTipoValor.TEXTO)
		dr("RH67_NM_REGIAO_RESERVISTA") = ProBanco(RH67_NM_REGIAO_RESERVISTA, eTipoValor.TEXTO)
		dr("RH67_NU_CONSELHO") = ProBanco(RH67_NU_CONSELHO, eTipoValor.TEXTO)
        dr("RH67_DT_CONSELHO") = ProBanco(RH67_DT_CONSELHO, eTipoValor.DATA)
        dr("RH67_SG_EXTENSAO_ARQ_IMAGEM") = ProBanco(RH67_SG_EXTENSAO_ARQ_IMAGEM, eTipoValor.TEXTO_LIVRE)
        dr("RH67_IM_PESSOA") = ProBanco(RH67_IM_PESSOA, eTipoValor.TEXTO_LIVRE)
        dr("RH67_IM_DOC_OFICIAL_IDENTIFIC") = ProBanco(RH67_IM_DOC_OFICIAL_IDENTIFIC, eTipoValor.TEXTO_LIVRE)
        dr("RH67_IM_CPF") = ProBanco(RH67_IM_CPF, eTipoValor.TEXTO_LIVRE)
        dr("RH67_IM_PIS_PASEP") = ProBanco(RH67_IM_PIS_PASEP, eTipoValor.TEXTO_LIVRE)
        dr("RH67_IM_COMPROVANTE_RESIDENCIA") = ProBanco(RH67_IM_COMPROVANTE_RESIDENCIA, eTipoValor.TEXTO_LIVRE)
        dr("RH67_SG_EXTENSAO_IMG_IDENTIFIC") = ProBanco(RH67_SG_EXTENSAO_IMG_IDENTIFIC, eTipoValor.TEXTO_LIVRE)
        dr("RH67_SG_EXTENSAO_IMG_CPF") = ProBanco(RH67_SG_EXTENSAO_IMG_CPF, eTipoValor.TEXTO_LIVRE)
        dr("RH67_SG_EXTENSAO_IMG_PIS_PASEP") = ProBanco(RH67_SG_EXTENSAO_IMG_PIS_PASEP, eTipoValor.TEXTO_LIVRE)
        dr("RH67_SG_EXTENSAO_IMG_COMP_RES") = ProBanco(RH67_SG_EXTENSAO_IMG_COMP_RES, eTipoValor.TEXTO_LIVRE)

        cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

    Public Sub Obter(ByVal DocumentoId As String)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH67_DOCUMENTO_PESSOA")
        strSQL.Append(" where RH01_ID_PESSOA = " & DocumentoId)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH67_ID_DOCUMENTO_PESSOA = DoBanco(dr("RH67_ID_DOCUMENTO_PESSOA"), eTipoValor.CHAVE)
            RH01_ID_PESSOA = DoBanco(dr("RH01_ID_PESSOA"), eTipoValor.CHAVE)
            TG12_ID_ESTADO_CIVIL = DoBanco(dr("TG12_ID_ESTADO_CIVIL"), eTipoValor.CHAVE)
            TG14_ID_GRAU_INSTRUCAO = DoBanco(dr("TG14_ID_GRAU_INSTRUCAO"), eTipoValor.CHAVE)
            TG03_ID_MUNICIPIO_NASCIMENTO = DoBanco(dr("TG03_ID_MUNICIPIO_NASCIMENTO"), eTipoValor.CHAVE)
            TG43_ID_ORGAO_EMISSOR_RG = DoBanco(dr("TG43_ID_ORGAO_EMISSOR_RG"), eTipoValor.CHAVE)
            RH67_SG_UF_EMISSOR_RG = DoBanco(dr("RH67_SG_UF_EMISSOR_RG"), eTipoValor.TEXTO)
            RH67_NU_RG = DoBanco(dr("RH67_NU_RG"), eTipoValor.TEXTO)
            RH67_NU_DIGITO_RG = DoBanco(dr("RH67_NU_DIGITO_RG"), eTipoValor.TEXTO)
            RH67_DT_EMISSAO_RG = DoBanco(dr("RH67_DT_EMISSAO_RG"), eTipoValor.TEXTO)
            RH67_NU_PIS_PASEP = DoBanco(dr("RH67_NU_PIS_PASEP"), eTipoValor.TEXTO)
            RH67_NU_TITULO_ELEITOR = DoBanco(dr("RH67_NU_TITULO_ELEITOR"), eTipoValor.TEXTO)
            RH67_CD_SECAO_ELEITOR = DoBanco(dr("RH67_CD_SECAO_ELEITOR"), eTipoValor.TEXTO)
            RH67_CD_ZONA_ELEITOR = DoBanco(dr("RH67_CD_ZONA_ELEITOR"), eTipoValor.TEXTO)
            RH67_NU_CERTIFICADO_RESERVISTA = DoBanco(dr("RH67_NU_CERTIFICADO_RESERVISTA"), eTipoValor.TEXTO)
            RH67_NM_REGIAO_RESERVISTA = DoBanco(dr("RH67_NM_REGIAO_RESERVISTA"), eTipoValor.TEXTO)
            RH67_NU_CONSELHO = DoBanco(dr("RH67_NU_CONSELHO"), eTipoValor.TEXTO)
            RH67_DT_CONSELHO = DoBanco(dr("RH67_DT_CONSELHO"), eTipoValor.TEXTO)
            RH67_SG_EXTENSAO_ARQ_IMAGEM = DoBanco(dr("RH67_SG_EXTENSAO_ARQ_IMAGEM"), eTipoValor.TEXTO_LIVRE)
            RH67_IM_PESSOA = DoBanco(dr("RH67_IM_PESSOA"), eTipoValor.TEXTO_LIVRE)
            RH67_IM_DOC_OFICIAL_IDENTIFIC = DoBanco(dr("RH67_IM_DOC_OFICIAL_IDENTIFIC"), eTipoValor.TEXTO_LIVRE)
            RH67_IM_CPF = DoBanco(dr("RH67_IM_CPF"), eTipoValor.TEXTO_LIVRE)
            RH67_IM_PIS_PASEP = DoBanco(dr("RH67_IM_PIS_PASEP"), eTipoValor.TEXTO_LIVRE)
            RH67_IM_COMPROVANTE_RESIDENCIA = DoBanco(dr("RH67_IM_COMPROVANTE_RESIDENCIA"), eTipoValor.TEXTO_LIVRE)
            RH67_SG_EXTENSAO_IMG_IDENTIFIC = DoBanco(dr("RH67_SG_EXTENSAO_IMG_IDENTIFIC"), eTipoValor.TEXTO_LIVRE)
            RH67_SG_EXTENSAO_IMG_CPF = DoBanco(dr("RH67_SG_EXTENSAO_IMG_CPF"), eTipoValor.TEXTO_LIVRE)
            RH67_SG_EXTENSAO_IMG_PIS_PASEP = DoBanco(dr("RH67_SG_EXTENSAO_IMG_PIS_PASEP"), eTipoValor.TEXTO_LIVRE)
            RH67_SG_EXTENSAO_IMG_COMP_RES = DoBanco(dr("RH67_SG_EXTENSAO_IMG_COMP_RES"), eTipoValor.TEXTO_LIVRE)

        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function PesquisarDocumento(ByVal PessoaId As Integer) As DataTable

        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append("select * ")
        strSQL.Append("	from( ")
        strSQL.Append("		select 1 as TipoDocumento, 'DOCUMENTO OFICIAL DE IDENTIFICACAO' as NomeDocumento, RH67_IM_DOC_OFICIAL_IDENTIFIC as DocumentoBase64, RH67_SG_EXTENSAO_IMG_IDENTIFIC as ExtensaoDocumento, RH01_ID_PESSOA, 'Imagens/Icones/' + REPLACE (RH67_SG_EXTENSAO_IMG_IDENTIFIC,'.','') + '.png' as ImageUrl 	 ")
        strSQL.Append("		from RH67_DOCUMENTO_PESSOA ")
        strSQL.Append("		where RH67_IM_DOC_OFICIAL_IDENTIFIC not like '' and RH67_SG_EXTENSAO_IMG_IDENTIFIC not like '' ")
        strSQL.Append("		union all ")
        strSQL.Append("		select 2 as TipoDocumento, 'CPF' as NomeDocumento, RH67_IM_CPF as DocumentoBase64, RH67_SG_EXTENSAO_IMG_CPF as ExtensaoDocumento, RH01_ID_PESSOA, 'Imagens/Icones/' + REPLACE (RH67_SG_EXTENSAO_IMG_CPF,'.','') + '.png' as ImageUrl 	 ")
        strSQL.Append("		from RH67_DOCUMENTO_PESSOA ")
        strSQL.Append("		where RH67_IM_CPF not like '' and RH67_SG_EXTENSAO_IMG_CPF not like '' ")
        strSQL.Append("		union all ")
        strSQL.Append("		select 3 as TipoDocumento, 'PIS/PASEP' as NomeDocumento, RH67_IM_PIS_PASEP as DocumentoBase64, RH67_SG_EXTENSAO_IMG_PIS_PASEP as ExtensaoDocumento, RH01_ID_PESSOA, 'Imagens/Icones/' + REPLACE (RH67_SG_EXTENSAO_IMG_PIS_PASEP,'.','') + '.png' as ImageUrl 	 ")
        strSQL.Append("		from RH67_DOCUMENTO_PESSOA ")
        strSQL.Append("		where RH67_IM_PIS_PASEP not like '' and RH67_SG_EXTENSAO_IMG_PIS_PASEP not like '' ")
        strSQL.Append("		union all ")
        strSQL.Append("		select 4 as TipoDocumento, 'COMPROVANTE DE RESIDENCIA' as NomeDocumento, RH67_IM_COMPROVANTE_RESIDENCIA as DocumentoBase64, RH67_SG_EXTENSAO_IMG_COMP_RES as ExtensaoDocumento, RH01_ID_PESSOA, 'Imagens/Icones/' + REPLACE (RH67_SG_EXTENSAO_IMG_COMP_RES,'.','') + '.png' as ImageUrl 	 ")
        strSQL.Append("		from RH67_DOCUMENTO_PESSOA ")
        strSQL.Append("		where RH67_IM_COMPROVANTE_RESIDENCIA not like '' and RH67_SG_EXTENSAO_IMG_COMP_RES not like '' ")
        strSQL.Append("	) as tabela ")
        strSQL.Append(" where RH01_ID_PESSOA = " & PessoaId)

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function Pesquisar(Optional ByVal Sort As String = "",
                              Optional DocumentoId As Integer = 0,
                              Optional PessoaId As Integer = 0,
                              Optional EstadoCivilId As Integer = 0,
                              Optional GrauInstrucaoId As Integer = 0,
                              Optional MunicipioNascimentoId As Integer = 0,
                              Optional OrgaoEmissorId As Integer = 0,
                              Optional SiglaUfEmissorId As Integer = 0,
                              Optional NumeroRg As String = "",
                              Optional RgDigito As String = "",
                              Optional RgDataEmissao As String = "",
                              Optional PisPasep As String = "",
                              Optional TituloEleitor As String = "",
                              Optional SecaoEleitor As String = "",
                              Optional ZonaEleitor As String = "",
                              Optional CertificadoReservista As String = "",
                              Optional RegiaoReservista As String = "",
                              Optional NumeroConselho As String = "",
                              Optional DataConselho As String = "",
                              Optional SiglaExtensaoArquivoImagem As String = "",
                              Optional ImagemPessoa As String = "",
                              Optional ImagemDocumentoOficialIdentificacao As String = "",
                              Optional ImagemCpf As String = "",
                              Optional ImagemPisPasep As String = "",
                              Optional ImagemComprovanteResidencia As String = "",
                              Optional ExtensaoImgDocumentoOficialIdentificacao As String = "",
                              Optional ExtensaoImgCpf As String = "",
                              Optional ExtensaoImgPisPasep As String = "",
                              Optional ExtensaoImgComprovanteResidencia As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH67_DOCUMENTO_PESSOA")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where RH67_ID_DOCUMENTO_PESSOA is not null")

        If DocumentoId > 0 Then
            strSQL.Append(" and RH67_ID_DOCUMENTO_PESSOA = " & DocumentoId)
        End If

        If PessoaId > 0 Then
            strSQL.Append(" and RH01_ID_PESSOA = " & PessoaId)
        End If

        If EstadoCivilId > 0 Then
            strSQL.Append(" and TG12_ID_ESTADO_CIVIL = " & EstadoCivilId)
        End If

        If GrauInstrucaoId > 0 Then
            strSQL.Append(" and TG14_ID_GRAU_INSTRUCAO = " & GrauInstrucaoId)
        End If

        If MunicipioNascimentoId > 0 Then
            strSQL.Append(" and TG03_ID_MUNICIPIO_NASCIMENTO = " & MunicipioNascimentoId)
        End If

        If OrgaoEmissorId > 0 Then
            strSQL.Append(" and TG43_ID_ORGAO_EMISSOR_RG = " & OrgaoEmissorId)
        End If

        If SiglaUfEmissorId > 0 Then
            strSQL.Append(" and RH67_SG_UF_EMISSOR_RG =" & SiglaUfEmissorId)
        End If

        If NumeroRg <> "" Then
            strSQL.Append(" and upper(RH67_NU_RG) like '%" & NumeroRg.ToUpper & "%'")
        End If

        If RgDigito <> "" Then
            strSQL.Append(" and upper(RH67_NU_DIGITO_RG) like '%" & RgDigito.ToUpper & "%'")
        End If

        If RgDataEmissao <> "" Then
            strSQL.Append(" and upper(RH67_DT_EMISSAO_RG) like '%" & RgDataEmissao.ToUpper & "%'")
        End If

        If PisPasep <> "" Then
            strSQL.Append(" and upper(RH67_NU_PIS_PASEP) like '%" & PisPasep.ToUpper & "%'")
        End If

        If TituloEleitor <> "" Then
            strSQL.Append(" and upper(RH67_NU_TITULO_ELEITOR) like '%" & TituloEleitor.ToUpper & "%'")
        End If

        If SecaoEleitor <> "" Then
            strSQL.Append(" and upper(RH67_CD_SECAO_ELEITOR) like '%" & SecaoEleitor.ToUpper & "%'")
        End If

        If ZonaEleitor <> "" Then
            strSQL.Append(" and upper(RH67_CD_ZONA_ELEITOR) like '%" & ZonaEleitor.ToUpper & "%'")
        End If

        If CertificadoReservista <> "" Then
            strSQL.Append(" and upper(RH67_NU_CERTIFICADO_RESERVISTA) like '%" & CertificadoReservista.ToUpper & "%'")
        End If

        If RegiaoReservista <> "" Then
            strSQL.Append(" and upper(RH67_NM_REGIAO_RESERVISTA) like '%" & RegiaoReservista.ToUpper & "%'")
        End If

        If NumeroConselho <> "" Then
            strSQL.Append(" and upper(RH67_NU_CONSELHO) like '%" & NumeroConselho.ToUpper & "%'")
        End If

        If DataConselho <> "" Then
            strSQL.Append(" and upper(RH67_DT_CONSELHO) like '%" & DataConselho.ToUpper & "%'")
        End If

        If SiglaExtensaoArquivoImagem <> "" Then
            strSQL.Append(" and upper(RH67_SG_EXTENSAO_ARQ_IMAGEM) like '%" & SiglaExtensaoArquivoImagem.ToUpper & "%'")
        End If

        If ImagemPessoa <> "" Then
            strSQL.Append(" and upper(RH67_IM_PESSOA) like '%" & ImagemPessoa.ToUpper & "%'")
        End If

        If ImagemDocumentoOficialIdentificacao <> "" Then
            strSQL.Append(" and upper(RH67_IM_DOC_OFICIAL_IDENTIFIC) like '%" & ImagemDocumentoOficialIdentificacao.ToUpper & "%'")
        End If

        If ImagemCpf <> "" Then
            strSQL.Append(" and upper(RH67_IM_CPF) like '%" & ImagemCpf.ToUpper & "%'")
        End If

        If ImagemPisPasep <> "" Then
            strSQL.Append(" and upper(RH67_IM_PIS_PASEP) like '%" & ImagemPisPasep.ToUpper & "%'")
        End If

        If ImagemComprovanteResidencia <> "" Then
            strSQL.Append(" and upper(RH67_IM_COMPROVANTE_RESIDENCIA) like '%" & ImagemComprovanteResidencia.ToUpper & "%'")
        End If

        If ExtensaoImgDocumentoOficialIdentificacao <> "" Then
            strSQL.Append(" and upper(RH67_SG_EXTENSAO_IMG_IDENTIFIC) like '%" & ExtensaoImgDocumentoOficialIdentificacao.ToUpper & "%'")
        End If

        If ExtensaoImgCpf <> "" Then
            strSQL.Append(" and upper(RH67_SG_EXTENSAO_IMG_CPF) like '%" & ExtensaoImgCpf.ToUpper & "%'")
        End If

        If ExtensaoImgPisPasep <> "" Then
            strSQL.Append(" and upper(RH67_SG_EXTENSAO_IMG_PIS_PASEP) like '%" & ExtensaoImgPisPasep.ToUpper & "%'")
        End If

        If ExtensaoImgComprovanteResidencia <> "" Then
            strSQL.Append(" and upper(RH67_SG_EXTENSAO_IMG_COMP_RES) like '%" & ExtensaoImgComprovanteResidencia.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH67_ID_DOCUMENTO_PESSOA", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH67_ID_DOCUMENTO_PESSOA as CODIGO, RH01_ID_PESSOA as DESCRICAO")
		strSQL.Append(" from RH67_DOCUMENTO_PESSOA")
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
		
		strSQL.Append(" select max(RH67_ID_DOCUMENTO_PESSOA) from RH67_DOCUMENTO_PESSOA")

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
	Public Function Excluir(ByVal DocumentoId as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH67_DOCUMENTO_PESSOA")
		strSQL.Append(" where RH67_ID_DOCUMENTO_PESSOA = " & DocumentoId)

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

