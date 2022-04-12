Imports Microsoft.VisualBasic
Imports System.Data

Public Class VinculoOriginal
    Implements IDisposable
    Private ID As Integer
    Private Constitucional_Outros as String
	Private Justificativa as String
	Private Justificado as String
    Private Nome_CNPJ As String
    Private CNPJ as String
	Private Mes_Folha as String
	Private Ano_Folha as String
	Private Tipo_Folha as String
	Private Tipo_Vinculo as String
	Private Nome_Servidor_Pensionista as String
	Private CPF as String
	Private Matricula as String
	Private Regime as String
	Private Cargo as String
	Private Natureza_Cargo as String
	Private Data_Exercicio as String
	Private Data_Aposentadoria as String
	Private Data_Exclusao as String
	Private Carga_Horaria as String
	Private Categoria_Situacao as String
	Private Valor_Bruto as String
	Private Valor_Liquido as String
	Private CBO_Cargo as String
	Private Codigo_Cargo as String
	Private Unidade_Lotacao as String
	Private Codigo_Unidade_TCE As String

	Private RH00_IM_CONTRA_CHEQUE As String
	Private RH00_SG_CONTRA_CHEQUE As String

	Private RH00_IM_TERMO_POSSE As String
	Private RH00_SG_TERMO_POSSE As String

	Private RH00_IM_QUADRO_HORARIO As String
	Private RH00_SG_QUADRO_HORARIO As String

	Private RH00_IM_ATO_EXONERECAO As String
	Private RH00_SG_ATO_EXONERACAO As String

	Private RH00_JUSTIFICATIVA As String

	Private RH00_ANALISE As String


	Public Property IdVinculo() as Integer
		Get
			Return ID
		End Get
		Set(ByVal Value As Integer)
			ID = Value
		End Set
	End Property
	Public Property ConstitucionalOutros() as String
		Get
			Return Constitucional_Outros
		End Get
		Set(ByVal Value As String)
			Constitucional_Outros = Value
		End Set
	End Property
    Public Property _Justificativa() As String
        Get
            Return Justificativa
        End Get
        Set(ByVal Value As String)
            Justificativa = Value
        End Set
    End Property
    Public Property _Justificado() As String
        Get
            Return Justificado
        End Get
        Set(ByVal Value As String)
            Justificado = Value
        End Set
    End Property
    Public Property NomeCNPJ() As String
        Get
            Return Nome_CNPJ
        End Get
        Set(ByVal Value As String)
            Nome_CNPJ = Value
        End Set
    End Property
    Public Property _Cnpj() As String
        Get
            Return CNPJ
        End Get
        Set(ByVal Value As String)
            CNPJ = Value
        End Set
    End Property
    Public Property MesFolha() as String
		Get
			Return Mes_Folha
		End Get
		Set(ByVal Value As String)
			Mes_Folha = Value
		End Set
	End Property
	Public Property AnoFolha() as String
		Get
			Return Ano_Folha
		End Get
		Set(ByVal Value As String)
			Ano_Folha = Value
		End Set
	End Property
	Public Property TipoFolha() as String
		Get
			Return Tipo_Folha
		End Get
		Set(ByVal Value As String)
			Tipo_Folha = Value
		End Set
	End Property
	Public Property TipoVinculo() as String
		Get
			Return Tipo_Vinculo
		End Get
		Set(ByVal Value As String)
			Tipo_Vinculo = Value
		End Set
	End Property
	Public Property NomeServidorPensionita() as String
		Get
			Return Nome_Servidor_Pensionista
		End Get
		Set(ByVal Value As String)
			Nome_Servidor_Pensionista = Value
		End Set
	End Property
    Public Property _Cpf() As String
        Get
            Return CPF
        End Get
        Set(ByVal Value As String)
            CPF = Value
        End Set
    End Property
    Public Property _Matricula() As String
        Get
            Return Matricula
        End Get
        Set(ByVal Value As String)
            Matricula = Value
        End Set
    End Property
    Public Property _Regime() As String
        Get
            Return Regime
        End Get
        Set(ByVal Value As String)
            Regime = Value
        End Set
    End Property
    Public Property _Cargo() As String
        Get
            Return Cargo
        End Get
        Set(ByVal Value As String)
            Cargo = Value
        End Set
    End Property
    Public Property naturezaCargo() as String
		Get
			Return Natureza_Cargo
		End Get
		Set(ByVal Value As String)
			Natureza_Cargo = Value
		End Set
	End Property
	Public Property DataExercicio() as String
		Get
			Return Data_Exercicio
		End Get
		Set(ByVal Value As String)
			Data_Exercicio = Value
		End Set
	End Property
	Public Property DataAposentadoria() as String
		Get
			Return Data_Aposentadoria
		End Get
		Set(ByVal Value As String)
			Data_Aposentadoria = Value
		End Set
	End Property
	Public Property DataExclusao() as String
		Get
			Return Data_Exclusao
		End Get
		Set(ByVal Value As String)
			Data_Exclusao = Value
		End Set
	End Property
	Public Property CargaHoraria() as String
		Get
			Return Carga_Horaria
		End Get
		Set(ByVal Value As String)
			Carga_Horaria = Value
		End Set
	End Property
	Public Property CategoriaSituacao() as String
		Get
			Return Categoria_Situacao
		End Get
		Set(ByVal Value As String)
			Categoria_Situacao = Value
		End Set
	End Property
	Public Property ValorBruto() as String
		Get
			Return Valor_Bruto
		End Get
		Set(ByVal Value As String)
			Valor_Bruto = Value
		End Set
	End Property
	Public Property ValorLiquido() as String
		Get
			Return Valor_Liquido
		End Get
		Set(ByVal Value As String)
			Valor_Liquido = Value
		End Set
	End Property
	Public Property CboCargo() as String
		Get
			Return CBO_Cargo
		End Get
		Set(ByVal Value As String)
			CBO_Cargo = Value
		End Set
	End Property
	Public Property CodigoCargo() as String
		Get
			Return Codigo_Cargo
		End Get
		Set(ByVal Value As String)
			Codigo_Cargo = Value
		End Set
	End Property
	Public Property UnidadeLotacao() as String
		Get
			Return Unidade_Lotacao
		End Get
		Set(ByVal Value As String)
			Unidade_Lotacao = Value
		End Set
	End Property
	Public Property CodigoUnidadeTce() As String
		Get
			Return Codigo_Unidade_TCE
		End Get
		Set(ByVal Value As String)
			Codigo_Unidade_TCE = Value
		End Set
	End Property
	Public Property ImgContraCheque() As String
		Get
			Return RH00_IM_CONTRA_CHEQUE
		End Get
		Set(value As String)
			RH00_IM_CONTRA_CHEQUE = value
		End Set
	End Property
	Public Property ExtensaoImgContraCheque() As String
		Get
			Return RH00_SG_CONTRA_CHEQUE
		End Get
		Set(value As String)
			RH00_SG_CONTRA_CHEQUE = value
		End Set
	End Property
	Public Property ImgTermoPosse() As String
		Get
			Return RH00_IM_TERMO_POSSE
		End Get
		Set(value As String)
			RH00_IM_TERMO_POSSE = value
		End Set
	End Property
	Public Property ExtensaoImgTermoPosse() As String
		Get
			Return RH00_SG_TERMO_POSSE
		End Get
		Set(value As String)
			RH00_SG_TERMO_POSSE = value
		End Set
	End Property
	Public Property ImgQuadroHorario() As String
		Get
			Return RH00_IM_QUADRO_HORARIO
		End Get
		Set(value As String)
			RH00_IM_QUADRO_HORARIO = value
		End Set
	End Property
	Public Property ExtensaoImgQuadroHorario() As String
		Get
			Return RH00_SG_QUADRO_HORARIO
		End Get
		Set(value As String)
			RH00_SG_QUADRO_HORARIO = value
		End Set
	End Property
	Public Property ImgAtoExoneracao() As String
		Get
			Return RH00_IM_ATO_EXONERECAO
		End Get
		Set(value As String)
			RH00_IM_ATO_EXONERECAO = value
		End Set
	End Property
	Public Property ExtensaoImgAtoExoneracao() As String
		Get
			Return RH00_SG_ATO_EXONERACAO
		End Get
		Set(value As String)
			RH00_SG_ATO_EXONERACAO = value
		End Set
	End Property

	Public Property JustificativaTextoLivre() As String
		Get
			Return RH00_JUSTIFICATIVA
		End Get
		Set(value As String)
			RH00_JUSTIFICATIVA = value
		End Set
	End Property

	Public Property Analise() As String
		Get
			Return RH00_ANALISE
		End Get
		Set(value As String)
			RH00_ANALISE = value
		End Set
	End Property

	Public Sub New(Optional ByVal IdVinculo As Integer = 0)
        If IdVinculo > 0 Then
            Obter(IdVinculo)
        End If
    End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder

		strSQL.Append(" select * ")
		strSQL.Append(" from VN00_VINCULO_ORIGINAL")
		strSQL.Append(" where ID = " & IdVinculo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("Justificativa") = ProBanco(Justificativa, eTipoValor.TEXTO_LIVRE)
		dr("Justificado") = ProBanco(Justificado, eTipoValor.TEXTO)
		dr("RH00_JUSTIFICATIVA") = ProBanco(RH00_JUSTIFICATIVA, eTipoValor.TEXTO_LIVRE)

		dr("RH00_IM_CONTRA_CHEQUE") = ProBanco(RH00_IM_CONTRA_CHEQUE, eTipoValor.TEXTO_LIVRE)
		dr("RH00_SG_CONTRA_CHEQUE") = ProBanco(RH00_SG_CONTRA_CHEQUE, eTipoValor.TEXTO_LIVRE)

		dr("RH00_IM_TERMO_POSSE") = ProBanco(RH00_IM_TERMO_POSSE, eTipoValor.TEXTO_LIVRE)
		dr("RH00_SG_TERMO_POSSE") = ProBanco(RH00_SG_TERMO_POSSE, eTipoValor.TEXTO_LIVRE)

		dr("RH00_IM_QUADRO_HORARIO") = ProBanco(RH00_IM_QUADRO_HORARIO, eTipoValor.TEXTO_LIVRE)
		dr("RH00_SG_QUADRO_HORARIO") = ProBanco(RH00_SG_QUADRO_HORARIO, eTipoValor.TEXTO_LIVRE)

		dr("RH00_IM_ATO_EXONERECAO") = ProBanco(RH00_IM_ATO_EXONERECAO, eTipoValor.TEXTO_LIVRE)
		dr("RH00_SG_ATO_EXONERACAO") = ProBanco(RH00_SG_ATO_EXONERACAO, eTipoValor.TEXTO_LIVRE)

		dr("RH00_ANALISE") = ProBanco(RH00_ANALISE, eTipoValor.NUMERO_INTEIRO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function PesquisarDocumento(ByVal VinculoId As Integer) As DataTable

		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append("select * ")
		strSQL.Append("	from( ")
		strSQL.Append("		select 1 as TipoDocumento, 'CONTRACHEQUE' as NomeDocumento, RH00_IM_CONTRA_CHEQUE as DocumentoBase64, RH00_SG_CONTRA_CHEQUE as ExtensaoDocumento, ID, 'Imagens/Icones/' + REPLACE(RH00_SG_CONTRA_CHEQUE,'.','') + '.png' as ImageUrl 	 ")
		strSQL.Append("		from VN00_VINCULO_ORIGINAL ")
		strSQL.Append("		where RH00_IM_CONTRA_CHEQUE not like '' and RH00_SG_CONTRA_CHEQUE  not like '' ")
		strSQL.Append("		union all ")
		strSQL.Append("		select 2 as TipoDocumento, 'TERMO DE POSSE' as NomeDocumento, RH00_IM_TERMO_POSSE as DocumentoBase64, RH00_SG_TERMO_POSSE  as ExtensaoDocumento, ID, 'Imagens/Icones/' + REPLACE(RH00_SG_TERMO_POSSE,'.','') + '.png' as ImageUrl 	 ")
		strSQL.Append("		from VN00_VINCULO_ORIGINAL ")
		strSQL.Append("		where RH00_IM_TERMO_POSSE  not like '' and RH00_SG_TERMO_POSSE  not like '' ")
		strSQL.Append("		union all ")
		strSQL.Append("		select 3 as TipoDocumento, 'QUADRO DE HORÁRIO' as NomeDocumento, RH00_IM_QUADRO_HORARIO  as DocumentoBase64, RH00_SG_QUADRO_HORARIO  as ExtensaoDocumento, ID, 'Imagens/Icones/' + REPLACE(RH00_SG_QUADRO_HORARIO,'.','') + '.png' as ImageUrl 	 ")
		strSQL.Append("		from VN00_VINCULO_ORIGINAL ")
		strSQL.Append("		where RH00_IM_QUADRO_HORARIO  not like '' and RH00_SG_QUADRO_HORARIO not like '' ")
		strSQL.Append("		union all ")
		strSQL.Append("		select 4 as TipoDocumento, 'ATO DE EXONERAÇÃO' as NomeDocumento, RH00_IM_ATO_EXONERECAO as DocumentoBase64, RH00_SG_ATO_EXONERACAO as ExtensaoDocumento, ID, 'Imagens/Icones/' + REPLACE(RH00_SG_ATO_EXONERACAO,'.','') + '.png' as ImageUrl 	 ")
		strSQL.Append("		from VN00_VINCULO_ORIGINAL ")
		strSQL.Append("		where RH00_IM_ATO_EXONERECAO not like '' and RH00_SG_ATO_EXONERACAO not like '' ")
		strSQL.Append("	) as tabela ")
		strSQL.Append(" where id = " & VinculoId)

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Sub Obter(ByVal IdVinculo as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
        strSQL.Append(" from VN00_VINCULO_ORIGINAL")
        strSQL.Append(" where ID = " & IdVinculo)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)

			ID = DoBanco(dr("ID"), eTipoValor.CHAVE)
			Constitucional_Outros = DoBanco(dr("Constitucional_Outros"), eTipoValor.TEXTO)
			Justificativa = DoBanco(dr("Justificativa"), eTipoValor.TEXTO_LIVRE)
			Justificado = DoBanco(dr("Justificado"), eTipoValor.TEXTO)
			Nome_CNPJ = DoBanco(dr("Nome_CNPJ"), eTipoValor.TEXTO)
			CNPJ = DoBanco(dr("CNPJ"), eTipoValor.TEXTO)
			Mes_Folha = DoBanco(dr("Mes_Folha"), eTipoValor.TEXTO)
			Ano_Folha = DoBanco(dr("Ano_Folha"), eTipoValor.TEXTO)
			Tipo_Folha = DoBanco(dr("Tipo_Folha"), eTipoValor.TEXTO)
			Tipo_Vinculo = DoBanco(dr("Tipo_Vinculo"), eTipoValor.TEXTO)
			Nome_Servidor_Pensionista = DoBanco(dr("Nome_Servidor_Pensionista"), eTipoValor.TEXTO)
			CPF = DoBanco(dr("CPF"), eTipoValor.TEXTO)
			Matricula = DoBanco(dr("Matricula"), eTipoValor.TEXTO)
			Regime = DoBanco(dr("Regime"), eTipoValor.TEXTO)
			Cargo = DoBanco(dr("Cargo"), eTipoValor.TEXTO)
			Natureza_Cargo = DoBanco(dr("Natureza_Cargo"), eTipoValor.TEXTO)
			Data_Exercicio = DoBanco(dr("Data_Exercicio"), eTipoValor.TEXTO)
			Data_Aposentadoria = DoBanco(dr("Data_Aposentadoria"), eTipoValor.TEXTO)
			Data_Exclusao = DoBanco(dr("Data_Exclusao"), eTipoValor.TEXTO)
			Carga_Horaria = DoBanco(dr("Carga_Horaria"), eTipoValor.TEXTO)
			Categoria_Situacao = DoBanco(dr("Categoria_Situacao"), eTipoValor.TEXTO)
			Valor_Bruto = DoBanco(dr("Valor_Bruto"), eTipoValor.TEXTO)
			Valor_Liquido = DoBanco(dr("Valor_Liquido"), eTipoValor.TEXTO)
			CBO_Cargo = DoBanco(dr("CBO_Cargo"), eTipoValor.TEXTO)
			Codigo_Cargo = DoBanco(dr("Codigo_Cargo"), eTipoValor.TEXTO)
			Unidade_Lotacao = DoBanco(dr("Unidade_Lotacao"), eTipoValor.TEXTO)
			Codigo_Unidade_TCE = DoBanco(dr("Codigo_Unidade_TCE"), eTipoValor.TEXTO)

			RH00_IM_CONTRA_CHEQUE = DoBanco(dr("RH00_IM_CONTRA_CHEQUE"), eTipoValor.TEXTO_LIVRE)
			RH00_SG_CONTRA_CHEQUE = DoBanco(dr("RH00_SG_CONTRA_CHEQUE"), eTipoValor.TEXTO_LIVRE)

			RH00_IM_TERMO_POSSE = DoBanco(dr("RH00_IM_TERMO_POSSE"), eTipoValor.TEXTO_LIVRE)
			RH00_SG_TERMO_POSSE = DoBanco(dr("RH00_SG_TERMO_POSSE"), eTipoValor.TEXTO_LIVRE)

			RH00_IM_QUADRO_HORARIO = DoBanco(dr("RH00_IM_QUADRO_HORARIO"), eTipoValor.TEXTO_LIVRE)
			RH00_SG_QUADRO_HORARIO = DoBanco(dr("RH00_SG_QUADRO_HORARIO"), eTipoValor.TEXTO_LIVRE)

			RH00_IM_ATO_EXONERECAO = DoBanco(dr("RH00_IM_ATO_EXONERECAO"), eTipoValor.TEXTO_LIVRE)
			RH00_SG_ATO_EXONERACAO = DoBanco(dr("RH00_SG_ATO_EXONERACAO"), eTipoValor.TEXTO_LIVRE)

			RH00_JUSTIFICATIVA = DoBanco(dr("RH00_JUSTIFICATIVA"), eTipoValor.TEXTO_LIVRE)

			RH00_ANALISE = DoBanco(dr("RH00_ANALISE"), eTipoValor.NUMERO_INTEIRO)



		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort As String = "", Optional IdVinculo As Integer = 0, Optional NomeServidorPensionita As String = "", Optional Cpf As String = "") As DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append(" select *,upper(Unidade_Lotacao) as LOTACAO ")
		strSQL.Append(" from VN00_VINCULO_ORIGINAL")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where ID is not null")
		strSQL.Append(" and RH00_EXCLUIDO = 0")

		If IdVinculo > 0 Then
			strSQL.Append(" and ID =  " & IdVinculo)
		End If

		If NomeServidorPensionita <> "" Then
			strSQL.Append(" and upper(Nome_Servidor_Pensionista) like '%" & NomeServidorPensionita.ToUpper & "%'")
		End If


		If Replace(Replace(Replace(Cpf, ".", ""), "-", ""), "_", "") <> "" Then
			strSQL.Append(" and upper(CPF) = '" & Replace(Replace(Cpf.ToUpper, ".", ""), "-", "") & "'")
		End If

		strSQL.Append(" Order By " & IIf(Sort = "", "ID", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function
	Public Function Pesquisar2(Optional ByVal Sort As String = "", Optional IdVinculo As Integer = 0, Optional NomeServidorPensionita As String = "", Optional Cpf As String = "" _
							   , Optional ByVal ExibirVinculos As Boolean = False, Optional ByVal Justificados As Integer = 0 _
							   , Optional ByVal Cargo As String = "", Optional ByVal Orgao As String = "") As DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append(" select ID,CPF, Matricula, Nome_Servidor_Pensionista, NOME_CNPJ,Justificado,upper(Unidade_Lotacao) as LOTACAO, case Justificado when 'S' then 'JUSTIFICADO' else 'PENDENTE' end as SITUACAO  ")
		strSQL.Append(" from VN00_VINCULO_ORIGINAL")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where ID is not null")

		If IdVinculo > 0 Then
			strSQL.Append(" and ID =  " & IdVinculo)
		End If

		If Cargo <> "" Then
			strSQL.Append(" and upper(cargo) like '%" & Cargo.ToUpper & "%'")
		End If

		If Orgao <> "" Then
			strSQL.Append(" and upper(Nome_CNPJ) like '%" & Orgao.ToUpper & "%'")
		End If

		If Justificados > 0 Then
			strSQL.Append(" and Justificado = 'S' ")
		End If

		If NomeServidorPensionita <> "" Then
			strSQL.Append(" and upper(Nome_Servidor_Pensionista) like '%" & NomeServidorPensionita.ToUpper & "%'")
		End If

		'If Cpf <> "" Then
		'    strSQL.Append(" and upper(CPF) like '%" & Cpf.ToUpper & "%'")
		'End If


		If Replace(Replace(Replace(Cpf, ".", ""), "-", ""), "_", "") <> "" Then
			strSQL.Append(" and upper(CPF) = '" & Replace(Replace(Cpf.ToUpper, ".", ""), "-", "") & "'")
		End If

		strSQL.Append(" Order By " & IIf(Sort = "", "Nome_Servidor_Pensionista", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function
	Public Function PesquisarNome(Optional ByVal Sort As String = "", Optional IdVinculo As Integer = 0, Optional NomeServidorPensionita As String = "", Optional Cpf As String = "" _
							   , Optional ByVal ExibirVinculos As Boolean = False, Optional ByVal Justificados As Integer = 0 _
							   , Optional ByVal Cargo As String = "", Optional ByVal Orgao As String = "") As DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append(" select distinct CPF, Nome_Servidor_Pensionista, justificado,isnuLL(RH00_ANALISE,0) as RH00_ANALISE  ")
		strSQL.Append(" from VN00_VINCULO_ORIGINAL")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where ID is not null")

		If IdVinculo > 0 Then
			strSQL.Append(" and ID =  " & IdVinculo)
		End If

		If Justificados > 0 Then
			strSQL.Append(" and Justificado = 'S' ")
		End If

		If NomeServidorPensionita <> "" Then
			strSQL.Append(" and upper(Nome_Servidor_Pensionista) like '%" & NomeServidorPensionita.ToUpper & "%'")
		End If

		If Cargo <> "" Then
			strSQL.Append(" and upper(cargo) like '%" & Cargo.ToUpper & "%'")
		End If

		If Orgao <> "" Then
			strSQL.Append(" and upper(Nome_CNPJ) like '%" & Orgao.ToUpper & "%'")
		End If

		'If Cpf <> "" Then
		'    strSQL.Append(" and upper(CPF) like '%" & Cpf.ToUpper & "%'")
		'End If


		If Replace(Replace(Replace(Cpf, ".", ""), "-", ""), "_", "") <> "" Then
			strSQL.Append(" and upper(CPF) = '" & Replace(Replace(Cpf.ToUpper, ".", ""), "-", "") & "'")
		End If

		strSQL.Append(" Order By " & IIf(Sort = "", "Nome_Servidor_Pensionista", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select ID as CODIGO, Constitucional_Outros as DESCRICAO")
        strSQL.Append(" from VN00_VINCULO_ORIGINAL")
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

        strSQL.Append(" select max(ID) from VN00_VINCULO_ORIGINAL")

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
	Public Function Excluir(ByVal IdVinculo as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
        strSQL.Append(" from VN00_VINCULO_ORIGINAL")
        strSQL.Append(" where ID = " & IdVinculo)

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

