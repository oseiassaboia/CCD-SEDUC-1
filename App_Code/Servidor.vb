Imports Microsoft.VisualBasic
Imports System.Data

Public Class Servidor
    Implements IDisposable

    Private RH02_ID_SERVIDOR as Integer
	Private RH01_ID_PESSOA as String
	Private RH16_ID_CARGO as String
	Private RH39_ID_CATEGORIA_FUNCIONAL as String
	Private RH35_ID_SIMBOLOGIA as String
	Private RH04_ID_ORGAO as String
	Private RH05_ID_TIPO_VINCULO as String
	Private RH07_ID_SITUACAO_SERVIDOR as String
	Private CA04_ID_USUARIO as Integer
	Private RH02_CD_MATRICULA as String
	Private RH02_DT_ADMISSAO as String
	Private RH02_DT_POSSE as String
	Private RH02_DT_NOMEACAO as String
	Private RH02_DT_SOLICITACAO_MATRIC as String
	Private RH02_DT_ANUENIO_FUNCAO as String
	Private RH02_DT_TRANSFERENCIA as String
	Private RH02_DT_RETORNO as String
	Private RH02_DT_DEMISSAO as String
	Private RH02_IN_ATO_PROVIMENTO as String
	Private RH02_DH_CADASTRO as String
	Private RH02_DH_ULTIMA_CARGA as String
	Private RH02_QT_HR_SEMANAL_REDUZIDA as String
	Private RH02_IN_REDUCAO_CG_HORARIA as Boolean
	Private RH02_NU_PORTARIA as String
    Private _ServidorCargo  As String
    Private RH02_IN_AMPLIACAO_CG_HORARIA as Boolean


	Public Property ServidorId() as Integer
		Get
			Return RH02_ID_SERVIDOR
		End Get
		Set(ByVal Value As Integer)
			RH02_ID_SERVIDOR = Value
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
	Public Property CargoId() as String
		Get
			Return RH16_ID_CARGO
		End Get
		Set(ByVal Value As String)
			RH16_ID_CARGO = Value
		End Set
	End Property
	Public Property CategoriaFuncaoId() as String
		Get
			Return RH39_ID_CATEGORIA_FUNCIONAL
		End Get
		Set(ByVal Value As String)
			RH39_ID_CATEGORIA_FUNCIONAL = Value
		End Set
	End Property
	Public Property Simbologia() as String
		Get
			Return RH35_ID_SIMBOLOGIA
		End Get
		Set(ByVal Value As String)
			RH35_ID_SIMBOLOGIA = Value
		End Set
	End Property
	Public Property OrgaoId() as String
		Get
			Return RH04_ID_ORGAO
		End Get
		Set(ByVal Value As String)
			RH04_ID_ORGAO = Value
		End Set
	End Property
	Public Property TipoVinculoId() as String
		Get
			Return RH05_ID_TIPO_VINCULO
		End Get
		Set(ByVal Value As String)
			RH05_ID_TIPO_VINCULO = Value
		End Set
	End Property
	Public Property SituacaoServidorId() as String
		Get
			Return RH07_ID_SITUACAO_SERVIDOR
		End Get
		Set(ByVal Value As String)
			RH07_ID_SITUACAO_SERVIDOR = Value
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
	Public Property Matricula() as String
		Get
			Return RH02_CD_MATRICULA
		End Get
		Set(ByVal Value As String)
			RH02_CD_MATRICULA = Value
		End Set
	End Property
	Public Property DataAdmissao() as String
		Get
			Return RH02_DT_ADMISSAO
		End Get
		Set(ByVal Value As String)
			RH02_DT_ADMISSAO = Value
		End Set
	End Property
	Public Property DataPosse() as String
		Get
			Return RH02_DT_POSSE
		End Get
		Set(ByVal Value As String)
			RH02_DT_POSSE = Value
		End Set
	End Property
	Public Property DataNomeacao() as String
		Get
			Return RH02_DT_NOMEACAO
		End Get
		Set(ByVal Value As String)
			RH02_DT_NOMEACAO = Value
		End Set
	End Property
	Public Property SolicitacaoMatricula() as String
		Get
			Return RH02_DT_SOLICITACAO_MATRIC
		End Get
		Set(ByVal Value As String)
			RH02_DT_SOLICITACAO_MATRIC = Value
		End Set
	End Property
	Public Property AnuenioFuncao() as String
		Get
			Return RH02_DT_ANUENIO_FUNCAO
		End Get
		Set(ByVal Value As String)
			RH02_DT_ANUENIO_FUNCAO = Value
		End Set
	End Property
	Public Property DataTransferencia() as String
		Get
			Return RH02_DT_TRANSFERENCIA
		End Get
		Set(ByVal Value As String)
			RH02_DT_TRANSFERENCIA = Value
		End Set
	End Property
	Public Property DataRetorno() as String
		Get
			Return RH02_DT_RETORNO
		End Get
		Set(ByVal Value As String)
			RH02_DT_RETORNO = Value
		End Set
	End Property
	Public Property DataDemissao() as String
		Get
			Return RH02_DT_DEMISSAO
		End Get
		Set(ByVal Value As String)
			RH02_DT_DEMISSAO = Value
		End Set
	End Property
	Public Property AtoProvimento() as String
		Get
			Return RH02_IN_ATO_PROVIMENTO
		End Get
		Set(ByVal Value As String)
			RH02_IN_ATO_PROVIMENTO = Value
		End Set
	End Property
	Public Property DataHoraCadastro() as String
		Get
			Return RH02_DH_CADASTRO
		End Get
		Set(ByVal Value As String)
			RH02_DH_CADASTRO = Value
		End Set
	End Property
	Public Property DataHoraUltimaCarga() as String
		Get
			Return RH02_DH_ULTIMA_CARGA
		End Get
		Set(ByVal Value As String)
			RH02_DH_ULTIMA_CARGA = Value
		End Set
	End Property
	Public Property QtdHoraSemanalReduzida() as String
		Get
			Return RH02_QT_HR_SEMANAL_REDUZIDA
		End Get
		Set(ByVal Value As String)
			RH02_QT_HR_SEMANAL_REDUZIDA = Value
		End Set
	End Property
	Public Property AmpliacaoCargaHoraria() as Boolean
		Get
			Return RH02_IN_AMPLIACAO_CG_HORARIA
		End Get
		Set(ByVal Value As Boolean)
		    RH02_IN_AMPLIACAO_CG_HORARIA = Value
		End Set
	End Property
    Public Property ReducaoCargaHoraria() as Boolean
        Get
            Return RH02_IN_REDUCAO_CG_HORARIA
        End Get
        Set(ByVal Value As Boolean)
            RH02_IN_REDUCAO_CG_HORARIA = Value
        End Set
    End Property
	Public Property NumeroPortaria() as String
		Get
			Return RH02_NU_PORTARIA
		End Get
		Set(ByVal Value As String)
			RH02_NU_PORTARIA = Value
		End Set
	End Property
    Public  Property ServidorCargo() As String
        get
            Return _ServidorCargo
        End Get

    set(value As String)
            'Não Implementado
    End Set

    End Property

	Public Sub New(Optional ByVal ServidorId as Integer = 0,Optional ByVal PessoaId As Integer = 0)
		If ServidorId > 0 or PessoaId > 0 Then
			Obter(ServidorId,PessoaId)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH02_SERVIDOR")
		strSQL.Append(" where RH02_ID_SERVIDOR = " & ServidorId)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

        dr("RH01_ID_PESSOA") = ProBanco(RH01_ID_PESSOA, eTipoValor.CHAVE)
        dr("RH16_ID_CARGO") = ProBanco(RH16_ID_CARGO, eTipoValor.CHAVE)
        dr("RH39_ID_CATEGORIA_FUNCIONAL") = ProBanco(RH39_ID_CATEGORIA_FUNCIONAL, eTipoValor.CHAVE)
        dr("RH35_ID_SIMBOLOGIA") = ProBanco(RH35_ID_SIMBOLOGIA, eTipoValor.CHAVE)
        dr("RH04_ID_ORGAO") = ProBanco(RH04_ID_ORGAO, eTipoValor.CHAVE)
        dr("RH05_ID_TIPO_VINCULO") = ProBanco(RH05_ID_TIPO_VINCULO, eTipoValor.CHAVE)
        dr("RH07_ID_SITUACAO_SERVIDOR") = ProBanco(RH07_ID_SITUACAO_SERVIDOR, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("RH02_CD_MATRICULA") = ProBanco(RH02_CD_MATRICULA, eTipoValor.TEXTO)
		dr("RH02_DT_ADMISSAO") = ProBanco(RH02_DT_ADMISSAO, eTipoValor.DATA)
		dr("RH02_DT_POSSE") = ProBanco(RH02_DT_POSSE, eTipoValor.DATA)
		dr("RH02_DT_NOMEACAO") = ProBanco(RH02_DT_NOMEACAO, eTipoValor.DATA)
		dr("RH02_DT_SOLICITACAO_MATRIC") = ProBanco(RH02_DT_SOLICITACAO_MATRIC, eTipoValor.DATA)
		dr("RH02_DT_ANUENIO_FUNCAO") = ProBanco(RH02_DT_ANUENIO_FUNCAO, eTipoValor.DATA)
		dr("RH02_DT_TRANSFERENCIA") = ProBanco(RH02_DT_TRANSFERENCIA, eTipoValor.DATA)
		dr("RH02_DT_RETORNO") = ProBanco(RH02_DT_RETORNO, eTipoValor.DATA)
		dr("RH02_DT_DEMISSAO") = ProBanco(RH02_DT_DEMISSAO, eTipoValor.DATA)
		dr("RH02_IN_ATO_PROVIMENTO") = ProBanco(RH02_IN_ATO_PROVIMENTO, eTipoValor.DATA_COMPLETA)
		dr("RH02_DH_CADASTRO") = ProBanco(RH02_DH_CADASTRO, eTipoValor.DATA_COMPLETA)
		dr("RH02_DH_ULTIMA_CARGA") = ProBanco(RH02_DH_ULTIMA_CARGA, eTipoValor.DATA_COMPLETA)
		dr("RH02_QT_HR_SEMANAL_REDUZIDA") = ProBanco(RH02_QT_HR_SEMANAL_REDUZIDA, eTipoValor.NUMERO_DECIMAL)
		dr("RH02_IN_REDUCAO_CG_HORARIA") = ProBanco(RH02_IN_REDUCAO_CG_HORARIA, eTipoValor.BOOLEANO)
	    dr("RH02_IN_AMPLIACAO_CG_HORARIA") = ProBanco(RH02_IN_AMPLIACAO_CG_HORARIA, eTipoValor.BOOLEANO)
		dr("RH02_NU_PORTARIA") = ProBanco(RH02_NU_PORTARIA, eTipoValor.TEXTO)


		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter( Optional ByVal ServidorId as Integer = 0 , optional ByVal PessoaID As Integer = 0)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
	    strSQL.Append(" ,case when RH16_id_cargo  in (586,587,588,589,590,591,592,593,594,595,596,597,598,599,600,639,640,641,642,674,675,676,677,678,679,680,681,683,701) then 'PROFESSOR' ")
	    strSQL.Append(" else 'ADM' end as SERVIDORCARGO ")
		strSQL.Append(" from RH02_SERVIDOR")
		strSQL.Append(" where RH02_ID_SERVIDOR > 0")

        if ServidorId > 0 Then
            strSQL.Append(" and RH02_ID_SERVIDOR =" & ServidorId)
        End If

	    if PessoaID > 0 Then
	    strSQL.Append(" and RH01_ID_PESSOA=" & PessoaID)
	    End If
		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH02_ID_SERVIDOR = DoBanco(dr("RH02_ID_SERVIDOR"), eTipoValor.CHAVE)
            RH01_ID_PESSOA = DoBanco(dr("RH01_ID_PESSOA"), eTipoValor.CHAVE)
            RH16_ID_CARGO = DoBanco(dr("RH16_ID_CARGO"), eTipoValor.CHAVE)
            RH39_ID_CATEGORIA_FUNCIONAL = DoBanco(dr("RH39_ID_CATEGORIA_FUNCIONAL"), eTipoValor.CHAVE)
            RH35_ID_SIMBOLOGIA = DoBanco(dr("RH35_ID_SIMBOLOGIA"), eTipoValor.CHAVE)
            RH04_ID_ORGAO = DoBanco(dr("RH04_ID_ORGAO"), eTipoValor.CHAVE)
            RH05_ID_TIPO_VINCULO = DoBanco(dr("RH05_ID_TIPO_VINCULO"), eTipoValor.CHAVE)
            RH07_ID_SITUACAO_SERVIDOR = DoBanco(dr("RH07_ID_SITUACAO_SERVIDOR"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            RH02_CD_MATRICULA = DoBanco(dr("RH02_CD_MATRICULA"), eTipoValor.TEXTO)
			RH02_DT_ADMISSAO = DoBanco(dr("RH02_DT_ADMISSAO"), eTipoValor.DATA)
			RH02_DT_POSSE = DoBanco(dr("RH02_DT_POSSE"), eTipoValor.DATA)
			RH02_DT_NOMEACAO = DoBanco(dr("RH02_DT_NOMEACAO"), eTipoValor.DATA)
			RH02_DT_SOLICITACAO_MATRIC = DoBanco(dr("RH02_DT_SOLICITACAO_MATRIC"), eTipoValor.DATA)
			RH02_DT_ANUENIO_FUNCAO = DoBanco(dr("RH02_DT_ANUENIO_FUNCAO"), eTipoValor.DATA)
			RH02_DT_TRANSFERENCIA = DoBanco(dr("RH02_DT_TRANSFERENCIA"), eTipoValor.DATA)
			RH02_DT_RETORNO = DoBanco(dr("RH02_DT_RETORNO"), eTipoValor.DATA)
			RH02_DT_DEMISSAO = DoBanco(dr("RH02_DT_DEMISSAO"), eTipoValor.DATA)
			RH02_IN_ATO_PROVIMENTO = DoBanco(dr("RH02_IN_ATO_PROVIMENTO"), eTipoValor.TEXTO)
			RH02_DH_CADASTRO = DoBanco(dr("RH02_DH_CADASTRO"), eTipoValor.DATA_COMPLETA)
			RH02_DH_ULTIMA_CARGA = DoBanco(dr("RH02_DH_ULTIMA_CARGA"), eTipoValor.DATA_COMPLETA)
			RH02_QT_HR_SEMANAL_REDUZIDA = DoBanco(dr("RH02_QT_HR_SEMANAL_REDUZIDA"), eTipoValor.NUMERO_DECIMAL)
		    RH02_IN_REDUCAO_CG_HORARIA = DoBanco(dr("RH02_IN_REDUCAO_CG_HORARIA"), eTipoValor.BOOLEANO)
		    RH02_IN_AMPLIACAO_CG_HORARIA = DoBanco(dr("RH02_IN_AMPLIACAO_CG_HORARIA"), eTipoValor.BOOLEANO)
            RH02_NU_PORTARIA = DoBanco(dr("RH02_NU_PORTARIA"), eTipoValor.TEXTO)
		    _ServidorCargo = DoBanco(dr("SERVIDORCARGO"), eTipoValor.TEXTO)

		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub
    Public Function PesquisarServidor(Optional ByVal Sort as String = "", Optional ServidorId as Integer = 0, Optional PessoaId as Integer = 0, Optional CargoId as Integer = 0, Optional CategoriaFuncaoId as Integer = 0, Optional Simbologia as Integer = 0, Optional OrgaoId as Integer = 0, Optional TipoVinculoId as Integer = 0, Optional SituacaoServidorId as Integer = 0, Optional UsuarioId as Integer = 0, Optional Matricula as String = "", Optional DataAdmissao as String = "", Optional DataPosse as String = "", Optional DataNomeacao as String = "", Optional SolicitacaoMatricula as String = "", Optional AnuenioFuncao as String = "", Optional DataTransferencia as String = "", Optional DataRetorno as String = "", Optional DataDemissao as String = "", Optional AtoProvimento as String = "", Optional DataHoraCadastro as String = "", Optional DataHoraUltimaCarga as String = "", Optional QtdHoraSemanalReduzida as String = "", Optional ReducaoCargaHoraria as String = "", Optional NumeroPortaria as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select *,'MATRICULA: '+isnull(SERVIDOR.RH02_CD_MATRICULA,'') +' - '+'CARGO: ' + Cargo.RH16_NM_CARGO  as DESCRICAO,Servidor.RH02_ID_SERVIDOR as CODIGO   ")
	    strSQL.Append(" ,case when Cargo.RH16_id_cargo  in (586,587,588,589,590,591,592,593,594,595,596,597,598,599,600,639,640,641,642,674,675,676,677,678,679,680,681,683,701) then 'PROFESSOR' ")
	    strSQL.Append(" else 'ADM' end as SERVIDORCARGO ")
		strSQL.Append(" from RH02_SERVIDOR Servidor")
		strSQL.Append(" left join RH16_CARGO Cargo on Cargo.RH16_ID_CARGO = Servidor.RH16_ID_CARGO ")
	    strSQL.Append(" left join RH05_TIPO_VINCULO TipoVinculo on TipoVinculo.RH05_ID_TIPO_VINCULO = Servidor.RH05_ID_TIPO_VINCULO ")
        strSQL.Append(" left join RH07_SITUACAO_SERVIDOR SituacaoServidor on SituacaoServidor.RH07_ID_SITUACAO_SERVIDOR = Servidor.RH07_ID_SITUACAO_SERVIDOR ")
        strSQL.Append(" left join RH35_SIMBOLOGIA SimbologiaServidor on SimbologiaServidor.RH35_ID_SIMBOLOGIA = Servidor.RH35_ID_SIMBOLOGIA ")
        strSQL.Append(" where sERVIDOR.RH02_ID_SERVIDOR is not null ")

        If ServidorId > 0 then 
			strSQL.Append(" and Servidor.RH02_ID_SERVIDOR = " & ServidorId)
		End If
		
		If PessoaId> 0 then
			strSQL.Append(" and Servidor.RH01_ID_PESSOA = " & PessoaId)
		End If
		
		If  CargoId > 0 then
			strSQL.Append(" and Servidor.RH16_ID_CARGO = " & CargoId)
		End If
		
		If CategoriaFuncaoId> 0 then
			strSQL.Append(" and Servidor.RH39_ID_CATEGORIA_FUNCIONAL = " & CategoriaFuncaoId)
		End If
		
		If Simbologia > 0 then
			strSQL.Append(" and Servidor.RH35_ID_SIMBOLOGIA = " & Simbologia)
		End If
		
		If OrgaoId > 0  then
			strSQL.Append(" and Servidor.RH04_ID_ORGAO = " & OrgaoId)
		End If
		
		If  TipoVinculoId> 0 then
			strSQL.Append(" and Servidor.RH05_ID_TIPO_VINCULO = " & TipoVinculoId)
		End If
		
		If SituacaoServidorId > 0Then
          
            strSQL.Append(" and Servidor.RH07_ID_SITUACAO_SERVIDOR in (1,11,10 )")
        End If
		
		If UsuarioId > 0 then 
			strSQL.Append(" and CA04_ID_USUARIO = " & UsuarioId)
		End If
		
		If Matricula <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_CD_MATRICULA) like '%" & Matricula.toUpper & "%'")
		End If
		
		If DataAdmissao <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_DT_ADMISSAO) like '%" & DataAdmissao.toUpper & "%'")
		End If
		
		If DataPosse <> "" then 
			strSQL.Append(" and upper(vRH02_DT_POSSE) like '%" & DataPosse.toUpper & "%'")
		End If
		
		If DataNomeacao <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_DT_NOMEACAO) like '%" & DataNomeacao.toUpper & "%'")
		End If
		
		If SolicitacaoMatricula <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_DT_SOLICITACAO_MATRIC) like '%" & SolicitacaoMatricula.toUpper & "%'")
		End If
		
		If AnuenioFuncao <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_DT_ANUENIO_FUNCAO) like '%" & AnuenioFuncao.toUpper & "%'")
		End If
		
		If DataTransferencia <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_DT_TRANSFERENCIA) like '%" & DataTransferencia.toUpper & "%'")
		End If
		
		If DataRetorno <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_DT_RETORNO) like '%" & DataRetorno.toUpper & "%'")
		End If
		
		If DataDemissao <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_DT_DEMISSAO) like '%" & DataDemissao.toUpper & "%'")
		End If
		
		If AtoProvimento <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_IN_ATO_PROVIMENTO) like '%" & AtoProvimento.toUpper & "%'")
		End If
		
		If isDate(DataHoraCadastro) then 
			strSQL.Append(" and Servidor.RH02_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
		End If
		
		If isDate(DataHoraUltimaCarga) then 
			strSQL.Append(" and Servidor.RH02_DH_ULTIMA_CARGA = Convert(DateTime, '" & DataHoraUltimaCarga & "', 103)")
		End If
		
		If IsNumeric(QtdHoraSemanalReduzida.Replace(".", "")) then
			strSQL.Append(" and Servidor.RH02_QT_HR_SEMANAL_REDUZIDA = " & QtdHoraSemanalReduzida.Replace(".", "").Replace(",", "."))
		End If
		
		If ReducaoCargaHoraria <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_IN_REDUCAO_CARGA_HORARIA) like '%" & ReducaoCargaHoraria.toUpper & "%'")
		End If
		
		If NumeroPortaria <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_NU_PORTARIA) like '%" & NumeroPortaria.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "Servidor.RH02_ID_SERVIDOR", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function
	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional ServidorId as Integer = 0, Optional PessoaId as Integer = 0, Optional CargoId as Integer = 0, Optional CategoriaFuncaoId as Integer = 0, Optional Simbologia as Integer = 0, Optional OrgaoId as Integer = 0, Optional TipoVinculoId as Integer = 0, Optional SituacaoServidorId as Integer = 0, Optional UsuarioId as Integer = 0, Optional Matricula as String = "", Optional DataAdmissao as String = "", Optional DataPosse as String = "", Optional DataNomeacao as String = "", Optional SolicitacaoMatricula as String = "", Optional AnuenioFuncao as String = "", Optional DataTransferencia as String = "", Optional DataRetorno as String = "", Optional DataDemissao as String = "", Optional AtoProvimento as String = "", Optional DataHoraCadastro as String = "", Optional DataHoraUltimaCarga as String = "", Optional QtdHoraSemanalReduzida as String = "", Optional ReducaoCargaHoraria as String = "", Optional NumeroPortaria as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select *,'MATRICULA: '+isnull(SERVIDOR.RH02_CD_MATRICULA,'') +' - '+'CARGO: ' +Cargo.RH16_NM_CARGO +' - '+'FUNÇÃO: '+ isnull(funcao.RH06_NM_FUNCAO,'SEM FUNÇÃO')+' LOTAÇÃO: '+ isnull(lotacao.RH36_NM_LOTACAO,' - ') as DESCRICAO,Servidor.RH02_ID_SERVIDOR as CODIGO   ")
	    strSQL.Append(" ,case when Cargo.RH16_id_cargo  in (586,587,588,589,590,591,592,593,594,595,596,597,598,599,600,639,640,641,642,674,675,676,677,678,679,680,681,683,701) then 'PROFESSOR' ")
	    strSQL.Append(" else 'ADM' end as SERVIDORCARGO ")
		strSQL.Append(" from RH02_SERVIDOR Servidor")
		strSQL.Append(" left join RH16_CARGO Cargo on Cargo.RH16_ID_CARGO = Servidor.RH16_ID_CARGO ")
	    strSQL.Append(" left join RH05_TIPO_VINCULO TipoVinculo on TipoVinculo.RH05_ID_TIPO_VINCULO = Servidor.RH05_ID_TIPO_VINCULO ")
        strSQL.Append(" left join RH07_SITUACAO_SERVIDOR SituacaoServidor on SituacaoServidor.RH07_ID_SITUACAO_SERVIDOR = Servidor.RH07_ID_SITUACAO_SERVIDOR ")
        strSQL.Append(" left join RH35_SIMBOLOGIA SimbologiaServidor on SimbologiaServidor.RH35_ID_SIMBOLOGIA = Servidor.RH35_ID_SIMBOLOGIA ")
        strSQL.Append(" left join RH14_LOTACAO_SERVIDOR LotacaoServidor on LotacaoServidor.RH02_ID_SERVIDOR = Servidor.RH02_ID_SERVIDOR ")
	    strSQL.Append(" left join rh36_lotacao Lotacao on LotacaoServidor.RH36_ID_LOTACAO = Lotacao.RH36_ID_LOTACAO ")
        strSQL.Append(" left join RH06_FUNCAO Funcao on LotacaoServidor.RH06_ID_FUNCAO = Funcao.RH06_ID_FUNCAO ")
        strSQL.Append(" where sERVIDOR.RH02_ID_SERVIDOR is not null  and LotacaoServidor.RH14_DT_DESLIGAMENTO is null	")
	    'strSQL.Append(" and sERVIDOR.RH07_ID_SITUACAO_SERVIDOR =" & Servidor.Situacao.ATIVO)


        If ServidorId > 0 then 
			strSQL.Append(" and Servidor.RH02_ID_SERVIDOR = " & ServidorId)
		End If
		
		If PessoaId> 0 then
			strSQL.Append(" and Servidor.RH01_ID_PESSOA = " & PessoaId)
		End If
		
		If  CargoId > 0 then
			strSQL.Append(" and Servidor.RH16_ID_CARGO = " & CargoId)
		End If
		
		If CategoriaFuncaoId> 0 then
			strSQL.Append(" and Servidor.RH39_ID_CATEGORIA_FUNCIONAL = " & CategoriaFuncaoId)
		End If
		
		If Simbologia > 0 then
			strSQL.Append(" and Servidor.RH35_ID_SIMBOLOGIA = " & Simbologia)
		End If
		
		If OrgaoId > 0  then
			strSQL.Append(" and Servidor.RH04_ID_ORGAO = " & OrgaoId)
		End If
		
		If  TipoVinculoId> 0 then
			strSQL.Append(" and Servidor.RH05_ID_TIPO_VINCULO = " & TipoVinculoId)
		End If
		
		If SituacaoServidorId = 10 Then
		    strSQL.Append(" and Servidor.RH07_ID_SITUACAO_SERVIDOR in (11,10 )")
            Else 
          
            strSQL.Append(" and Servidor.RH07_ID_SITUACAO_SERVIDOR in (1,11,10 )")
        End If
		
		If UsuarioId > 0 then 
			strSQL.Append(" and CA04_ID_USUARIO = " & UsuarioId)
		End If
		
		If Matricula <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_CD_MATRICULA) like '%" & Matricula.toUpper & "%'")
		End If
		
		If DataAdmissao <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_DT_ADMISSAO) like '%" & DataAdmissao.toUpper & "%'")
		End If
		
		If DataPosse <> "" then 
			strSQL.Append(" and upper(vRH02_DT_POSSE) like '%" & DataPosse.toUpper & "%'")
		End If
		
		If DataNomeacao <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_DT_NOMEACAO) like '%" & DataNomeacao.toUpper & "%'")
		End If
		
		If SolicitacaoMatricula <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_DT_SOLICITACAO_MATRIC) like '%" & SolicitacaoMatricula.toUpper & "%'")
		End If
		
		If AnuenioFuncao <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_DT_ANUENIO_FUNCAO) like '%" & AnuenioFuncao.toUpper & "%'")
		End If
		
		If DataTransferencia <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_DT_TRANSFERENCIA) like '%" & DataTransferencia.toUpper & "%'")
		End If
		
		If DataRetorno <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_DT_RETORNO) like '%" & DataRetorno.toUpper & "%'")
		End If
		
		If DataDemissao <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_DT_DEMISSAO) like '%" & DataDemissao.toUpper & "%'")
		End If
		
		If AtoProvimento <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_IN_ATO_PROVIMENTO) like '%" & AtoProvimento.toUpper & "%'")
		End If
		
		If isDate(DataHoraCadastro) then 
			strSQL.Append(" and Servidor.RH02_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
		End If
		
		If isDate(DataHoraUltimaCarga) then 
			strSQL.Append(" and Servidor.RH02_DH_ULTIMA_CARGA = Convert(DateTime, '" & DataHoraUltimaCarga & "', 103)")
		End If
		
		If IsNumeric(QtdHoraSemanalReduzida.Replace(".", "")) then
			strSQL.Append(" and Servidor.RH02_QT_HR_SEMANAL_REDUZIDA = " & QtdHoraSemanalReduzida.Replace(".", "").Replace(",", "."))
		End If
		
		If ReducaoCargaHoraria <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_IN_REDUCAO_CARGA_HORARIA) like '%" & ReducaoCargaHoraria.toUpper & "%'")
		End If
		
		If NumeroPortaria <> "" then 
			strSQL.Append(" and upper(Servidor.RH02_NU_PORTARIA) like '%" & NumeroPortaria.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "Servidor.RH02_ID_SERVIDOR", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder

        strSQL.Append(" select servidor.RH02_ID_SERVIDOR as CODIGO, servidor.RH02_CD_MATRICULA +' - '+ Cargo.RH16_NM_CARGO  as DESCRICAO")
        strSQL.Append(" from RH02_SERVIDOR Servidor")
        strSQL.Append(" inner join RH16_ID_CARGO as Cargo on  cargo.RH16_ID_CARGO = servidor.RH16_ID_CARGO ")
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
		
		strSQL.Append(" select max(RH02_ID_SERVIDOR) from RH02_SERVIDOR")

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
	Public Function Excluir(ByVal ServidorId as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH02_SERVIDOR")
		strSQL.Append(" where RH02_ID_SERVIDOR = " & ServidorId)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return LinhasAfetadas
	End Function

    Public Enum Situacao As Short
        APOSENTADO                  = 4     
        ATIVO                       = 1
        DESLIGADO                   = 2
        FALECIDO                    = 7
        LICENCA_REMUNERADA          = 6
        MATRICULA_SOLICITADA        = 11
        NAO_CONFIRMADO              = 3
        PRE_CADASTRADO              = 10
        SUSPENSO                    = 8
        TRANSFERIDO                 = 5
        
    End Enum

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
