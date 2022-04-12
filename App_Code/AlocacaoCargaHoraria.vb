Imports Microsoft.VisualBasic
Imports System.Data

Public Class AlocacaoCargaHoraria

    Implements  IDisposable

	Private RH80_ID_ALOCACAO_CARGA_HORARIA as Integer
	Private RH14_ID_LOTACAO_SERVIDOR as Integer
	Private RH78_ID_SERVIDOR_CARGA_HORARIA as Integer
	Private TG06_ID_TURNO as Integer
	Private TG62_ID_POLO as Integer
	Private CA04_ID_USUARIO as Integer
	Private CA04_ID_USUARIO_RECEBEU as Integer
	Private RH80_DH_RECEBIMENTO as String
	Private RH80_QT_HORA_ALOCADA as String
	Private RH80_DH_CADASTRO as String
	Private RH80_DH_DESATIVACAO as String
	Private RH80_DS_MOTIVO as String

	Public Property IdAlocacaoCargaHoraria() as Integer
		Get
			Return RH80_ID_ALOCACAO_CARGA_HORARIA
		End Get
		Set(ByVal Value As Integer)
			RH80_ID_ALOCACAO_CARGA_HORARIA = Value
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
	Public Property IdServidorCargaHoraria() as Integer
		Get
			Return RH78_ID_SERVIDOR_CARGA_HORARIA
		End Get
		Set(ByVal Value As Integer)
			RH78_ID_SERVIDOR_CARGA_HORARIA = Value
		End Set
	End Property
	Public Property IdTurno() as Integer
		Get
			Return TG06_ID_TURNO
		End Get
		Set(ByVal Value As Integer)
			TG06_ID_TURNO = Value
		End Set
	End Property
	Public Property IdPolo() as Integer
		Get
			Return TG62_ID_POLO
		End Get
		Set(ByVal Value As Integer)
			TG62_ID_POLO = Value
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
	Public Property IdusuarioRecebeu() as Integer
		Get
			Return CA04_ID_USUARIO_RECEBEU
		End Get
		Set(ByVal Value As Integer)
			CA04_ID_USUARIO_RECEBEU = Value
		End Set
	End Property
	Public Property DataHoraRecebimento() as String
		Get
			Return RH80_DH_RECEBIMENTO
		End Get
		Set(ByVal Value As String)
			RH80_DH_RECEBIMENTO = Value
		End Set
	End Property
	Public Property QtdHotaAlocada() as String
		Get
			Return RH80_QT_HORA_ALOCADA
		End Get
		Set(ByVal Value As String)
			RH80_QT_HORA_ALOCADA = Value
		End Set
	End Property
	Public Property DataHoraCadastro() as String
		Get
			Return RH80_DH_CADASTRO
		End Get
		Set(ByVal Value As String)
			RH80_DH_CADASTRO = Value
		End Set
	End Property
	Public Property DataHoraDesativado() as String
		Get
			Return RH80_DH_DESATIVACAO
		End Get
		Set(ByVal Value As String)
			RH80_DH_DESATIVACAO = Value
		End Set
	End Property
	Public Property DescricaoMotivo() as String
		Get
			Return RH80_DS_MOTIVO
		End Get
		Set(ByVal Value As String)
			RH80_DS_MOTIVO = Value
		End Set
	End Property

	Public Sub New(Optional ByVal IdAlocacaoCargaHoraria as Integer = 0)
		If IdAlocacaoCargaHoraria > 0 Then
			Obter(IdAlocacaoCargaHoraria)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH80_ALOCACAO_CARGA_HORARIA")
		strSQL.Append(" where RH80_ID_ALOCACAO_CARGA_HORARIA = " & IdAlocacaoCargaHoraria)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH14_ID_LOTACAO_SERVIDOR") = ProBanco(RH14_ID_LOTACAO_SERVIDOR, eTipoValor.CHAVE)
		dr("RH78_ID_SERVIDOR_CARGA_HORARIA") = ProBanco(RH78_ID_SERVIDOR_CARGA_HORARIA, eTipoValor.CHAVE)
		dr("TG06_ID_TURNO") = ProBanco(TG06_ID_TURNO, eTipoValor.CHAVE)
		dr("TG62_ID_POLO") = ProBanco(TG62_ID_POLO, eTipoValor.CHAVE)
		dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
		dr("CA04_ID_USUARIO_RECEBEU") = ProBanco(CA04_ID_USUARIO_RECEBEU, eTipoValor.CHAVE)
		dr("RH80_DH_RECEBIMENTO") = ProBanco(RH80_DH_RECEBIMENTO, eTipoValor.DATA)
		dr("RH80_QT_HORA_ALOCADA") = ProBanco(RH80_QT_HORA_ALOCADA, eTipoValor.NUMERO_INTEIRO)
		dr("RH80_DH_CADASTRO") = ProBanco(RH80_DH_CADASTRO, eTipoValor.DATA)
		dr("RH80_DH_DESATIVACAO") = ProBanco(RH80_DH_DESATIVACAO, eTipoValor.DATA)
		dr("RH80_DS_MOTIVO") = ProBanco(RH80_DS_MOTIVO, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal IdAlocacaoCargaHoraria as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH80_ALOCACAO_CARGA_HORARIA")
		strSQL.Append(" where RH80_ID_ALOCACAO_CARGA_HORARIA = " & IdAlocacaoCargaHoraria)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH80_ID_ALOCACAO_CARGA_HORARIA = DoBanco(dr("RH80_ID_ALOCACAO_CARGA_HORARIA"), eTipoValor.CHAVE)
			RH14_ID_LOTACAO_SERVIDOR = DoBanco(dr("RH14_ID_LOTACAO_SERVIDOR"), eTipoValor.CHAVE)
			RH78_ID_SERVIDOR_CARGA_HORARIA = DoBanco(dr("RH78_ID_SERVIDOR_CARGA_HORARIA"), eTipoValor.CHAVE)
			TG06_ID_TURNO = DoBanco(dr("TG06_ID_TURNO"), eTipoValor.CHAVE)
			TG62_ID_POLO = DoBanco(dr("TG62_ID_POLO"), eTipoValor.CHAVE)
			CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
			CA04_ID_USUARIO_RECEBEU = DoBanco(dr("CA04_ID_USUARIO_RECEBEU"), eTipoValor.CHAVE)
			RH80_DH_RECEBIMENTO = DoBanco(dr("RH80_DH_RECEBIMENTO"), eTipoValor.DATA)
			RH80_QT_HORA_ALOCADA = DoBanco(dr("RH80_QT_HORA_ALOCADA"), eTipoValor.NUMERO_INTEIRO)
			RH80_DH_CADASTRO = DoBanco(dr("RH80_DH_CADASTRO"), eTipoValor.DATA)
			RH80_DH_DESATIVACAO = DoBanco(dr("RH80_DH_DESATIVACAO"), eTipoValor.DATA)
			RH80_DS_MOTIVO = DoBanco(dr("RH80_DS_MOTIVO"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort As String = "", Optional IdAlocacaoCargaHoraria As Integer = 0, Optional IdLotacaoServidor As Integer = 0, Optional IdServidorCargaHoraria As Integer = 0 _
							  , Optional IdTurno As Integer = 0, Optional IdPolo As Integer = 0, Optional IdUsuario As Integer = 0, Optional IdusuarioRecebeu As Integer = 0, Optional DataHoraRecebimento As String = "" _
							  , Optional QtdHotaAlocada As String = "", Optional DataHoraCadastro As String = "", Optional DataHoraDesativado As String = "", Optional DescricaoMotivo As String = "" _
							  , Optional CodigoPessoa As Integer = 0, Optional ByVal RegistroAtivo As Boolean = True, Optional ByVal CodigoServidor As Integer = 0 _
							  , Optional ByVal Periodo As String = "", Optional ByVal IdLotacao As Integer = 0, Optional ByVal IdPeriodo As Integer = 0) As DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder

		strSQL.Append(" select *, ")
		strSQL.Append(" 'LOTAÇÃO: '+ RH36_NM_LOTACAO + ' - TURNO: '+ TG06_NM_TURNO + ' - HORA/AULA ALOCADA: ' + CONVERT(VARCHAR,RH80_QT_HORA_ALOCADA) AS DESCRICAO,case when isnull(RH80_DH_DESATIVACAO,0)   = ''  then 0 else 1 end as Desativado ")
		strSQL.Append(" from RH80_ALOCACAO_CARGA_HORARIA ALOCACAOCARGAHORARIA")
		strSQL.Append(" inner join RH78_SERVIDOR_CARGA_HORARIA ServidorCargaHoraria on ServidorCargaHoraria.RH78_ID_SERVIDOR_CARGA_HORARIA = ALOCACAOCARGAHORARIA.RH78_ID_SERVIDOR_CARGA_HORARIA ")
		strSQL.Append(" inner join RH77_CARGA_HORARIA CARGAHORARIA ON CARGAHORARIA.RH77_ID_CARGA_HORARIA = ServidorCargaHoraria.RH77_ID_CARGA_HORARIA  ")
		strSQL.Append(" inner join RH79_TIPO_CARGA_HORARIA TIPOCARGA ON TIPOCARGA.RH79_ID_TIPO_CARGA_HORARIA = CARGAHORARIA.RH79_ID_TIPO_CARGA_HORARIA  ")
		strSQL.Append(" inner join RH14_LOTACAO_SERVIDOR LOTACAOSERVIDOR on LOTACAOSERVIDOR.RH14_ID_LOTACAO_SERVIDOR = ALOCACAOCARGAHORARIA.RH14_ID_LOTACAO_SERVIDOR ")
		strSQL.Append(" inner join RH36_LOTACAO LOTACAO on LOTACAO.RH36_ID_LOTACAO = LOTACAOSERVIDOR.RH36_ID_LOTACAO ")
		strSQL.Append(" inner join RH02_SERVIDOR SERVIDOR  on SERVIDOR.RH02_ID_SERVIDOR = LOTACAOSERVIDOR.RH02_ID_SERVIDOR ")
		strSQL.Append(" inner join DBGERAL..tg06_turno turno on ALOCACAOCARGAHORARIA.TG06_ID_TURNO = turno.TG06_ID_TURNO ")
		strSQL.Append(" left join	RH88_PERIODO			rh88	on	LOTACAOSERVIDOR.RH88_ID_PERIODO = rh88.RH88_ID_PERIODO ")
		strSQL.Append(" where ALOCACAOCARGAHORARIA.RH80_ID_ALOCACAO_CARGA_HORARIA is not null")

		If RegistroAtivo Then
			strSQL.Append(" and RH80_DH_DESATIVACAO is null ")
		End If

		If IdLotacao > 0 Then
			strSQL.Append(" and LOTACAOSERVIDOR.RH36_ID_LOTACAO =" & IdLotacao)
		End If

		If CodigoPessoa > 0 Then
			strSQL.Append(" and SERVIDOR.RH01_ID_PESSOA = " & CodigoPessoa)
		End If

		If CodigoServidor > 0 Then
			strSQL.Append(" and SERVIDOR.RH02_ID_SERVIDOR = " & CodigoServidor)
		End If

		If IdAlocacaoCargaHoraria > 0 Then
			strSQL.Append(" and ALOCACAOCARGAHORARIA.RH80_ID_ALOCACAO_CARGA_HORARIA = " & IdAlocacaoCargaHoraria)
		End If

		If IdLotacaoServidor > 0 Then
			strSQL.Append(" and LOTACAOSERVIDOR.RH14_ID_LOTACAO_SERVIDOR = " & IdLotacaoServidor)
		End If

		If IdServidorCargaHoraria > 0 Then
			strSQL.Append(" and ServidorCargaHoraria.RH78_ID_SERVIDOR_CARGA_HORARIA = " & IdServidorCargaHoraria)
		End If

		If IdTurno > 0 Then
			strSQL.Append(" and ALOCACAOCARGAHORARIA.TG06_ID_TURNO = " & IdTurno)
		End If

		If IdPolo > 0 Then
			strSQL.Append(" and TG62_ID_POLO = " & IdPolo)
		End If

		If IdUsuario > 0 Then
			strSQL.Append(" and CA04_ID_USUARIO = " & IdUsuario)
		End If

		If IdusuarioRecebeu > 0 Then
			strSQL.Append(" and CA04_ID_USUARIO_RECEBEU = " & IdusuarioRecebeu)
		End If


		If Periodo <> "" Then
			strSQL.Append(" and rh88.RH88_NM_PERIODO = " & Periodo)
		End If

		If IdPeriodo > 0 Then
			strSQL.Append(" and rh88.RH88_ID_PERIODO = " & IdPeriodo)
		End If

		If IsDate(DataHoraRecebimento) Then
			strSQL.Append(" and RH80_DH_RECEBIMENTO = Convert(DateTime, '" & DataHoraRecebimento & "', 103)")
		End If

		If IsNumeric(QtdHotaAlocada.Replace(".", "")) Then
			strSQL.Append(" and RH80_QT_HORA_ALOCADA = " & QtdHotaAlocada.Replace(".", "").Replace(",", "."))
		End If

		If IsDate(DataHoraCadastro) Then
			strSQL.Append(" and RH80_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
		End If

		If IsDate(DataHoraDesativado) Then
			strSQL.Append(" and RH80_DH_DESATIVACAO = Convert(DateTime, '" & DataHoraDesativado & "', 103)")
		End If

		If DescricaoMotivo <> "" Then
			strSQL.Append(" and upper(RH80_DS_MOTIVO) like '%" & DescricaoMotivo.ToUpper & "%'")
		End If

		strSQL.Append(" Order By " & IIf(Sort = "", "ALOCACAOCARGAHORARIA.RH80_ID_ALOCACAO_CARGA_HORARIA", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH80_ID_ALOCACAO_CARGA_HORARIA as CODIGO, RH14_ID_LOTACAO_SERVIDOR as DESCRICAO")
		strSQL.Append(" from RH80_ALOCACAO_CARGA_HORARIA")
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
		
		strSQL.Append(" select max(RH80_ID_ALOCACAO_CARGA_HORARIA) from RH80_ALOCACAO_CARGA_HORARIA")

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
    'Public Function Excluir(ByVal IdAlocacaoCargaHoraria As String) As Integer
    '    Dim cnn As New Conexao
    '    Dim strSQL As New StringBuilder
    '    Dim LinhasAfetadas As Integer

    '    strSQL.Append(" delete ")
    '    strSQL.Append(" from RH80_ALOCACAO_CARGA_HORARIA")
    '    strSQL.Append(" where RH80_ID_ALOCACAO_CARGA_HORARIA = " & IdAlocacaoCargaHoraria)

    '    LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

    '    cnn.FecharBanco()
    '    cnn = Nothing

    '    Return LinhasAfetadas
    'End Function


    Public Function DesabilitarAlocacaoPorCargaHoraria(ByVal CodigoServidorCargaHoraria As Integer, ByVal CodigoUsuarioDesabilitacao As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" Update RH80_ALOCACAO_CARGA_HORARIA ")
        strSQL.Append(" set RH80_DH_DESATIVACAO = getdate(), RH80_DS_MOTIVO = 'DESATIVACAO AUTOMATICA', CA04_ID_USUARIO_RECEBEU =" & CodigoUsuarioDesabilitacao)
        strSQL.Append(" where RH78_ID_SERVIDOR_CARGA_HORARIA = " & CodigoServidorCargaHoraria)

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
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

'******************************************************************************
'*                                 07/03/2019                                 *
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

