Imports Microsoft.VisualBasic
Imports System.Data

Public Class Habilitacao
    Implements IDisposable

    Private RH74_ID_HABILITACAO As Integer
    Private RH72_ID_HABILIDADE As Integer
    Private RH80_ID_ALOCACAO_CARGA_HORARIA As Integer
    Private CA04_ID_USUARIO As Integer
    Private CA04_ID_USUARIO_ALT As Integer
    Private RH74_QT_HORA_ALOCADA As String
    Private RH74_DH_CADASTRO As String
    Private RH74_DH_DESATIVACAO As String

    Public Property IdHabilitacao() As Integer
        Get
            Return RH74_ID_HABILITACAO
        End Get
        Set(ByVal Value As Integer)
            RH74_ID_HABILITACAO = Value
        End Set
    End Property
    Public Property IdHabilidade() As Integer
        Get
            Return RH72_ID_HABILIDADE
        End Get
        Set(ByVal Value As Integer)
            RH72_ID_HABILIDADE = Value
        End Set
    End Property
    Public Property IdAlocacaoCargaHoraria() As Integer
        Get
            Return RH80_ID_ALOCACAO_CARGA_HORARIA
        End Get
        Set(ByVal Value As Integer)
            RH80_ID_ALOCACAO_CARGA_HORARIA = Value
        End Set
    End Property
    Public Property IdUsuario() As Integer
        Get
            Return CA04_ID_USUARIO
        End Get
        Set(ByVal Value As Integer)
            CA04_ID_USUARIO = Value
        End Set
    End Property
    Public Property IdUsuarioAlteracao() As Integer
        Get
            Return CA04_ID_USUARIO_ALT
        End Get
        Set(ByVal Value As Integer)
            CA04_ID_USUARIO_ALT = Value
        End Set
    End Property
    Public Property QtdHoraAlocada() As String
        Get
            Return RH74_QT_HORA_ALOCADA
        End Get
        Set(ByVal Value As String)
            RH74_QT_HORA_ALOCADA = Value
        End Set
    End Property
    Public Property DataHoraCadastro() As String
        Get
            Return RH74_DH_CADASTRO
        End Get
        Set(ByVal Value As String)
            RH74_DH_CADASTRO = Value
        End Set
    End Property
    Public Property DataHoraDesativacao() As String
        Get
            Return RH74_DH_DESATIVACAO
        End Get
        Set(ByVal Value As String)
            RH74_DH_DESATIVACAO = Value
        End Set
    End Property

    Public Sub New(Optional ByVal IdHabilitacao As Integer = 0)
        If IdHabilitacao > 0 Then
            Obter(IdHabilitacao)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH74_HABILITACAO")
        strSQL.Append(" where RH74_ID_HABILITACAO = " & IdHabilitacao)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH72_ID_HABILIDADE") = ProBanco(RH72_ID_HABILIDADE, eTipoValor.CHAVE)
        dr("RH80_ID_ALOCACAO_CARGA_HORARIA") = ProBanco(RH80_ID_ALOCACAO_CARGA_HORARIA, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO_ALT") = ProBanco(CA04_ID_USUARIO_ALT, eTipoValor.CHAVE)
        dr("RH74_QT_HORA_ALOCADA") = ProBanco(RH74_QT_HORA_ALOCADA, eTipoValor.numero_decimal)
        dr("RH74_DH_CADASTRO") = ProBanco(RH74_DH_CADASTRO, eTipoValor.DATA)
        dr("RH74_DH_DESATIVACAO") = ProBanco(RH74_DH_DESATIVACAO, eTipoValor.DATA)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal IdHabilitacao As String)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH74_HABILITACAO")
        strSQL.Append(" where RH74_ID_HABILITACAO = " & IdHabilitacao)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH74_ID_HABILITACAO = DoBanco(dr("RH74_ID_HABILITACAO"), eTipoValor.CHAVE)
            RH72_ID_HABILIDADE = DoBanco(dr("RH72_ID_HABILIDADE"), eTipoValor.CHAVE)
            RH80_ID_ALOCACAO_CARGA_HORARIA = DoBanco(dr("RH80_ID_ALOCACAO_CARGA_HORARIA"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            CA04_ID_USUARIO_ALT = DoBanco(dr("CA04_ID_USUARIO_ALT"), eTipoValor.CHAVE)
            RH74_QT_HORA_ALOCADA = DoBanco(dr("RH74_QT_HORA_ALOCADA"), eTipoValor.numero_inteiro)
            RH74_DH_CADASTRO = DoBanco(dr("RH74_DH_CADASTRO"), eTipoValor.DATA)
            RH74_DH_DESATIVACAO = DoBanco(dr("RH74_DH_DESATIVACAO"), eTipoValor.DATA)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional IdHabilitacao As Integer = 0, Optional IdHabilidade As Integer = 0 _
                              , Optional IdAlocacaoCargaHoraria As Integer = 0, Optional IdUsuario As Integer = 0, Optional IdUsuarioAlteracao As Integer = 0 _
                              , Optional QtdHoraAlocada As Integer = 0, Optional DataHoraCadastro As String = "", Optional DataHoraDesativacao As String = "" _
                              , Optional ByVal IdServidor As Integer = 0, Optional IdPessoa As Integer = 0, Optional ByVal RegistroAtivo As Boolean = True _
                              , Optional ServidorCargaHoraria As Integer = 0, Optional ByVal IdServidorLotacao As Integer = 0, Optional ByVal Periodo As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select *, ")
        strSQL.Append(" case when RH72_TP_HABILIDADE = 1 then 'COMUM' WHEN RH72_TP_HABILIDADE = 2 THEN 'DESVIO' END  AS TIPO_HABILIDADE, case when isnull(RH74_DH_DESATIVACAO,0)   = ''  then 0 else 1 end as Desativado ")
        strSQL.Append(" from RH74_HABILITACAO HABILITACAO")
        strSQL.Append(" inner join RH72_HABILIDADE HABILIDADE on HABILIDADE.RH72_ID_HABILIDADE = HABILITACAO.RH72_ID_HABILIDADE ")
        strSQL.Append(" INNER JOIN RH02_SERVIDOR SERVIDOR ON SERVIDOR.RH02_ID_SERVIDOR = HABILIDADE.RH02_ID_SERVIDOR ")
        strSQL.Append(" inner join DBDIARIO..DE09_DISCIPLINA DISCIPLINA ON HABILIDADE.DE09_ID_DISCIPLINA = DISCIPLINA.DE09_ID_DISCIPLINA ")
        strSQL.Append(" INNER JOIN RH80_ALOCACAO_CARGA_HORARIA ALOCACAOCARGAHORARIA ON ALOCACAOCARGAHORARIA.RH80_ID_ALOCACAO_CARGA_HORARIA = HABILITACAO.RH80_ID_ALOCACAO_CARGA_HORARIA ")
        strSQL.Append(" INNER join RH14_LOTACAO_SERVIDOR LotacaoServidor on LotacaoServidor.RH14_ID_LOTACAO_SERVIDOR = ALOCACAOCARGAHORARIA.RH14_ID_LOTACAO_SERVIDOR ")
        strSQL.Append(" INNER join RH36_LOTACAO Lotacao on LotacaoServidor.RH36_ID_LOTACAO = Lotacao.RH36_ID_LOTACAO   ")
        strSQL.Append(" inner join dbgeral..TG06_TURNO TURNO ON TURNO.TG06_ID_TURNO = ALOCACAOCARGAHORARIA.TG06_ID_TURNO ")
        strSQL.Append(" inner join RH78_SERVIDOR_CARGA_HORARIA ServidorCargaHoraria ON ServidorCargaHoraria.RH78_ID_SERVIDOR_CARGA_HORARIA = ALOCACAOCARGAHORARIA.RH78_ID_SERVIDOR_CARGA_HORARIA ")
        strSQL.Append(" inner join RH77_CARGA_HORARIA CargaHoraria ON CargaHoraria.RH77_ID_CARGA_HORARIA = ServidorCargaHoraria.RH77_ID_CARGA_HORARIA ")
        strSQL.Append(" inner join RH79_TIPO_CARGA_HORARIA TipoCargaHoraria ON TipoCargaHoraria.RH79_ID_TIPO_CARGA_HORARIA = CargaHoraria.RH79_ID_TIPO_CARGA_HORARIA ")
        strSQL.Append(" left join	RH88_PERIODO			rh88	on	LotacaoServidor.RH88_ID_PERIODO = rh88.RH88_ID_PERIODO ")
        strSQL.Append(" where HABILITACAO.RH74_ID_HABILITACAO is not null and LotacaoServidor.RH14_DT_DESLIGAMENTO is null ")
        strSQL.Append(" and servidor.RH07_ID_SITUACAO_SERVIDOR in (1,10,11) ")


        If RegistroAtivo Then
            strSQL.Append(" and HABILITACAO.RH74_DH_DESATIVACAO is null ")
        End If

        If IdServidorLotacao > 0 Then
            strSQL.Append(" and LotacaoServidor.RH14_ID_LOTACAO_SERVIDOR = " & IdServidorLotacao)
        End If

        If Periodo <> "" Then
            strSQL.Append(" and rh88.RH88_NM_PERIODO = " & Periodo)
        End If


        If IdPessoa > 0 Then
            strSQL.Append(" and SERVIDOR.RH01_ID_PESSOA = " & IdPessoa)
        End If

        If IdServidor > 0 Then
            strSQL.Append(" and SERVIDOR.RH02_ID_SERVIDOR = " & IdServidor)
        End If

        If IdHabilitacao > 0 Then
            strSQL.Append(" and HABILITACAO.RH74_ID_HABILITACAO = " & IdHabilitacao)
        End If

        If IdHabilidade > 0 Then
            strSQL.Append(" and HABILIDADE.RH72_ID_HABILIDADE = " & IdHabilidade)
        End If

        If IdAlocacaoCargaHoraria > 0 Then
            strSQL.Append(" and HABILITACAO.RH80_ID_ALOCACAO_CARGA_HORARIA = " & IdAlocacaoCargaHoraria)
        End If

        If IdUsuario > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & IdUsuario)
        End If

        If IdUsuarioAlteracao > 0 Then
            strSQL.Append(" and HABILITACAO.CA04_ID_USUARIO_ALT = " & IdUsuarioAlteracao)
        End If

        If QtdHoraAlocada > 0 Then
            strSQL.Append(" and RH74_QT_HORA_ALOCADA = " & QtdHoraAlocada)
        End If

        If ServidorCargaHoraria > 0 Then
            strSQL.Append(" and ServidorCargaHoraria.RH78_ID_SERVIDOR_CARGA_HORARIA = " & ServidorCargaHoraria)
        End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" and RH74_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
        End If

        If IsDate(DataHoraDesativacao) Then
            strSQL.Append(" and RH74_DH_DESATIVACAO = Convert(DateTime, '" & DataHoraDesativacao & "', 103)")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH74_ID_HABILITACAO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select RH74_ID_HABILITACAO as CODIGO, RH72_ID_HABILIDADE as DESCRICAO")
        strSQL.Append(" from RH74_HABILITACAO")
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

        strSQL.Append(" select max(RH74_ID_HABILITACAO) from RH74_HABILITACAO")

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
    'Public Function Excluir(ByVal IdHabilitacao As String) As Integer
    '    Dim cnn As New Conexao
    '    Dim strSQL As New StringBuilder
    '    Dim LinhasAfetadas As Integer

    '    strSQL.Append(" delete ")
    '    strSQL.Append(" from RH74_HABILITACAO")
    '    strSQL.Append(" where RH74_ID_HABILITACAO = " & IdHabilitacao)

    '    LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

    '    cnn.FecharBanco()
    '    cnn = Nothing

    '    Return LinhasAfetadas
    'End Function

    Public Function DesabilitarHabilitacao(Optional  ByVal CodigoUsuarioDesabilitacao As Integer = 0, optional _
                                                                   ByVal CodigoAlocacaoCargaHoraria As integer = 0   ) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" Update RH74_HABILITACAO  ")
        strSQL.Append(" set RH74_DH_DESATIVACAO = getdate(),  CA04_ID_USUARIO_ALT =" & CodigoUsuarioDesabilitacao)
        strSQL.Append(" WHERE RH80_ID_ALOCACAO_CARGA_HORARIA= " & CodigoAlocacaoCargaHoraria)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

 Public Function DesabilitarPorHabilidade(ByVal Codigohabilidade As integer, byval CodigoUsuarioAlteracao As integer ) As integer

     Dim cnn As New Conexao
     Dim strSQL As New StringBuilder
     Dim LinhasAfetadas As Integer

     strSQL.Append(" Update RH74_HABILITACAO set RH74_DH_DESATIVACAO = getdate(), CA04_ID_USUARIO_ALT= " &CodigoUsuarioAlteracao )
     strSQL.Append(" where RH72_ID_HABILIDADE=" & Codigohabilidade)


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

