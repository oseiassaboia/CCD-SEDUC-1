Imports Microsoft.VisualBasic
Imports System.Data

Public Class CargaHoraria

    Implements IDisposable

    Private RH77_ID_CARGA_HORARIA As Integer
    Private RH79_ID_TIPO_CARGA_HORARIA As Integer
    Private CA04_ID_USUARIO As Integer
    Private RH77_TP_CONTRATACAO As String
    Private RH77_QT_MIN_HR_AULA As String
    Private RH77_QT_MAX_HR_AULA As String
    Private RH77_QT_MIN_HR_PLANEJAMENTO As String
    Private RH77_QT_MAX_HR_PLANEJAMENTO As String
    Private RH77_DT_INICIO_VIGENCIA As String
    Private RH77_DT_TERMINO_VIGENCIA As String
    Private RH77_IN_PERMITE_REDUCAO As Boolean
    Private RH77_DH_CADASTRO As String

    Public Property Id() As Integer
        Get
            Return RH77_ID_CARGA_HORARIA
        End Get
        Set(ByVal Value As Integer)
            RH77_ID_CARGA_HORARIA = Value
        End Set
    End Property
    Public Property IdTipoCargaHoraria() As Integer
        Get
            Return RH79_ID_TIPO_CARGA_HORARIA
        End Get
        Set(ByVal Value As Integer)
            RH79_ID_TIPO_CARGA_HORARIA = Value
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
    Public Property TipoContratacao() As String
        Get
            Return RH77_TP_CONTRATACAO
        End Get
        Set(ByVal Value As String)
            RH77_TP_CONTRATACAO = Value
        End Set
    End Property
    Public Property QtdMinimaHrAula() As String
        Get
            Return RH77_QT_MIN_HR_AULA
        End Get
        Set(ByVal Value As String)
            RH77_QT_MIN_HR_AULA = Value
        End Set
    End Property
    Public Property QtdMaxHrAula() As String
        Get
            Return RH77_QT_MAX_HR_AULA
        End Get
        Set(ByVal Value As String)
            RH77_QT_MAX_HR_AULA = Value
        End Set
    End Property
    Public Property QtdMinHrPlanejamento() As String
        Get
            Return RH77_QT_MIN_HR_PLANEJAMENTO
        End Get
        Set(ByVal Value As String)
            RH77_QT_MIN_HR_PLANEJAMENTO = Value
        End Set
    End Property
    Public Property QtdMaxHrPlanejamento() As String
        Get
            Return RH77_QT_MAX_HR_PLANEJAMENTO
        End Get
        Set(ByVal Value As String)
            RH77_QT_MAX_HR_PLANEJAMENTO = Value
        End Set
    End Property
    Public Property DataInicioVigencia() As String
        Get
            Return RH77_DT_INICIO_VIGENCIA
        End Get
        Set(ByVal Value As String)
            RH77_DT_INICIO_VIGENCIA = Value
        End Set
    End Property
    Public Property DataTerminoVigencia() As String
        Get
            Return RH77_DT_TERMINO_VIGENCIA
        End Get
        Set(ByVal Value As String)
            RH77_DT_TERMINO_VIGENCIA = Value
        End Set
    End Property
    Public Property PermitiReducao() As Boolean
        Get
            Return RH77_IN_PERMITE_REDUCAO
        End Get
        Set(ByVal Value As Boolean)
            RH77_IN_PERMITE_REDUCAO = Value
        End Set
    End Property
    Public Property DataHoraCadastro() As String
        Get
            Return RH77_DH_CADASTRO
        End Get
        Set(ByVal Value As String)
            RH77_DH_CADASTRO = Value
        End Set
    End Property

    Public Sub New(Optional ByVal Id As Integer = 0)
        If Id > 0 Then
            Obter(Id)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH77_CARGA_HORARIA")
        strSQL.Append(" where RH77_ID_CARGA_HORARIA = " & Id)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH79_ID_TIPO_CARGA_HORARIA") = ProBanco(RH79_ID_TIPO_CARGA_HORARIA, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("RH77_TP_CONTRATACAO") = ProBanco(RH77_TP_CONTRATACAO, eTipoValor.CHAVE)
        dr("RH77_QT_MIN_HR_AULA") = ProBanco(RH77_QT_MIN_HR_AULA, eTipoValor.TEXTO)
        dr("RH77_QT_MAX_HR_AULA") = ProBanco(RH77_QT_MAX_HR_AULA, eTipoValor.TEXTO)
        dr("RH77_QT_MIN_HR_PLANEJAMENTO") = ProBanco(RH77_QT_MIN_HR_PLANEJAMENTO, eTipoValor.TEXTO)
        dr("RH77_QT_MAX_HR_PLANEJAMENTO") = ProBanco(RH77_QT_MAX_HR_PLANEJAMENTO, eTipoValor.TEXTO)
        dr("RH77_DT_INICIO_VIGENCIA") = ProBanco(RH77_DT_INICIO_VIGENCIA, eTipoValor.TEXTO)
        dr("RH77_DT_TERMINO_VIGENCIA") = ProBanco(RH77_DT_TERMINO_VIGENCIA, eTipoValor.TEXTO)
        dr("RH77_IN_PERMITE_REDUCAO") = ProBanco(RH77_IN_PERMITE_REDUCAO, eTipoValor.TEXTO)
        dr("RH77_DH_CADASTRO") = ProBanco(RH77_DH_CADASTRO, eTipoValor.DATA)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal Id As String)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH77_CARGA_HORARIA")
        strSQL.Append(" where RH77_ID_CARGA_HORARIA = " & Id)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH77_ID_CARGA_HORARIA = DoBanco(dr("RH77_ID_CARGA_HORARIA"), eTipoValor.CHAVE)
            RH79_ID_TIPO_CARGA_HORARIA = DoBanco(dr("RH79_ID_TIPO_CARGA_HORARIA"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            RH77_TP_CONTRATACAO = DoBanco(dr("RH77_TP_CONTRATACAO"), eTipoValor.TEXTO)
            RH77_QT_MIN_HR_AULA = DoBanco(dr("RH77_QT_MIN_HR_AULA"), eTipoValor.TEXTO)
            RH77_QT_MAX_HR_AULA = DoBanco(dr("RH77_QT_MAX_HR_AULA"), eTipoValor.TEXTO)
            RH77_QT_MIN_HR_PLANEJAMENTO = DoBanco(dr("RH77_QT_MIN_HR_PLANEJAMENTO"), eTipoValor.TEXTO)
            RH77_QT_MAX_HR_PLANEJAMENTO = DoBanco(dr("RH77_QT_MAX_HR_PLANEJAMENTO"), eTipoValor.TEXTO)
            RH77_DT_INICIO_VIGENCIA = DoBanco(dr("RH77_DT_INICIO_VIGENCIA"), eTipoValor.TEXTO)
            RH77_DT_TERMINO_VIGENCIA = DoBanco(dr("RH77_DT_TERMINO_VIGENCIA"), eTipoValor.TEXTO)
            RH77_IN_PERMITE_REDUCAO = DoBanco(dr("RH77_IN_PERMITE_REDUCAO"), eTipoValor.TEXTO)
            RH77_DH_CADASTRO = DoBanco(dr("RH77_DH_CADASTRO"), eTipoValor.DATA)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Id As Integer = 0, Optional IdTipoCargaHoraria As Integer = 0, Optional IdUsuario As Integer = 0, Optional TipoContratacao As String = "", Optional QtdMinimaHrAula As Integer = 0, Optional QtdMaxHrAula As Integer = 0, Optional QtdMinHrPlanejamento As Integer = 0, Optional QtdMaxHrPlanejamento As Integer = 0, Optional DataInicioVigencia As String = "", Optional DataTerminoVigencia As String = "", Optional PermitiReducao As Boolean = False, Optional DataHoraCadastro As String = "", Optional ByVal TipoCargaHorariaIN As String = "", Optional ByVal CodigoServidor As Integer = 0, Optional ByVal Periodo As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select *, RH79_NM_TIPO_CARGA_HORARIA+' -  Tipo Contratação: '+ case when  RH77_TP_CONTRATACAO = 'E' then 'EFETIVO' WHEN RH77_TP_CONTRATACAO ='C' THEN  'CONTRATO' END + ' Máx Hora/Aula: ' + convert(varchar,RH77_QT_MAX_HR_AULA)+' hr' as DESCRICAO ")
        strSQL.Append(" from RH77_CARGA_HORARIA CargaHoraria")
        strSQL.Append(" inner join RH79_TIPO_CARGA_HORARIA TipoCargaHoraria on TipoCargaHoraria.RH79_ID_TIPO_CARGA_HORARIA = CargaHoraria.RH79_ID_TIPO_CARGA_HORARIA")
        strSQL.Append(" where RH77_ID_CARGA_HORARIA is not null")

        If CodigoServidor > 0 And Periodo > 0 Then

            strSQL.Append(" and  not exists ")
            strSQL.Append(" (select * from RH78_SERVIDOR_CARGA_HORARIA RH78 ")
            strSQL.Append(" left join RH88_PERIODO RH88 ON RH78.RH88_ID_PERIODO = RH88.RH88_ID_PERIODO ")
            strSQL.Append("  WHERE RH78.RH77_ID_CARGA_HORARIA = CargaHoraria.RH77_ID_CARGA_HORARIA ")
            strSQL.Append(" and RH02_ID_SERVIDOR = " & CodigoServidor & " And RH78.RH88_ID_PERIODO =" & Periodo & ")")

        End If

        If Id > 0 Then
            strSQL.Append(" and RH77_ID_CARGA_HORARIA = " & Id)
        End If

        If IdTipoCargaHoraria > 0 Then
            strSQL.Append(" and TipoCargaHoraria.RH79_ID_TIPO_CARGA_HORARIA = " & IdTipoCargaHoraria)
        End If

        If IdUsuario > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & IdUsuario)
        End If

        If TipoContratacao <> "" Then
            strSQL.Append(" and upper(RH77_TP_CONTRATACAO) like '%" & TipoContratacao.ToUpper & "%'")
        End If

        If QtdMinimaHrAula > 0 Then
            strSQL.Append(" and RH77_QT_MIN_HR_AULA = " & QtdMinimaHrAula)
        End If

        If QtdMaxHrAula > 0 Then
            strSQL.Append(" and RH77_QT_MAX_HR_AULA = " & QtdMaxHrAula)
        End If

        If QtdMinHrPlanejamento > 0 Then
            strSQL.Append(" and RH77_QT_MIN_HR_PLANEJAMENTO = " & QtdMinHrPlanejamento)
        End If

        If QtdMaxHrPlanejamento > 0 Then
            strSQL.Append(" and RH77_QT_MAX_HR_PLANEJAMENTO = " & QtdMaxHrPlanejamento)
        End If

        If DataInicioVigencia <> "" Then
            strSQL.Append(" and upper(RH77_DT_INICIO_VIGENCIA) like '%" & DataInicioVigencia.ToUpper & "%'")
        End If

        If DataTerminoVigencia <> "" Then
            strSQL.Append(" and upper(RH77_DT_TERMINO_VIGENCIA) like '%" & DataTerminoVigencia.ToUpper & "%'")
        End If

        If IdTipoCargaHoraria > 0 Then
            strSQL.Append(" and TipoCargaHoraria.RH79_ID_TIPO_CARGA_HORARIA = " & IdTipoCargaHoraria)
        End If

        If TipoCargaHorariaIN <> "" Then
            strSQL.Append(" and TipoCargaHoraria.RH79_ID_TIPO_CARGA_HORARIA in (" & TipoCargaHorariaIN & ")")
        End If

        'If PermitiReducao <> "" Then
        '    strSQL.Append(" and upper(RH77_IN_PERMITE_REDUCAO) like '%" & PermitiReducao.ToUpper & "%'")
        'End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" and RH77_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH77_ID_CARGA_HORARIA", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select RH77_ID_CARGA_HORARIA as CODIGO, ")
        strSQL.Append(" RH79_NM_TIPO_CARGA_HORARIA+' -  Tipo Contratação: '+ case when  RH77_TP_CONTRATACAO = 'E' then 'EFETIVO' WHEN RH77_TP_CONTRATACAO ='C' THEN  'CONTRATO' END + ' Máx Hora/Aula: ' + convert(varchar,RH77_QT_MAX_HR_AULA)+' hr' as DESCRICAO  ")
        strSQL.Append(" from RH77_CARGA_HORARIA CargaHoraria")
        strSQL.Append(" inner join RH79_TIPO_CARGA_HORARIA TipoCargaHoraria on TipoCargaHoraria.RH79_ID_TIPO_CARGA_HORARIA = CargaHoraria.RH79_ID_TIPO_CARGA_HORARIA")
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

        strSQL.Append(" select max(RH77_ID_CARGA_HORARIA) from RH77_CARGA_HORARIA")

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
    Public Function Excluir(ByVal Id As String) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from RH77_CARGA_HORARIA")
        strSQL.Append(" where RH77_ID_CARGA_HORARIA = " & Id)

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
'*                                 26/02/2019                                 *
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

