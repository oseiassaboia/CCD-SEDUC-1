Imports Microsoft.VisualBasic
Imports System.Data
Public Class Predio

    Implements IDisposable
    Private TG59_ID_PREDIO As Integer
    Private TG56_ID_LOGRADOURO_CEP As Integer
    Private TG57_ID_TIPO_PREDIO As String
    Private TG58_ID_TIPO_POSSE As String
    Private TG59_NU_PREDIO As String
    Private TG59_NM_PREDIO As String
    Private TG59_DS_COMPLEMENTO As String
    Private TG59_NU_LATITUDE As String
    Private TG59_NU_LONGITUDE As String
    Private TG59_NU_CONTRATO_ENERGIA As String

    Public Property IdPredio() As Integer
        Get
            Return TG59_ID_PREDIO
        End Get
        Set(ByVal Value As Integer)
            TG59_ID_PREDIO = Value
        End Set
    End Property
    Public Property IdLogradouro() As Integer
        Get
            Return TG56_ID_LOGRADOURO_CEP
        End Get
        Set(ByVal Value As Integer)
            TG56_ID_LOGRADOURO_CEP = Value
        End Set
    End Property
    Public Property IdTipoPredio() As String
        Get
            Return TG57_ID_TIPO_PREDIO
        End Get
        Set(ByVal Value As String)
            TG57_ID_TIPO_PREDIO = Value
        End Set
    End Property
    Public Property IdTipoPosse() As String
        Get
            Return TG58_ID_TIPO_POSSE
        End Get
        Set(ByVal Value As String)
            TG58_ID_TIPO_POSSE = Value
        End Set
    End Property
    Public Property NumeroPredio() As String
        Get
            Return TG59_NU_PREDIO
        End Get
        Set(ByVal Value As String)
            TG59_NU_PREDIO = Value
        End Set
    End Property
    Public Property NomePredio() As String
        Get
            Return TG59_NM_PREDIO
        End Get
        Set(ByVal Value As String)
            TG59_NM_PREDIO = Value
        End Set
    End Property
    Public Property Complemento() As String
        Get
            Return TG59_DS_COMPLEMENTO
        End Get
        Set(ByVal Value As String)
            TG59_DS_COMPLEMENTO = Value
        End Set
    End Property
    Public Property NuLatitude() As String
        Get
            Return TG59_NU_LATITUDE
        End Get
        Set(ByVal Value As String)
            TG59_NU_LATITUDE = Value
        End Set
    End Property
    Public Property Nulongitude() As String
        Get
            Return TG59_NU_LONGITUDE
        End Get
        Set(ByVal Value As String)
            TG59_NU_LONGITUDE = Value
        End Set
    End Property
    Public Property NuContratoEnerdia() As String
        Get
            Return TG59_NU_CONTRATO_ENERGIA
        End Get
        Set(ByVal Value As String)
            TG59_NU_CONTRATO_ENERGIA = Value
        End Set
    End Property

    Public Sub New(Optional ByVal IdPredio As Integer = 0)
        If IdPredio > 0 Then
            Obter(IdPredio)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New ConexaoGeral
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from TG59_PREDIO")
        strSQL.Append(" where TG59_ID_PREDIO = " & IdPredio)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("TG56_ID_LOGRADOURO_CEP") = ProBanco(TG56_ID_LOGRADOURO_CEP, eTipoValor.CHAVE)
        dr("TG57_ID_TIPO_PREDIO") = ProBanco(TG57_ID_TIPO_PREDIO, eTipoValor.NUMERO_DECIMAL)
        dr("TG58_ID_TIPO_POSSE") = ProBanco(TG58_ID_TIPO_POSSE, eTipoValor.NUMERO_DECIMAL)
        dr("TG59_NU_PREDIO") = ProBanco(TG59_NU_PREDIO, eTipoValor.TEXTO)
        dr("TG59_NM_PREDIO") = ProBanco(TG59_NM_PREDIO, eTipoValor.TEXTO)
        dr("TG59_DS_COMPLEMENTO") = ProBanco(TG59_DS_COMPLEMENTO, eTipoValor.TEXTO)
        dr("TG59_NU_LATITUDE") = ProBanco(TG59_NU_LATITUDE, eTipoValor.TEXTO)
        dr("TG59_NU_LONGITUDE") = ProBanco(TG59_NU_LONGITUDE, eTipoValor.TEXTO)
        dr("TG59_NU_CONTRATO_ENERGIA") = ProBanco(TG59_NU_CONTRATO_ENERGIA, eTipoValor.TEXTO)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal IdPredio As String)
        Dim cnn As New ConexaoGeral
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from TG59_PREDIO")
        strSQL.Append(" where TG59_ID_PREDIO = " & IdPredio)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG59_ID_PREDIO = DoBanco(dr("TG59_ID_PREDIO"), eTipoValor.CHAVE)
            TG56_ID_LOGRADOURO_CEP = DoBanco(dr("TG56_ID_LOGRADOURO_CEP"), eTipoValor.CHAVE)
            TG57_ID_TIPO_PREDIO = DoBanco(dr("TG57_ID_TIPO_PREDIO"), eTipoValor.NUMERO_DECIMAL)
            TG58_ID_TIPO_POSSE = DoBanco(dr("TG58_ID_TIPO_POSSE"), eTipoValor.NUMERO_DECIMAL)
            TG59_NU_PREDIO = DoBanco(dr("TG59_NU_PREDIO"), eTipoValor.TEXTO)
            TG59_NM_PREDIO = DoBanco(dr("TG59_NM_PREDIO"), eTipoValor.TEXTO)
            TG59_DS_COMPLEMENTO = DoBanco(dr("TG59_DS_COMPLEMENTO"), eTipoValor.TEXTO)
            TG59_NU_LATITUDE = DoBanco(dr("TG59_NU_LATITUDE"), eTipoValor.TEXTO)
            TG59_NU_LONGITUDE = DoBanco(dr("TG59_NU_LONGITUDE"), eTipoValor.TEXTO)
            TG59_NU_CONTRATO_ENERGIA = DoBanco(dr("TG59_NU_CONTRATO_ENERGIA"), eTipoValor.TEXTO)
        End If


        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional IdPredio As Integer = 0, Optional IdLogradouro As Integer = 0, Optional IdTipoPredio As String = "", Optional IdTipoPosse As String = "", Optional NumeroPredio As String = "", Optional NomePredio As String = "", Optional Complemento As String = "", Optional NuLatitude As String = "", Optional Nulongitude As String = "", Optional NuContratoEnerdia As String = "") As DataTable
        Dim cnn As New ConexaoGeral
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from TG59_PREDIO")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where TG59_ID_PREDIO is not null")

        If IdPredio > 0 Then
            strSQL.Append(" and TG59_ID_PREDIO = " & IdPredio)
        End If

        If IdLogradouro > 0 Then
            strSQL.Append(" and TG56_ID_LOGRADOURO_CEP = " & IdLogradouro)
        End If

        If IsNumeric(IdTipoPredio.Replace(".", "")) Then
            strSQL.Append(" and TG57_ID_TIPO_PREDIO = " & IdTipoPredio.Replace(".", "").Replace(",", "."))
        End If

        If IsNumeric(IdTipoPosse.Replace(".", "")) Then
            strSQL.Append(" and TG58_ID_TIPO_POSSE = " & IdTipoPosse.Replace(".", "").Replace(",", "."))
        End If

        If NumeroPredio <> "" Then
            strSQL.Append(" and upper(TG59_NU_PREDIO) like '%" & NumeroPredio.ToUpper & "%'")
        End If

        If NomePredio <> "" Then
            strSQL.Append(" and upper(TG59_NM_PREDIO) like '%" & NomePredio.ToUpper & "%'")
        End If

        If Complemento <> "" Then
            strSQL.Append(" and upper(TG59_DS_COMPLEMENTO) like '%" & Complemento.ToUpper & "%'")
        End If

        If NuLatitude <> "" Then
            strSQL.Append(" and upper(TG59_NU_LATITUDE) like '%" & NuLatitude.ToUpper & "%'")
        End If

        If Nulongitude <> "" Then
            strSQL.Append(" and upper(TG59_NU_LONGITUDE) like '%" & Nulongitude.ToUpper & "%'")
        End If

        If NuContratoEnerdia <> "" Then
            strSQL.Append(" and upper(TG59_NU_CONTRATO_ENERGIA) like '%" & NuContratoEnerdia.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "TG59_ID_PREDIO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New ConexaoGeral
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select TG59_ID_PREDIO as CODIGO, TG57_ID_TIPO_PREDIO as DESCRICAO")
        strSQL.Append(" from TG59_PREDIO")
        strSQL.Append(" order by 2 ")

        dt = cnn.AbrirDataTable(strSQL.ToString)


        cnn = Nothing

        Return dt
    End Function

    Public Function ObterUltimo() As Integer
        Dim cnn As New ConexaoGeral
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(TG59_ID_PREDIO) from TG59_PREDIO")

        With cnn.AbrirDataTable(strSQL.ToString)
            If Not IsDBNull(.Rows(0)(0)) Then
                CodigoUltimo = .Rows(0)(0)
            Else
                CodigoUltimo = 0
            End If
        End With


        cnn = Nothing

        Return CodigoUltimo

    End Function
    Public Function Excluir(ByVal IdPredio As String) As Integer
        Dim cnn As New ConexaoGeral
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from TG59_PREDIO")
        strSQL.Append(" where TG59_ID_PREDIO = " & IdPredio)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)


        cnn = Nothing

        Return LinhasAfetadas
    End Function

    Public Function PesquisarEnderecoPredio(ByVal Predio As Integer) As DataTable
        Dim cnn As New ConexaoGeral
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select TG59.TG59_ID_PREDIO  ")
        strSQL.Append("   , TG56.TG56_ID_LOGRADOURO_CEP, TG55.TG55_NU_CEP, TG54.TG54_NM_LOGRADOURO, TG16.TG16_NM_TIPO_LOGRADOURO ")
        strSQL.Append("	  , TG04.TG04_NM_BAIRRO, TG03.TG03_NM_MUNICIPIO, TG02.TG02_SG_UF ")
        strSQL.Append(" From TG59_PREDIO As TG59 ")
        strSQL.Append(" Left Join DBGERAL.DBO.TG56_LOGRADOURO_CEP  AS TG56 ON TG56.TG56_ID_LOGRADOURO_CEP = TG59.TG56_ID_LOGRADOURO_CEP ")
        strSQL.Append(" Left Join DBGERAL.DBO.TG55_CEP AS TG55 ON TG55.TG55_ID_CEP = TG56.TG55_ID_CEP ")
        strSQL.Append(" Left Join DBGERAL.DBO.TG54_LOGRADOURO AS TG54 ON TG54.TG54_ID_LOGRADOURO = TG56.TG54_ID_LOGRADOURO ")
        strSQL.Append(" Left Join DBGERAL.DBO.TG16_TIPO_LOGRADOURO AS TG16 ON TG16.TG16_ID_TIPO_LOGRADOURO = TG54.TG16_ID_TIPO_LOGRADOURO ")
        strSQL.Append(" Left Join DBGERAL.DBO.TG04_BAIRRO AS TG04 ON TG04.TG04_ID_BAIRRO = TG54.TG04_ID_BAIRRO ")
        strSQL.Append(" Left Join DBGERAL.DBO.TG03_MUNICIPIO AS TG03 ON TG03.TG03_ID_MUNICIPIO = TG04.TG03_ID_MUNICIPIO ")
        strSQL.Append(" Left Join DBGERAL.DBO.TG02_UF AS TG02 ON TG02.TG02_ID_UF = TG03.TG02_ID_UF ")
        strSQL.Append(" where TG59.TG59_ID_PREDIO =  " & Predio)

        Return cnn.AbrirDataTable(strSQL.ToString)
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
'*                                 20/09/2019                                 *
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

