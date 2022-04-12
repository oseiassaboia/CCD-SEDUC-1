Imports Microsoft.VisualBasic
Imports System.Data

Public Class Endereco

    Implements IDisposable
    Private RH69_ID_ENDERECO As Integer
    Private RH01_ID_PESSOA As String
    Private TG57_ID_LOGRADOURO_CEP As Integer
    Private RH69_NU_PREDIO As String
    Private RH69_DS_COMPLEMENTO As String
    Private RH69_TP_ZONA_RESIDENCIA As Integer

    Public Property EnderecoId() As Integer
        Get
            Return RH69_ID_ENDERECO
        End Get
        Set(ByVal Value As Integer)
            RH69_ID_ENDERECO = Value
        End Set
    End Property
    Public Property PessoaId() As String
        Get
            Return RH01_ID_PESSOA
        End Get
        Set(ByVal Value As String)
            RH01_ID_PESSOA = Value
        End Set
    End Property
    Public Property LogradouroId() As Integer
        Get
            Return TG57_ID_LOGRADOURO_CEP
        End Get
        Set(ByVal Value As Integer)
            TG57_ID_LOGRADOURO_CEP = Value
        End Set
    End Property
    Public Property PredioNumero() As String
        Get
            Return RH69_NU_PREDIO
        End Get
        Set(ByVal Value As String)
            RH69_NU_PREDIO = Value
        End Set
    End Property
    Public Property Complemento() As String
        Get
            Return RH69_DS_COMPLEMENTO
        End Get
        Set(ByVal Value As String)
            RH69_DS_COMPLEMENTO = Value
        End Set
    End Property
    Public Property TipoZonaResidencia() As Integer
        Get
            Return RH69_TP_ZONA_RESIDENCIA
        End Get
        Set(value As Integer)
            RH69_TP_ZONA_RESIDENCIA = value
        End Set
    End Property

    Public Sub New(Optional ByVal EnderecoId As Integer = 0)
        If EnderecoId > 0 Then
            Obter(EnderecoId)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH69_ENDERECO")
        strSQL.Append(" where RH69_ID_ENDERECO = " & EnderecoId)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH01_ID_PESSOA") = ProBanco(RH01_ID_PESSOA, eTipoValor.CHAVE)
        dr("TG57_ID_LOGRADOURO_CEP") = ProBanco(TG57_ID_LOGRADOURO_CEP, eTipoValor.CHAVE)
        dr("RH69_NU_PREDIO") = ProBanco(RH69_NU_PREDIO, eTipoValor.TEXTO)
        dr("RH69_DS_COMPLEMENTO") = ProBanco(RH69_DS_COMPLEMENTO, eTipoValor.TEXTO)
        dr("RH69_TP_ZONA_RESIDENCIA") = ProBanco(RH69_TP_ZONA_RESIDENCIA, eTipoValor.NUMERO_INTEIRO)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal Codigo As String)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH69_ENDERECO")
        strSQL.Append(" where RH69_ID_ENDERECO = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH69_ID_ENDERECO = DoBanco(dr("RH69_ID_ENDERECO"), eTipoValor.CHAVE)
            RH01_ID_PESSOA = DoBanco(dr("RH01_ID_PESSOA"), eTipoValor.CHAVE)
            TG57_ID_LOGRADOURO_CEP = DoBanco(dr("TG57_ID_LOGRADOURO_CEP"), eTipoValor.CHAVE)
            RH69_NU_PREDIO = DoBanco(dr("RH69_NU_PREDIO"), eTipoValor.TEXTO)
            RH69_DS_COMPLEMENTO = DoBanco(dr("RH69_DS_COMPLEMENTO"), eTipoValor.TEXTO)
            RH69_TP_ZONA_RESIDENCIA = DoBanco(dr("RH69_TP_ZONA_RESIDENCIA"), eTipoValor.NUMERO_INTEIRO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function SalvarEnderecoXML(ByVal Cep As String, ByVal TipoLogradouro As String, ByVal Logradouro As String, ByVal LogradouroCorreios As Boolean, ByVal Bairro As String, ByVal BairroCorreios As Boolean, ByVal Estado As String, ByVal Municipio As String) As DataTable
        Dim cnn As New ConexaoGeral
        Dim dt As DataTable
        Dim strSQL As New StringBuilder


        strSQL.Append(" Declare @P_NU_CEP                  VARCHAR(8) ")
        strSQL.Append(" Declare @P_NM_TIPO_LOGRADOURO      VARCHAR(20) ")
        strSQL.Append(" Declare @P_NM_LOGRADOURO           VARCHAR(MAX) ")
        strSQL.Append(" Declare @P_IN_LOGRADOURO_CORREIOS  BIT         ")
        strSQL.Append(" Declare @P_NM_BAIRRO               VARCHAR(MAX)")
        strSQL.Append(" Declare @P_IN_BAIRRO_CORREIOS      BIT         ")
        strSQL.Append(" Declare @P_NM_MUNICIPIO            VARCHAR(MAX) ")
        strSQL.Append(" Declare @P_SG_UF                   Char(2) ")
        strSQL.Append(" Declare @P_ID_LOGRADOURO_CEP       INT  ")

        strSQL.Append(" Set @P_NU_CEP = '" & Cep & "'")
        strSQL.Append(" Set @P_NM_TIPO_LOGRADOURO = '" & TipoLogradouro & "'")
        strSQL.Append(" Set @P_NM_LOGRADOURO = '" & Logradouro & "'")
        strSQL.Append(" Set @P_IN_LOGRADOURO_CORREIOS = " & IIf(LogradouroCorreios = False, 0, 1))
        strSQL.Append(" Set @P_NM_BAIRRO = '" & Replace(Bairro, "'", "") & "'")
        strSQL.Append(" Set @P_IN_BAIRRO_CORREIOS = " & IIf(BairroCorreios = False, 0, 1))
        strSQL.Append(" Set @P_NM_MUNICIPIO = '" & Municipio & "'")
        strSQL.Append(" Set @P_SG_UF = '" & Estado & "'")


        strSQL.Append(" exec SP_ENDERECO  @P_NU_CEP, @P_NM_TIPO_LOGRADOURO, @P_NM_LOGRADOURO, @P_IN_LOGRADOURO_CORREIOS, @P_NM_BAIRRO, @P_IN_BAIRRO_CORREIOS, @P_NM_MUNICIPIO, @P_SG_UF, @P_ID_LOGRADOURO_CEP OUT ")
        strSQL.Append(" Select @P_ID_LOGRADOURO_CEP ")

        dt = cnn.AbrirDataTable(strSQL.ToString)


        cnn = Nothing

        Return dt
    End Function

    Public Sub ObterLogradouroCep(ByVal Codigo As String)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH69_ENDERECO")
        strSQL.Append(" where TG57_ID_LOGRADOURO_CEP = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH69_ID_ENDERECO = DoBanco(dr("RH69_ID_ENDERECO"), eTipoValor.CHAVE)
            RH01_ID_PESSOA = DoBanco(dr("RH01_ID_PESSOA"), eTipoValor.CHAVE)
            TG57_ID_LOGRADOURO_CEP = DoBanco(dr("TG57_ID_LOGRADOURO_CEP"), eTipoValor.CHAVE)
            RH69_NU_PREDIO = DoBanco(dr("RH69_NU_PREDIO"), eTipoValor.TEXTO)
            RH69_DS_COMPLEMENTO = DoBanco(dr("RH69_DS_COMPLEMENTO"), eTipoValor.TEXTO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub
    Public Sub ObterLogradouroPessoa(ByVal Codigo As String)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH69_ENDERECO")
        strSQL.Append(" where RH01_ID_PESSOA = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            RH69_ID_ENDERECO = DoBanco(dr("RH69_ID_ENDERECO"), eTipoValor.CHAVE)
            RH01_ID_PESSOA = DoBanco(dr("RH01_ID_PESSOA"), eTipoValor.CHAVE)
            TG57_ID_LOGRADOURO_CEP = DoBanco(dr("TG57_ID_LOGRADOURO_CEP"), eTipoValor.CHAVE)
            RH69_NU_PREDIO = DoBanco(dr("RH69_NU_PREDIO"), eTipoValor.TEXTO)
            RH69_DS_COMPLEMENTO = DoBanco(dr("RH69_DS_COMPLEMENTO"), eTipoValor.TEXTO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function PesquisarEnderecoPessoa(ByVal Pessoa As Integer) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select RH69.RH69_ID_ENDERECO, RH69.RH01_ID_PESSOA ,RH69.TG57_ID_LOGRADOURO_CEP   ,  ")
        strSQL.Append(" TG56.TG56_ID_LOGRADOURO_CEP, TG55.TG55_NU_CEP, TG54.TG54_NM_LOGRADOURO, TG16.TG16_NM_TIPO_LOGRADOURO 	  , ")
        strSQL.Append(" TG04.TG04_NM_BAIRRO, TG03.TG03_NM_MUNICIPIO, TG02.TG02_SG_UF,RH69.RH69_NU_PREDIO  , RH69.RH69_DS_COMPLEMENTO	   ")
        strSQL.Append("	   ")
        strSQL.Append(" From RH69_ENDERECO As RH69 ")
        strSQL.Append(" Left Join DBGERAL.DBO.TG56_LOGRADOURO_CEP  AS TG56 ON TG56.TG56_ID_LOGRADOURO_CEP = RH69.TG57_ID_LOGRADOURO_CEP  ")
        strSQL.Append(" Left Join DBGERAL.DBO.TG55_CEP AS TG55 ON TG55.TG55_ID_CEP = TG56.TG55_ID_CEP ")
        strSQL.Append(" Left Join DBGERAL.DBO.TG54_LOGRADOURO AS TG54 ON TG54.TG54_ID_LOGRADOURO = TG56.TG54_ID_LOGRADOURO ")
        strSQL.Append(" Left Join DBGERAL.DBO.TG16_TIPO_LOGRADOURO AS TG16 ON TG16.TG16_ID_TIPO_LOGRADOURO = TG54.TG16_ID_TIPO_LOGRADOURO ")
        strSQL.Append(" Left Join DBGERAL.DBO.TG04_BAIRRO AS TG04 ON TG04.TG04_ID_BAIRRO = TG54.TG04_ID_BAIRRO ")
        strSQL.Append(" Left Join DBGERAL.DBO.TG03_MUNICIPIO AS TG03 ON TG03.TG03_ID_MUNICIPIO = TG04.TG03_ID_MUNICIPIO ")
        strSQL.Append(" Left Join DBGERAL.DBO.TG02_UF AS TG02 ON TG02.TG02_ID_UF = TG03.TG02_ID_UF ")
        strSQL.Append(" where RH69.RH01_ID_PESSOA =  " & Pessoa)



        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional EnderecoId As Integer = 0, Optional PessoaId As String = "", Optional LogradouroId As Integer = 0, Optional PredioNumero As String = "", Optional Complemento As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from RH69_ENDERECO")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where RH69_ID_ENDERECO is not null")

        If EnderecoId > 0 Then
            strSQL.Append(" and RH69_ID_ENDERECO = " & EnderecoId)
        End If

        If PessoaId > 0 Then
            strSQL.Append(" and RH01_ID_PESSOA = " & PessoaId)
        End If

        If LogradouroId > 0 Then
            strSQL.Append(" and TG56_ID_LOGRADOURO_CEP = " & LogradouroId)
        End If

        If PredioNumero <> "" Then
            strSQL.Append(" and upper(RH69_NU_PREDIO) like '%" & PredioNumero.ToUpper & "%'")
        End If

        If Complemento <> "" Then
            strSQL.Append(" and upper(RH69_DS_COMPLEMENTO) like '%" & Complemento.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "RH69_ID_ENDERECO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select RH69_ID_ENDERECO as CODIGO, RH01_ID_PESSOA as DESCRICAO")
        strSQL.Append(" from RH69_ENDERECO")
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

        strSQL.Append(" select max(RH69_ID_ENDERECO) from RH69_ENDERECO")

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
    Public Function Excluir(ByVal EnderecoId As String) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from RH69_ENDERECO")
        strSQL.Append(" where RH69_ID_ENDERECO = " & EnderecoId)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

    Public Function PesquisarCepLocal(ByVal cep As String) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select * from( ")
        strSQL.Append(" select uf, NomeSemAcento As cidade,'' AS logradouro,'' AS bairro, cep, '' AS tipo_logradouro")
        strSQL.Append(" From DBCEP.dbo.cep_unico")
        strSQL.Append(" union all ")
        strSQL.Append(" Select UF, cidade,logradouro,bairro,cep,tp_logradouro as tipo_logradouro")
        strSQL.Append(" From DBCEP.dbo.logradouros")
        strSQL.Append(" ) as endereco")
        strSQL.Append(" where Replace(cep, '-', '') =  '" & cep & "' ")


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
'*                                 19/09/2018                                 *
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

