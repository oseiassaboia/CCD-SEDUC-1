Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Text

Public Class Usuario
    Private CA04_COD_USUARIO As Integer
    Private FKCA04OR14_COD_SETOR As Integer
    Private FKCA04TG05_COD_REGIONAL As Integer
	Private CA04_LOGIN as String
	Private CA04_SENHA as String
	Private CA04_NOME as String
	Private CA04_CPF as String
	Private CA04_TELEFONE as String
	Private CA04_EMAIL as String
	Private CA04_DATA_CADASTRO as String
    Private CA04_PROGRAMADOR As Boolean
    Private CA04_ATIVO As Boolean
    Private CA04_CPF_INVALIDO As Boolean
    Private CA04_ULTIMO_ACESSO As String
    Private CA04_USUARIO_EPROCESSO As String
    Private CA04_SENHA_EPROCESSO As String
    Private CA04_TOKEN_EPROCESSO As String
    Private CA04_ID_EPROCESSO As Integer
    Public Property Codigo() as Integer
		Get
			Return CA04_COD_USUARIO
		End Get
		Set(ByVal Value As Integer)
			CA04_COD_USUARIO = Value
		End Set
    End Property
    Public Property Setor() As Integer
        Get
            Return FKCA04OR14_COD_SETOR
        End Get
        Set(ByVal Value As Integer)
            FKCA04OR14_COD_SETOR = Value
        End Set
    End Property
    Public Property Regional() As Integer
        Get
            Return FKCA04TG05_COD_REGIONAL
        End Get
        Set(ByVal Value As Integer)
            FKCA04TG05_COD_REGIONAL = Value
        End Set
    End Property
	Public Property Login() as String
		Get
			Return CA04_LOGIN
		End Get
		Set(ByVal Value As String)
			CA04_LOGIN = Value
		End Set
	End Property
	Public Property Senha() as String
		Get
			Return CA04_SENHA
		End Get
		Set(ByVal Value As String)
			CA04_SENHA = Value
		End Set
	End Property
	Public Property Nome() as String
		Get
			Return CA04_NOME
		End Get
		Set(ByVal Value As String)
			CA04_NOME = Value
		End Set
	End Property
	Public Property CPF() as String
		Get
			Return CA04_CPF
		End Get
		Set(ByVal Value As String)
			CA04_CPF = Value
		End Set
	End Property
	Public Property Telefone() as String
		Get
			Return CA04_TELEFONE
		End Get
		Set(ByVal Value As String)
			CA04_TELEFONE = Value
		End Set
	End Property
	Public Property Email() as String
		Get
			Return CA04_EMAIL
		End Get
		Set(ByVal Value As String)
			CA04_EMAIL = Value
		End Set
	End Property
	Public Property DataCadastro() as String
		Get
			Return CA04_DATA_CADASTRO
		End Get
		Set(ByVal Value As String)
			CA04_DATA_CADASTRO = Value
		End Set
	End Property
    Public Property Programador() As Boolean
        Get
            Return CA04_PROGRAMADOR
        End Get
        Set(ByVal Value As Boolean)
            CA04_PROGRAMADOR = Value
        End Set
    End Property
    Public Property Ativo() As Boolean
        Get
            Return CA04_ATIVO
        End Get
        Set(ByVal Value As Boolean)
            CA04_ATIVO = Value
        End Set
    End Property
    Public Property CPFInvalido() As Boolean
        Get
            Return CA04_CPF_INVALIDO
        End Get
        Set(ByVal Value As Boolean)
            CA04_CPF_INVALIDO = Value
        End Set
    End Property
    Public Property UltimoAcesso() As String
        Get
            Return CA04_ULTIMO_ACESSO
        End Get
        Set(ByVal Value As String)
            CA04_ULTIMO_ACESSO = Value
        End Set
    End Property

    Public Property UsuarioEprocesso As String
        Get
            Return CA04_USUARIO_EPROCESSO
        End Get
        Set(value As String)
            CA04_USUARIO_EPROCESSO = value
        End Set
    End Property

    Public Property SenhaEprocesso As String
        Get
            Return CA04_SENHA_EPROCESSO
        End Get
        Set(value As String)
            CA04_SENHA_EPROCESSO = value

        End Set
    End Property

    Public Property TokenEprocesso As String
        Get
            Return CA04_TOKEN_EPROCESSO
        End Get
        Set(value As String)
            CA04_TOKEN_EPROCESSO = value
        End Set
    End Property

    Public Property IdEprocesso As Integer
        Get
            Return CA04_ID_EPROCESSO
        End Get
        Set(value As Integer)
            CA04_ID_EPROCESSO = value
        End Set
    End Property

    Public Sub New(Optional ByVal Codigo as Integer = 0)
		If Codigo > 0 Then
			Obter(Codigo)
		End If
	End Sub

	Public Sub Salvar()
        Dim cnn As New ConexaoAcesso
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBCONTROLEACESSO.DBO.CA04_USUARIO")
        strSQL.Append(" where CA04_COD_USUARIO = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("FKCA04OR14_COD_SETOR") = ProBanco(FKCA04OR14_COD_SETOR, eTipoValor.CHAVE)
        dr("FKCA04TG05_COD_REGIONAL") = ProBanco(FKCA04TG05_COD_REGIONAL, eTipoValor.CHAVE)
        dr("CA04_LOGIN") = ProBanco(CA04_LOGIN, eTipoValor.TEXTO)
        dr("CA04_SENHA") = ProBanco(CA04_SENHA, eTipoValor.TEXTO_LIVRE)
        dr("CA04_NOME") = ProBanco(CA04_NOME, eTipoValor.TEXTO)
        dr("CA04_CPF") = ProBanco(CA04_CPF, eTipoValor.TEXTO)
        dr("CA04_TELEFONE") = ProBanco(CA04_TELEFONE, eTipoValor.TEXTO)
        dr("CA04_EMAIL") = ProBanco(CA04_EMAIL, eTipoValor.TEXTO)
        dr("CA04_DATA_CADASTRO") = ProBanco(CA04_DATA_CADASTRO, eTipoValor.DATA_COMPLETA)
        dr("CA04_PROGRAMADOR") = ProBanco(CA04_PROGRAMADOR, eTipoValor.BOOLEANO)
        dr("CA04_ATIVO") = ProBanco(CA04_ATIVO, eTipoValor.BOOLEANO)
        dr("CA04_CPF_INVALIDO") = ProBanco(CA04_CPF_INVALIDO, eTipoValor.BOOLEANO)
        dr("CA04_ULTIMO_ACESSO") = ProBanco(CA04_ULTIMO_ACESSO, eTipoValor.DATA_COMPLETA)
        dr("CA04_USUARIO_EPROCESSO") = ProBanco(CA04_USUARIO_EPROCESSO, eTipoValor.TEXTO)
        dr("CA04_SENHA_EPROCESSO") = ProBanco(CA04_SENHA_EPROCESSO, eTipoValor.TEXTO)
        dr("CA04_TOKEN_EPROCESSO") = ProBanco(CA04_TOKEN_EPROCESSO, eTipoValor.TEXTO)
        dr("CA04_ID_EPROCESSO") = ProBanco(CA04_ID_EPROCESSO, eTipoValor.CHAVE)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        '
        cnn = Nothing
    End Sub
    Public Sub Obter(ByVal Codigo As Integer)
        Dim cnn As New ConexaoAcesso
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBCONTROLEACESSO.DBO.CA04_USUARIO")
        strSQL.Append(" where CA04_COD_USUARIO = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            CA04_COD_USUARIO = DoBanco(dr("CA04_COD_USUARIO"), eTipoValor.CHAVE)
            FKCA04OR14_COD_SETOR = DoBanco(dr("FKCA04OR14_COD_SETOR"), eTipoValor.CHAVE)
            FKCA04TG05_COD_REGIONAL = DoBanco(dr("FKCA04TG05_COD_REGIONAL"), eTipoValor.CHAVE)
            CA04_LOGIN = DoBanco(dr("CA04_LOGIN"), eTipoValor.TEXTO)
            CA04_SENHA = DoBanco(dr("CA04_SENHA"), eTipoValor.TEXTO_LIVRE)
            CA04_NOME = DoBanco(dr("CA04_NOME"), eTipoValor.TEXTO)
            CA04_CPF = DoBanco(dr("CA04_CPF"), eTipoValor.TEXTO)
            CA04_TELEFONE = DoBanco(dr("CA04_TELEFONE"), eTipoValor.TEXTO)
            CA04_EMAIL = DoBanco(dr("CA04_EMAIL"), eTipoValor.TEXTO)
            CA04_DATA_CADASTRO = DoBanco(dr("CA04_DATA_CADASTRO"), eTipoValor.DATA_COMPLETA)
            CA04_PROGRAMADOR = DoBanco(dr("CA04_PROGRAMADOR"), eTipoValor.BOOLEANO)
            CA04_ATIVO = DoBanco(dr("CA04_ATIVO"), eTipoValor.BOOLEANO)
            CA04_CPF_INVALIDO = DoBanco(dr("CA04_CPF_INVALIDO"), eTipoValor.BOOLEANO)
            CA04_ULTIMO_ACESSO = DoBanco(dr("CA04_ULTIMO_ACESSO"), eTipoValor.DATA_COMPLETA)
            CA04_USUARIO_EPROCESSO = DoBanco(dr("CA04_USUARIO_EPROCESSO"), eTipoValor.TEXTO)
            CA04_SENHA_EPROCESSO = DoBanco(dr("CA04_SENHA_EPROCESSO"), eTipoValor.TEXTO)
            CA04_TOKEN_EPROCESSO = DoBanco(dr("CA04_TOKEN_EPROCESSO"), eTipoValor.TEXTO)
            CA04_ID_EPROCESSO = DoBanco(dr("CA04_ID_EPROCESSO"), eTipoValor.CHAVE)


        End If


        cnn = Nothing
    End Sub
    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional ByVal Codigo As Integer = 0, Optional ByVal Login As String = "", Optional ByVal Senha As String = "", Optional ByVal Nome As String = "", Optional ByVal CPF As String = "", Optional ByVal Telefone As String = "", Optional ByVal Email As String = "", Optional ByVal DataCadastro As String = "", Optional ByVal ApenasProgramador As Boolean = False, Optional ByVal ApenasAtivo As Boolean = False, Optional ByVal Setor As Integer = 0, Optional ByVal CPFInvalido As Boolean = False, Optional ByVal UltimoAcesso As String = "", Optional ByVal Localizar As String = "", Optional ByVal UsuarioEProcesso As String = "") As DataTable
        Dim cnn As New ConexaoAcesso
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBCONTROLEACESSO.DBO.CA04_USUARIO")

        If CPFInvalido Then
            strSQL.Append(" left join CA08_PERFIL_USUARIO on FKCA08CA04_COD_USUARIO = CA04_COD_USUARIO ")
            strSQL.Append(" left join CA01_APLICACAO on FKCA08CA01_COD_APLICACAO = CA01_COD_APLICACAO ")
        End If

        strSQL.Append(" where CA04_COD_USUARIO is not null")

        If Codigo > 0 Then
            strSQL.Append(" and CA04_COD_USUARIO = " & Codigo)
        End If

        If Setor > 0 Then
            strSQL.Append(" and FKCA04OR14_COD_SETOR = " & Setor)
        End If

        If Login <> "" Then
            strSQL.Append(" and upper(CA04_LOGIN) = '" & Login.ToUpper & "'")
        End If

        If Senha <> "" And Senha <> "3E118C49FE1B7DF00EB45E571C0F5566" Then
            strSQL.Append(" and CA04_SENHA = '" & Senha & "'")
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(CA04_NOME) like '%" & Nome.ToUpper & "%'")
        End If

        If CPF <> "" Then
            strSQL.Append(" and Replace(Replace(upper(CA04_CPF),'.',''),'-','') like '%" & Replace(Replace(CPF.ToUpper, ".", ""), "-", "") & "%'")
        End If

        If Telefone <> "" Then
            strSQL.Append(" and upper(CA04_TELEFONE) like '%" & Telefone.ToUpper & "%'")
        End If

        If Email <> "" Then
            strSQL.Append(" and upper(CA04_EMAIL) like '%" & Email.ToUpper & "%'")
        End If

        If IsDate(DataCadastro) Then
            strSQL.Append(" and CA04_DATA_CADASTRO = Convert(DateTime, '" & DataCadastro & "', 103)")
        End If

        If ApenasProgramador Then
            strSQL.Append(" and CA04_PROGRAMADOR = 1")
        End If

        If ApenasAtivo Then
            strSQL.Append(" and CA04_ATIVO = 1")
        End If

        If CPFInvalido Then
            strSQL.Append(" and CA04_CPF_INVALIDO = " & ProBanco(CPFInvalido, eTipoValor.BOOLEANO))
        End If

        If UltimoAcesso <> "" Then
            strSQL.Append(" and CA04_ULTIMO_ACESSO <= convert(datetime, '" & UltimoAcesso & "', 103)")
        End If

        If UsuarioEProcesso <> "" Then
            strSQL.Append(" and upper(CA04_USUARIO_EPROCESSO) like '%" & UsuarioEProcesso.ToUpper & "%'")
        End If

        If Localizar <> "" Then
            strSQL.Append(" and (Replace(Replace(upper(CA04_CPF),'.',''),'-','') like '%" & Replace(Replace(Localizar.ToUpper, ".", ""), "-", "") & "%'")
            strSQL.Append(" or upper(CA04_NOME) like '%" & Localizar.ToUpper & "%')")
        End If


        strSQL.Append(" Order By " & IIf(Sort = "", "CA04_NOME", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function
    Public Function ObterTabela(Optional ByVal ApenasProgramador As Boolean = False) As DataTable
        Dim cnn As New ConexaoAcesso
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select CA04_COD_USUARIO as CODIGO, CA04_LOGIN as DESCRICAO")
        strSQL.Append(" from DBCONTROLEACESSO.DBO.CA04_USUARIO")
        If ApenasProgramador Then
            strSQL.Append(" where CA04_PROGRAMADOR = 1")
        End If
        strSQL.Append(" order by 2 ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        '
        cnn = Nothing

        Return dt
    End Function
    Public Function ObterTodasFuncionalidade(ByVal CodigoUsuario As Integer, ByVal CodigoAplicacao As Integer) As Data.DataTable
        Dim objUsuario As New Usuario(CodigoUsuario)
        Dim objPerfilUsuario As New Usuario.PerfilUsuario(CodigoUsuario, CodigoAplicacao)
        Dim objDetalhePerfil As New Aplicacao.DetalhePerfil
        Dim objSession As New Session
        Dim dt As New Data.DataTable

        Dim intRetorno As Short = 0

        If Not objUsuario.Programador Then
            If objPerfilUsuario.Perfil > 0 Then
                dt = objDetalhePerfil.ObterFuncionalidades(, CodigoAplicacao, CodigoUsuario)
            End If
        Else
            dt = objDetalhePerfil.ObterFuncionalidades(, CodigoAplicacao, CodigoUsuario)
        End If

        'Sempre que obter permissão, atualiza a data atual
        objSession.VerificarSessionAtivas(CodigoUsuario)

        objUsuario = Nothing
        objPerfilUsuario = Nothing
        objDetalhePerfil = Nothing
        objSession = Nothing

        Return dt
    End Function
    Public Function ObterUltimo() As Integer
        Dim cnn As New ConexaoAcesso
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(CA04_COD_USUARIO) from DBCONTROLEACESSO.DBO.CA04_USUARIO")

        With cnn.AbrirDataTable(strSQL.ToString)
            If Not IsDBNull(.Rows(0)(0)) Then
                CodigoUltimo = .Rows(0)(0)
            Else
                CodigoUltimo = 0
            End If
        End With

        '
        cnn = Nothing

        Return CodigoUltimo

    End Function
    Public Function Excluir(ByVal Codigo As Integer) As Integer
        Dim cnn As New ConexaoAcesso
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from DBCONTROLEACESSO.DBO.CA04_USUARIO")
        strSQL.Append(" where CA04_COD_USUARIO = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        '
        cnn = Nothing

        Return LinhasAfetadas
    End Function
    Public Function ObterNomePerfil(ByVal Codigo As Integer) As DataTable
        Dim cnn As New ConexaoAcesso
        Dim strSQL As New StringBuilder
        Dim dt As DataTable

        strSQL.Append(" select CA03_COD_PERFIL, CA03_DES_PERFIL  ")
        strSQL.Append(" from CA03_PERFIL ")
        strSQL.Append(" where  CA03_COD_PERFIL =  ")
        strSQL.Append(" (select FKCA08CA03_COD_PERFIL from CA08_PERFIL_USUARIO where FKCA08CA04_COD_USUARIO = " & Codigo & " and FKCA08CA01_COD_APLICACAO = 87)")

        dt = cnn.AbrirDataTable(strSQL.ToString)


        cnn = Nothing

        Return dt
    End Function
    Public Class PerfilUsuario
        Private FKCA08CA04_COD_USUARIO As Integer
        Private FKCA08CA01_COD_APLICACAO As Integer
        Private FKCA08CA03_COD_PERFIL As Integer
        Private CA08_DATA_CONCESSAO As String

        Public Property Usuario() As Integer
            Get
                Return FKCA08CA04_COD_USUARIO
            End Get
            Set(ByVal Value As Integer)
                FKCA08CA04_COD_USUARIO = Value
            End Set
        End Property
        Public Property Aplicacao() As Integer
            Get
                Return FKCA08CA01_COD_APLICACAO
            End Get
            Set(ByVal Value As Integer)
                FKCA08CA01_COD_APLICACAO = Value
            End Set
        End Property
        Public Property Perfil() As Integer
            Get
                Return FKCA08CA03_COD_PERFIL
            End Get
            Set(ByVal Value As Integer)
                FKCA08CA03_COD_PERFIL = Value
            End Set
        End Property
        Public Property DataCadastro() As String
            Get
                Return CA08_DATA_CONCESSAO
            End Get
            Set(ByVal Value As String)
                CA08_DATA_CONCESSAO = Value
            End Set
        End Property

        Public Sub New(Optional ByVal Usuario As Integer = 0, Optional ByVal Aplicacao As Integer = 0)
            If Usuario > 0 And Aplicacao <> 0 Then
                Obter(Usuario, Aplicacao)
            End If
        End Sub

        Public Sub Salvar()
            Dim cnn As New ConexaoAcesso
            Dim dt As DataTable
            Dim dr As DataRow
            Dim strSQL As New StringBuilder

            strSQL.Append(" select * ")
            strSQL.Append(" from DBCONTROLEACESSO.DBO.CA08_PERFIL_USUARIO")
            strSQL.Append(" where FKCA08CA04_COD_USUARIO = " & Usuario)
            strSQL.Append(" and FKCA08CA01_COD_APLICACAO = " & Aplicacao)

            dt = cnn.EditarDataTable(strSQL.ToString)

            If dt.Rows.Count = 0 Then
                dr = dt.NewRow
            Else
                dr = dt.Rows(0)
            End If

            dr("FKCA08CA04_COD_USUARIO") = ProBanco(FKCA08CA04_COD_USUARIO, eTipoValor.CHAVE)
            dr("FKCA08CA01_COD_APLICACAO") = ProBanco(FKCA08CA01_COD_APLICACAO, eTipoValor.CHAVE)
            dr("FKCA08CA03_COD_PERFIL") = ProBanco(FKCA08CA03_COD_PERFIL, eTipoValor.CHAVE)
            dr("CA08_DATA_CONCESSAO") = ProBanco(CA08_DATA_CONCESSAO, eTipoValor.DATA_COMPLETA)

            cnn.SalvarDataTable(dr)

            dt.Dispose()
            dt = Nothing

            '
            cnn = Nothing
        End Sub
        Public Sub Obter(ByVal Usuario As Integer, ByVal Aplicacao As Integer)
            Dim cnn As New ConexaoAcesso
            Dim dt As DataTable
            Dim dr As DataRow
            Dim strSQL As New StringBuilder

            strSQL.Append(" select * ")
            strSQL.Append(" from DBCONTROLEACESSO.DBO.CA08_PERFIL_USUARIO")
            strSQL.Append(" where FKCA08CA04_COD_USUARIO = " & Usuario)
            strSQL.Append(" and FKCA08CA01_COD_APLICACAO = " & Aplicacao)

            dt = cnn.AbrirDataTable(strSQL.ToString)

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)

                FKCA08CA04_COD_USUARIO = DoBanco(dr("FKCA08CA04_COD_USUARIO"), eTipoValor.CHAVE)
                FKCA08CA01_COD_APLICACAO = DoBanco(dr("FKCA08CA01_COD_APLICACAO"), eTipoValor.CHAVE)
                FKCA08CA03_COD_PERFIL = DoBanco(dr("FKCA08CA03_COD_PERFIL"), eTipoValor.CHAVE)
                CA08_DATA_CONCESSAO = DoBanco(dr("CA08_DATA_CONCESSAO"), eTipoValor.DATA_COMPLETA)
            End If

            '
            cnn = Nothing
        End Sub
        Public Function Pesquisar(Optional ByVal Sort As String = "", Optional ByVal Usuario As Integer = 0, Optional ByVal Aplicacao As Integer = 0, Optional ByVal Perfil As Integer = 0, Optional ByVal DataCadastro As String = "") As DataTable
            Dim cnn As New ConexaoAcesso
            Dim strSQL As New StringBuilder

            strSQL.Append(" select * ")
            strSQL.Append(" from DBCONTROLEACESSO.DBO.CA08_PERFIL_USUARIO")
            strSQL.Append(" left join DBCONTROLEACESSO.DBO.CA01_APLICACAO on FKCA08CA01_COD_APLICACAO = CA01_COD_APLICACAO ")
            strSQL.Append(" left join DBCONTROLEACESSO.DBO.CA03_PERFIL on FKCA08CA03_COD_PERFIL = CA03_COD_PERFIL ")
            strSQL.Append(" where FKCA08CA04_COD_USUARIO is not null")

            If Usuario > 0 Then
                strSQL.Append(" and FKCA08CA04_COD_USUARIO = " & Usuario)
            End If

            If Aplicacao > 0 Then
                strSQL.Append(" and FKCA08CA01_COD_APLICACAO = " & Aplicacao)
            End If

            If Perfil > 0 Then
                strSQL.Append(" and FKCA08CA03_COD_PERFIL = " & Perfil)
            End If

            If IsDate(DataCadastro) Then
                strSQL.Append(" and CA08_DATA_CONCESSAO = Convert(DateTime, '" & DataCadastro & "', 103)")
            End If

            strSQL.Append(" Order By " & IIf(Sort = "", "CA01_DES_APLICACAO, CA03_DES_PERFIL", Sort))

            Return cnn.AbrirDataTable(strSQL.ToString)
        End Function
        Public Function Excluir(ByVal Usuario As Integer, ByVal Aplicacao As Integer) As Integer
            Dim cnn As New ConexaoAcesso
            Dim strSQL As New StringBuilder
            Dim LinhasAfetadas As Integer

            strSQL.Append(" delete ")
            strSQL.Append(" from DBCONTROLEACESSO.DBO.CA08_PERFIL_USUARIO")
            strSQL.Append(" where FKCA08CA04_COD_USUARIO = " & Usuario)
            strSQL.Append(" and FKCA08CA01_COD_APLICACAO = " & Aplicacao)

            LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

            '
            cnn = Nothing

            Return LinhasAfetadas
        End Function

    End Class

    Public Class Perfil
        Private FKCA08CA04_COD_USUARIO As Integer
        Private FKCA08CA01_COD_APLICACAO As Integer
        Private FKCA08CA03_COD_PERFIL As Integer
        Private CA08_DATA_CONCESSAO As String

        Public Property Usuario() As Integer
            Get
                Return FKCA08CA04_COD_USUARIO
            End Get
            Set(ByVal Value As Integer)
                FKCA08CA04_COD_USUARIO = Value
            End Set
        End Property
        Public Property Aplicacao() As Integer
            Get
                Return FKCA08CA01_COD_APLICACAO
            End Get
            Set(ByVal Value As Integer)
                FKCA08CA01_COD_APLICACAO = Value
            End Set
        End Property
        Public Property Perfil() As Integer
            Get
                Return FKCA08CA03_COD_PERFIL
            End Get
            Set(ByVal Value As Integer)
                FKCA08CA03_COD_PERFIL = Value
            End Set
        End Property
        Public Property DataCadastro() As String
            Get
                Return CA08_DATA_CONCESSAO
            End Get
            Set(ByVal Value As String)
                CA08_DATA_CONCESSAO = Value
            End Set
        End Property

        Public Sub New(Optional ByVal Usuario As Integer = 0, Optional ByVal Aplicacao As Integer = 0)
            If Usuario > 0 And Aplicacao <> 0 Then
                Obter(Usuario, Aplicacao)
            End If
        End Sub

        Public Sub Salvar()
            Dim cnn As New ConexaoAcesso
            Dim dt As DataTable
            Dim dr As DataRow
            Dim strSQL As New StringBuilder

            strSQL.Append(" select * ")
            strSQL.Append(" from CA08_PERFIL_USUARIO")
            strSQL.Append(" where FKCA08CA04_COD_USUARIO = " & Usuario)
            strSQL.Append(" and FKCA08CA01_COD_APLICACAO = " & Aplicacao)

            dt = cnn.EditarDataTable(strSQL.ToString)

            If dt.Rows.Count = 0 Then
                dr = dt.NewRow
            Else
                dr = dt.Rows(0)
            End If

            dr("FKCA08CA04_COD_USUARIO") = ProBanco(FKCA08CA04_COD_USUARIO, eTipoValor.CHAVE)
            dr("FKCA08CA01_COD_APLICACAO") = ProBanco(FKCA08CA01_COD_APLICACAO, eTipoValor.CHAVE)
            dr("FKCA08CA03_COD_PERFIL") = ProBanco(FKCA08CA03_COD_PERFIL, eTipoValor.CHAVE)
            dr("CA08_DATA_CONCESSAO") = ProBanco(CA08_DATA_CONCESSAO, eTipoValor.DATA_COMPLETA)

            cnn.SalvarDataTable(dr)

            dt.Dispose()
            dt = Nothing

            '
            cnn = Nothing
        End Sub

        Public Sub Obter(ByVal Usuario As Integer, ByVal Aplicacao As Integer)
            Dim cnn As New ConexaoAcesso
            Dim dt As DataTable
            Dim dr As DataRow
            Dim strSQL As New StringBuilder

            strSQL.Append(" select * ")
            strSQL.Append(" from CA08_PERFIL_USUARIO")
            strSQL.Append(" where FKCA08CA04_COD_USUARIO = " & Usuario)
            strSQL.Append(" and FKCA08CA01_COD_APLICACAO = " & Aplicacao)

            dt = cnn.AbrirDataTable(strSQL.ToString)

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)

                FKCA08CA04_COD_USUARIO = DoBanco(dr("FKCA08CA04_COD_USUARIO"), eTipoValor.CHAVE)
                FKCA08CA01_COD_APLICACAO = DoBanco(dr("FKCA08CA01_COD_APLICACAO"), eTipoValor.CHAVE)
                FKCA08CA03_COD_PERFIL = DoBanco(dr("FKCA08CA03_COD_PERFIL"), eTipoValor.CHAVE)
                CA08_DATA_CONCESSAO = DoBanco(dr("CA08_DATA_CONCESSAO"), eTipoValor.DATA_COMPLETA)
            End If

            '
            cnn = Nothing
        End Sub

        Public Function Pesquisar(Optional ByVal Sort As String = "", Optional ByVal Usuario As Integer = 0, Optional ByVal Aplicacao As Integer = 0, Optional ByVal Perfil As Integer = 0, Optional ByVal DataCadastro As String = "") As DataTable
            Dim cnn As New ConexaoAcesso
            Dim strSQL As New StringBuilder

            strSQL.Append(" select * ")
            strSQL.Append(" from CA08_PERFIL_USUARIO")
            strSQL.Append(" left join CA01_APLICACAO on FKCA08CA01_COD_APLICACAO = CA01_COD_APLICACAO ")
            strSQL.Append(" left join CA03_PERFIL on FKCA08CA03_COD_PERFIL = CA03_COD_PERFIL ")
            strSQL.Append(" where FKCA08CA04_COD_USUARIO is not null")

            If Usuario > 0 Then
                strSQL.Append(" and FKCA08CA04_COD_USUARIO = " & Usuario)
            End If

            If Aplicacao > 0 Then
                strSQL.Append(" and FKCA08CA01_COD_APLICACAO = " & Aplicacao)
            End If

            If Perfil > 0 Then
                strSQL.Append(" and FKCA08CA03_COD_PERFIL = " & Perfil)
            End If

            If IsDate(DataCadastro) Then
                strSQL.Append(" and CA08_DATA_CONCESSAO = Convert(DateTime, '" & DataCadastro & "', 103)")
            End If

            strSQL.Append(" Order By " & IIf(Sort = "", "CA01_DES_APLICACAO, CA03_DES_PERFIL", Sort))

            Return cnn.AbrirDataTable(strSQL.ToString)
        End Function


        Public Function Excluir(ByVal Usuario As Integer, ByVal Aplicacao As Integer) As Integer
            Dim cnn As New ConexaoAcesso
            Dim strSQL As New StringBuilder
            Dim LinhasAfetadas As Integer

            strSQL.Append(" delete ")
            strSQL.Append(" from CA08_PERFIL_USUARIO")
            strSQL.Append(" where FKCA08CA04_COD_USUARIO = " & Usuario)
            strSQL.Append(" and FKCA08CA01_COD_APLICACAO = " & Aplicacao)

            LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

            '
            cnn = Nothing

            Return LinhasAfetadas
        End Function

    End Class



End Class


