Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Text

Public Class Aplicacao
	Private CA01_COD_APLICACAO as Integer
	Private FKCA01CA05_COD_LIGUAGEM as Integer
	Private FKCA01CA06_COD_SGBD as Integer
	Private FKCA01CA04_COD_USUARIO_RESPONSAVEL as Integer
	Private FKCA01CA04_COD_USUARIO_PROGRAMADOR as Integer
	Private CA01_DES_APLICACAO as String
	Private CA01_ENDERECO_IP as String
	Private CA01_ENDERECO_ACESSO as String
	Private CA01_DIRETORIO_PUBLICACAO as String
	Private CA01_DATA_CADASTRO as String
    Private CA01_IMAGEM_NEW As String
    Private CA01_ANOTACOES As String

	Public Property Codigo() as Integer
		Get
			Return CA01_COD_APLICACAO
		End Get
		Set(ByVal Value As Integer)
			CA01_COD_APLICACAO = Value
		End Set
	End Property
	Public Property Linguagem() as Integer
		Get
			Return FKCA01CA05_COD_LIGUAGEM
		End Get
		Set(ByVal Value As Integer)
			FKCA01CA05_COD_LIGUAGEM = Value
		End Set
	End Property
	Public Property Sgbd() as Integer
		Get
			Return FKCA01CA06_COD_SGBD
		End Get
		Set(ByVal Value As Integer)
			FKCA01CA06_COD_SGBD = Value
		End Set
	End Property
	Public Property UsuarioResponsavel() as Integer
		Get
			Return FKCA01CA04_COD_USUARIO_RESPONSAVEL
		End Get
		Set(ByVal Value As Integer)
			FKCA01CA04_COD_USUARIO_RESPONSAVEL = Value
		End Set
	End Property
	Public Property UsuarioProgramador() as Integer
		Get
			Return FKCA01CA04_COD_USUARIO_PROGRAMADOR
		End Get
		Set(ByVal Value As Integer)
			FKCA01CA04_COD_USUARIO_PROGRAMADOR = Value
		End Set
	End Property
	Public Property Descricao() as String
		Get
			Return CA01_DES_APLICACAO
		End Get
		Set(ByVal Value As String)
			CA01_DES_APLICACAO = Value
		End Set
	End Property
	Public Property IPAplicacao() as String
		Get
			Return CA01_ENDERECO_IP
		End Get
		Set(ByVal Value As String)
			CA01_ENDERECO_IP = Value
		End Set
	End Property
	Public Property EnderecoAcesso() as String
		Get
			Return CA01_ENDERECO_ACESSO
		End Get
		Set(ByVal Value As String)
			CA01_ENDERECO_ACESSO = Value
		End Set
	End Property
	Public Property DiretorioPublicacao() as String
		Get
			Return CA01_DIRETORIO_PUBLICACAO
		End Get
		Set(ByVal Value As String)
			CA01_DIRETORIO_PUBLICACAO = Value
		End Set
	End Property
	Public Property DataCadastro() as String
		Get
			Return CA01_DATA_CADASTRO
		End Get
		Set(ByVal Value As String)
			CA01_DATA_CADASTRO = Value
		End Set
	End Property
	Public Property Imagem() as String
		Get
            Return CA01_IMAGEM_NEW
        End Get
		Set(ByVal Value As String)
            CA01_IMAGEM_NEW = Value
        End Set
    End Property
    Public Property Anotacoes() As String
        Get
            Return CA01_ANOTACOES
        End Get
        Set(ByVal Value As String)
            CA01_ANOTACOES = Value
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
        strSQL.Append(" from CA01_APLICACAO")
        strSQL.Append(" where CA01_COD_APLICACAO = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
            dr("CA01_DATA_CADASTRO") = ProBanco(CA01_DATA_CADASTRO, eTipoValor.DATA_COMPLETA)
        Else
            dr = dt.Rows(0)
        End If

        dr("FKCA01CA05_COD_LIGUAGEM") = ProBanco(FKCA01CA05_COD_LIGUAGEM, eTipoValor.CHAVE)
        dr("FKCA01CA06_COD_SGBD") = ProBanco(FKCA01CA06_COD_SGBD, eTipoValor.CHAVE)
        dr("FKCA01CA04_COD_USUARIO_RESPONSAVEL") = ProBanco(FKCA01CA04_COD_USUARIO_RESPONSAVEL, eTipoValor.CHAVE)
        dr("FKCA01CA04_COD_USUARIO_PROGRAMADOR") = ProBanco(FKCA01CA04_COD_USUARIO_PROGRAMADOR, eTipoValor.CHAVE)
        dr("CA01_DES_APLICACAO") = ProBanco(CA01_DES_APLICACAO, eTipoValor.TEXTO)
        dr("CA01_ENDERECO_IP") = ProBanco(CA01_ENDERECO_IP, eTipoValor.TEXTO_LIVRE)
        dr("CA01_ENDERECO_ACESSO") = ProBanco(CA01_ENDERECO_ACESSO, eTipoValor.TEXTO_LIVRE)
        dr("CA01_DIRETORIO_PUBLICACAO") = ProBanco(CA01_DIRETORIO_PUBLICACAO, eTipoValor.TEXTO_LIVRE)
        dr("CA01_IMAGEM_NEW") = ProBanco(CA01_IMAGEM_NEW, eTipoValor.TEXTO_LIVRE)
        dr("CA01_ANOTACOES") = ProBanco(CA01_ANOTACOES, eTipoValor.TEXTO_LIVRE)


        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        ''
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal Codigo As Integer)
        Dim cnn As New ConexaoAcesso
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from CA01_APLICACAO")
        strSQL.Append(" where CA01_COD_APLICACAO = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            CA01_COD_APLICACAO = DoBanco(dr("CA01_COD_APLICACAO"), eTipoValor.CHAVE)
            FKCA01CA05_COD_LIGUAGEM = DoBanco(dr("FKCA01CA05_COD_LIGUAGEM"), eTipoValor.CHAVE)
            FKCA01CA06_COD_SGBD = DoBanco(dr("FKCA01CA06_COD_SGBD"), eTipoValor.CHAVE)
            FKCA01CA04_COD_USUARIO_RESPONSAVEL = DoBanco(dr("FKCA01CA04_COD_USUARIO_RESPONSAVEL"), eTipoValor.CHAVE)
            FKCA01CA04_COD_USUARIO_PROGRAMADOR = DoBanco(dr("FKCA01CA04_COD_USUARIO_PROGRAMADOR"), eTipoValor.CHAVE)
            CA01_DES_APLICACAO = DoBanco(dr("CA01_DES_APLICACAO"), eTipoValor.TEXTO)
            CA01_ENDERECO_IP = DoBanco(dr("CA01_ENDERECO_IP"), eTipoValor.TEXTO_LIVRE)
            CA01_ENDERECO_ACESSO = DoBanco(dr("CA01_ENDERECO_ACESSO"), eTipoValor.TEXTO_LIVRE)
            CA01_DIRETORIO_PUBLICACAO = DoBanco(dr("CA01_DIRETORIO_PUBLICACAO"), eTipoValor.TEXTO_LIVRE)
            CA01_DATA_CADASTRO = DoBanco(dr("CA01_DATA_CADASTRO"), eTipoValor.DATA_COMPLETA)
            CA01_IMAGEM_NEW = DoBanco(dr("CA01_IMAGEM_NEW"), eTipoValor.TEXTO_LIVRE)
            CA01_ANOTACOES = DoBanco(dr("CA01_ANOTACOES"), eTipoValor.TEXTO_LIVRE)

        End If

        ''
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional ByVal Codigo As Integer = 0, Optional ByVal Linguagem As Integer = 0, Optional ByVal Sgbd As Integer = 0, Optional ByVal UsuarioResponsavel As Integer = 0, Optional ByVal UsuarioProgramador As Integer = 0, Optional ByVal Descricao As String = "", Optional ByVal IPAplicacao As String = "", Optional ByVal EnderecoAcesso As String = "", Optional ByVal DiretorioPublicacao As String = "", Optional ByVal DataCadastro As String = "", Optional ByVal Imagem As String = "") As DataTable
        Dim cnn As New ConexaoAcesso
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from CA01_APLICACAO")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where CA01_COD_APLICACAO is not null")

        If Codigo > 0 Then
            strSQL.Append(" and CA01_COD_APLICACAO = " & Codigo)
        End If

        If Linguagem > 0 Then
            strSQL.Append(" and FKCA01CA05_COD_LIGUAGEM = " & Linguagem)
        End If

        If Sgbd > 0 Then
            strSQL.Append(" and FKCA01CA06_COD_SGBD = " & Sgbd)
        End If

        If UsuarioResponsavel > 0 Then
            strSQL.Append(" and FKCA01CA04_COD_USUARIO_RESPONSAVEL = " & UsuarioResponsavel)
        End If

        If UsuarioProgramador > 0 Then
            strSQL.Append(" and FKCA01CA04_COD_USUARIO_PROGRAMADOR = " & UsuarioProgramador)
        End If

        If Descricao <> "" Then
            strSQL.Append(" and upper(CA01_DES_APLICACAO) like '%" & Descricao.ToUpper & "%'")
        End If

        If IPAplicacao <> "" Then
            strSQL.Append(" and upper(CA01_ENDERECO_IP) like '%" & IPAplicacao.ToUpper & "%'")
        End If

        If EnderecoAcesso <> "" Then
            strSQL.Append(" and upper(CA01_ENDERECO_ACESSO) like '%" & EnderecoAcesso.ToUpper & "%'")
        End If

        If DiretorioPublicacao <> "" Then
            strSQL.Append(" and upper(CA01_DIRETORIO_PUBLICACAO) like '%" & DiretorioPublicacao.ToUpper & "%'")
        End If

        If IsDate(DataCadastro) Then
            strSQL.Append(" and CA01_DATA_CADASTRO = Convert(DateTime, '" & DataCadastro & "', 103)")
        End If

        If Imagem <> "" Then
            strSQL.Append(" and upper(CA01_IMAGEM) like '%" & Imagem.ToUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "CA01_DES_APLICACAO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New ConexaoAcesso
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select CA01_COD_APLICACAO as CODIGO, CA01_DES_APLICACAO as DESCRICAO ")
        strSQL.Append(" from CA01_APLICACAO ")
        'strSQL.Append(" where CA01_COD_APLICACAO is not null")

        strSQL.Append(" where CA01_COD_APLICACAO in ( 91  )")
        'If Codigo > 0 Then
        '    strSQL.Append(" LEFT JOIN CA08_PERFIL_USUARIO on CA01_COD_APLICACAO = FKCA08CA01_COD_APLICACAO ")
        '    strSQL.Append(" WHERE FKCA08CA04_COD_USUARIO = " & Codigo)
        'End If

        strSQL.Append(" order by 2 ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        ''
        cnn = Nothing

        Return dt
    End Function

    Public Function ObterUltimo() As Integer
        Dim cnn As New ConexaoAcesso
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" select max(CA01_COD_APLICACAO) from CA01_APLICACAO")

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
    Public Function Excluir(ByVal Codigo As Integer) As Integer
        Dim cnn As New ConexaoAcesso
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from CA01_APLICACAO")
        strSQL.Append(" where CA01_COD_APLICACAO = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        ''
        cnn = Nothing

        Return LinhasAfetadas
    End Function

    Public Class Perfil
        Private CA03_COD_PERFIL As Integer
        Private FKCA03CA01_COD_APLICACAO As Integer
        Private CA03_DES_PERFIL As String

        Public Property Codigo() As Integer
            Get
                Return CA03_COD_PERFIL
            End Get
            Set(ByVal Value As Integer)
                CA03_COD_PERFIL = Value
            End Set
        End Property
        Public Property Aplicacao() As Integer
            Get
                Return FKCA03CA01_COD_APLICACAO
            End Get
            Set(ByVal Value As Integer)
                FKCA03CA01_COD_APLICACAO = Value
            End Set
        End Property
        Public Property Descricao() As String
            Get
                Return CA03_DES_PERFIL
            End Get
            Set(ByVal Value As String)
                CA03_DES_PERFIL = Value
            End Set
        End Property

        Public Sub New(Optional ByVal Codigo As Integer = 0)
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
            strSQL.Append(" from CA03_PERFIL")
            strSQL.Append(" where CA03_COD_PERFIL = " & Codigo)

            dt = cnn.EditarDataTable(strSQL.ToString)

            If dt.Rows.Count = 0 Then
                dr = dt.NewRow
            Else
                dr = dt.Rows(0)
            End If

            dr("FKCA03CA01_COD_APLICACAO") = ProBanco(FKCA03CA01_COD_APLICACAO, eTipoValor.CHAVE)
            dr("CA03_DES_PERFIL") = ProBanco(CA03_DES_PERFIL, eTipoValor.TEXTO)

            cnn.SalvarDataTable(dr)

            dt.Dispose()
            dt = Nothing

            ''
            cnn = Nothing
        End Sub

        Public Sub Obter(ByVal Codigo As Integer)
            Dim cnn As New ConexaoAcesso
            Dim dt As DataTable
            Dim dr As DataRow
            Dim strSQL As New StringBuilder

            strSQL.Append(" select * ")
            strSQL.Append(" from CA03_PERFIL")
            strSQL.Append(" where CA03_COD_PERFIL = " & Codigo)

            dt = cnn.AbrirDataTable(strSQL.ToString)

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)

                CA03_COD_PERFIL = DoBanco(dr("CA03_COD_PERFIL"), eTipoValor.CHAVE)
                FKCA03CA01_COD_APLICACAO = DoBanco(dr("FKCA03CA01_COD_APLICACAO"), eTipoValor.CHAVE)
                CA03_DES_PERFIL = DoBanco(dr("CA03_DES_PERFIL"), eTipoValor.TEXTO)
            End If

            ''
            cnn = Nothing
        End Sub

        Public Function Pesquisar(Optional ByVal Sort As String = "", Optional ByVal Codigo As Integer = 0, Optional ByVal Aplicacao As Integer = 0, Optional ByVal Descricao As String = "") As DataTable
            Dim cnn As New ConexaoAcesso
            Dim strSQL As New StringBuilder

            strSQL.Append(" select * ")
            strSQL.Append(" from CA03_PERFIL")
            'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
            strSQL.Append(" where CA03_COD_PERFIL is not null")

            If Codigo > 0 Then
                strSQL.Append(" and CA03_COD_PERFIL = " & Codigo)
            End If

            If Aplicacao > 0 Then
                strSQL.Append(" and FKCA03CA01_COD_APLICACAO = " & Aplicacao)
            End If

            If Descricao <> "" Then
                strSQL.Append(" and upper(CA03_DES_PERFIL) like '%" & Descricao.ToUpper & "%'")
            End If

            strSQL.Append(" Order By " & IIf(Sort = "", "CA03_COD_PERFIL", Sort))

            Return cnn.AbrirDataTable(strSQL.ToString)
        End Function

        Public Function ObterTabela(Optional ByVal Aplicacao As Integer = 0) As DataTable
            Dim cnn As New ConexaoAcesso
            Dim dt As DataTable
            Dim strSQL As New StringBuilder

            strSQL.Append(" select CA03_COD_PERFIL as CODIGO, CA03_DES_PERFIL as DESCRICAO")
            strSQL.Append(" from CA03_PERFIL")
            strSQL.Append(" where CA03_COD_PERFIL is not null")

            If Aplicacao > 0 Then
                strSQL.Append(" and FKCA03CA01_COD_APLICACAO = " & Aplicacao)
            End If

            strSQL.Append(" and CA03_COD_PERFIL <> 1")
            strSQL.Append(" order by 2 ")

            dt = cnn.AbrirDataTable(strSQL.ToString)

            ''
            cnn = Nothing

            Return dt
        End Function

        Public Function ObterUltimo() As Integer
            Dim cnn As New ConexaoAcesso
            Dim strSQL As New StringBuilder
            Dim CodigoUltimo As Integer

            strSQL.Append(" select max(CA03_COD_PERFIL) from CA03_PERFIL")

            With cnn.AbrirDataTable(strSQL.ToString)
                If Not IsDBNull(.Rows(0)(0)) Then
                    CodigoUltimo = .Rows(0)(0)
                Else
                    CodigoUltimo = 0
                End If
            End With

            ''
            cnn = Nothing

            Return CodigoUltimo

        End Function
        Public Function Excluir(ByVal Codigo As Integer) As Integer
            Dim cnn As New ConexaoAcesso
            Dim strSQL As New StringBuilder
            Dim LinhasAfetadas As Integer

            strSQL.Append(" delete ")
            strSQL.Append(" from CA03_PERFIL")
            strSQL.Append(" where CA03_COD_PERFIL = " & Codigo)

            LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

            ''
            cnn = Nothing

            Return LinhasAfetadas
        End Function

    End Class


    Public Class Funcionalidade
        Private CA02_COD_FUNCIONALIDADE As Integer
        Private FKCA02CA01_COD_APLICACAO As Integer
        Private CA02_DES_FUNCIONALIDADE As String
        Private CA02_ANOTACOES As String

        Public Property Codigo() As Integer
            Get
                Return CA02_COD_FUNCIONALIDADE
            End Get
            Set(ByVal Value As Integer)
                CA02_COD_FUNCIONALIDADE = Value
            End Set
        End Property
        Public Property Aplicacao() As Integer
            Get
                Return FKCA02CA01_COD_APLICACAO
            End Get
            Set(ByVal Value As Integer)
                FKCA02CA01_COD_APLICACAO = Value
            End Set
        End Property
        Public Property Descricao() As String
            Get
                Return CA02_DES_FUNCIONALIDADE
            End Get
            Set(ByVal Value As String)
                CA02_DES_FUNCIONALIDADE = Value
            End Set
        End Property
        Public Property Anotacoes() As String
            Get
                Return CA02_ANOTACOES
            End Get
            Set(ByVal Value As String)
                CA02_ANOTACOES = Value
            End Set
        End Property


        Public Sub New(Optional ByVal Codigo As Integer = 0)
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
            strSQL.Append(" from CA02_FUNCIONALIDADE")
            strSQL.Append(" where CA02_COD_FUNCIONALIDADE = " & Codigo)

            dt = cnn.EditarDataTable(strSQL.ToString)

            If dt.Rows.Count = 0 Then
                dr = dt.NewRow
            Else
                dr = dt.Rows(0)
            End If

            dr("FKCA02CA01_COD_APLICACAO") = ProBanco(FKCA02CA01_COD_APLICACAO, eTipoValor.CHAVE)
            dr("CA02_DES_FUNCIONALIDADE") = ProBanco(CA02_DES_FUNCIONALIDADE, eTipoValor.TEXTO)
            dr("CA02_ANOTACOES") = ProBanco(CA02_ANOTACOES, eTipoValor.TEXTO_LIVRE)

            cnn.SalvarDataTable(dr)

            dt.Dispose()
            dt = Nothing

            ''
            cnn = Nothing
        End Sub

        Public Sub Obter(ByVal Codigo As Integer)
            Dim cnn As New ConexaoAcesso
            Dim dt As DataTable
            Dim dr As DataRow
            Dim strSQL As New StringBuilder

            strSQL.Append(" select * ")
            strSQL.Append(" from CA02_FUNCIONALIDADE")
            strSQL.Append(" where CA02_COD_FUNCIONALIDADE = " & Codigo)

            dt = cnn.AbrirDataTable(strSQL.ToString)

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)

                CA02_COD_FUNCIONALIDADE = DoBanco(dr("CA02_COD_FUNCIONALIDADE"), eTipoValor.CHAVE)
                FKCA02CA01_COD_APLICACAO = DoBanco(dr("FKCA02CA01_COD_APLICACAO"), eTipoValor.CHAVE)
                CA02_DES_FUNCIONALIDADE = DoBanco(dr("CA02_DES_FUNCIONALIDADE"), eTipoValor.TEXTO)
                CA02_ANOTACOES = DoBanco(dr("CA02_ANOTACOES"), eTipoValor.TEXTO_LIVRE)

            End If

            ''
            cnn = Nothing
        End Sub

        Public Function Pesquisar(Optional ByVal Sort As String = "", Optional ByVal Codigo As Integer = 0, Optional ByVal Aplicacao As Integer = 0, Optional ByVal Descricao As String = "") As DataTable
            Dim cnn As New ConexaoAcesso
            Dim strSQL As New StringBuilder

            strSQL.Append(" select * ")
            strSQL.Append(" from CA02_FUNCIONALIDADE")
            'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
            strSQL.Append(" where CA02_COD_FUNCIONALIDADE is not null")

            If Codigo > 0 Then
                strSQL.Append(" and CA02_COD_FUNCIONALIDADE = " & Codigo)
            End If

            If Aplicacao > 0 Then
                strSQL.Append(" and FKCA02CA01_COD_APLICACAO = " & Aplicacao)
            End If

            If Descricao <> "" Then
                strSQL.Append(" and upper(CA02_DES_FUNCIONALIDADE) like '%" & Descricao.ToUpper & "%'")
            End If

            strSQL.Append(" Order By " & IIf(Sort = "", "CA02_DES_FUNCIONALIDADE", Sort))

            Return cnn.AbrirDataTable(strSQL.ToString)
        End Function

        Public Function ObterTabela(Optional ByVal Aplicacao As Integer = 0) As DataTable
            Dim cnn As New ConexaoAcesso
            Dim dt As DataTable
            Dim strSQL As New StringBuilder

            strSQL.Append(" select CA02_COD_FUNCIONALIDADE as CODIGO, CA02_DES_FUNCIONALIDADE as DESCRICAO")
            strSQL.Append(" from CA02_FUNCIONALIDADE")
            If Aplicacao > 0 Then
                strSQL.Append(" where FKCA02CA01_COD_APLICACAO = " & Aplicacao)
            End If
            strSQL.Append(" order by 2 ")

            dt = cnn.AbrirDataTable(strSQL.ToString)

            ''
            cnn = Nothing

            Return dt
        End Function

        Public Function ObterUltimo() As Integer
            Dim cnn As New ConexaoAcesso
            Dim strSQL As New StringBuilder
            Dim CodigoUltimo As Integer

            strSQL.Append(" select max(CA02_COD_FUNCIONALIDADE) from CA02_FUNCIONALIDADE")

            With cnn.AbrirDataTable(strSQL.ToString)
                If Not IsDBNull(.Rows(0)(0)) Then
                    CodigoUltimo = .Rows(0)(0)
                Else
                    CodigoUltimo = 0
                End If
            End With

            ''
            cnn = Nothing

            Return CodigoUltimo

        End Function
        Public Function Excluir(ByVal Codigo As Integer) As Integer
            Dim cnn As New ConexaoAcesso
            Dim strSQL As New StringBuilder
            Dim LinhasAfetadas As Integer

            strSQL.Append(" delete ")
            strSQL.Append(" from CA02_FUNCIONALIDADE")
            strSQL.Append(" where CA02_COD_FUNCIONALIDADE = " & Codigo)

            LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

            ''
            cnn = Nothing

            Return LinhasAfetadas
        End Function



    End Class
    Public Class DetalhePerfil
        Private FKCA07CA02_COD_FUNCIONALIDADE As Integer
        Private FKCA07CA03_COD_PERFIL As Integer
        Private CA07_SOMENTE_LEITURA As Boolean

        Public Property Funcionalidade() As Integer
            Get
                Return FKCA07CA02_COD_FUNCIONALIDADE
            End Get
            Set(ByVal Value As Integer)
                FKCA07CA02_COD_FUNCIONALIDADE = Value
            End Set
        End Property
        Public Property Perfil() As Integer
            Get
                Return FKCA07CA03_COD_PERFIL
            End Get
            Set(ByVal Value As Integer)
                FKCA07CA03_COD_PERFIL = Value
            End Set
        End Property
        Public Property SomenteLeitura() As Boolean
            Get
                Return CA07_SOMENTE_LEITURA
            End Get
            Set(ByVal Value As Boolean)
                CA07_SOMENTE_LEITURA = Value
            End Set
        End Property

        Public Sub New(Optional ByVal Perfil As Integer = 0, Optional ByVal Funcionalidade As Integer = 0)
            If Funcionalidade > 0 And Perfil <> 0 Then
                Obter(Perfil, Funcionalidade)
            End If
        End Sub

        Public Sub Salvar()
            Dim cnn As New ConexaoAcesso
            Dim dt As DataTable
            Dim dr As DataRow
            Dim strSQL As New StringBuilder

            strSQL.Append(" select * ")
            strSQL.Append(" from CA07_DETALHE_PERFIL")
            strSQL.Append(" where FKCA07CA02_COD_FUNCIONALIDADE = " & Funcionalidade)
            strSQL.Append(" and FKCA07CA03_COD_PERFIL = " & Perfil)

            dt = cnn.EditarDataTable(strSQL.ToString)

            If dt.Rows.Count = 0 Then
                dr = dt.NewRow
            Else
                dr = dt.Rows(0)
            End If

            dr("FKCA07CA02_COD_FUNCIONALIDADE") = ProBanco(FKCA07CA02_COD_FUNCIONALIDADE, eTipoValor.CHAVE)
            dr("FKCA07CA03_COD_PERFIL") = ProBanco(FKCA07CA03_COD_PERFIL, eTipoValor.CHAVE)
            dr("CA07_SOMENTE_LEITURA") = ProBanco(CA07_SOMENTE_LEITURA, eTipoValor.BOOLEANO)

            cnn.SalvarDataTable(dr)

            dt.Dispose()
            dt = Nothing

            '
            cnn = Nothing
        End Sub

        Public Sub Obter(ByVal Perfil As Integer, ByVal Funcionalidade As Integer)
            Dim cnn As New ConexaoAcesso
            Dim dt As DataTable
            Dim dr As DataRow
            Dim strSQL As New StringBuilder

            strSQL.Append(" select * ")
            strSQL.Append(" from CA07_DETALHE_PERFIL")
            strSQL.Append(" where FKCA07CA02_COD_FUNCIONALIDADE = " & Funcionalidade)
            strSQL.Append(" and FKCA07CA03_COD_PERFIL = " & Perfil)

            dt = cnn.AbrirDataTable(strSQL.ToString)

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)

                FKCA07CA02_COD_FUNCIONALIDADE = DoBanco(dr("FKCA07CA02_COD_FUNCIONALIDADE"), eTipoValor.CHAVE)
                FKCA07CA03_COD_PERFIL = DoBanco(dr("FKCA07CA03_COD_PERFIL"), eTipoValor.CHAVE)
                CA07_SOMENTE_LEITURA = DoBanco(dr("CA07_SOMENTE_LEITURA"), eTipoValor.BOOLEANO)
            End If

            '
            cnn = Nothing
        End Sub

        Public Function Pesquisar(Optional ByVal Sort As String = "", Optional ByVal Aplicacao As Integer = 0, Optional ByVal Funcionalidade As Integer = 0, Optional ByVal Perfil As Integer = 0, Optional ByVal ApenasSomenteLeitura As Boolean = False) As DataTable
            Dim cnn As New ConexaoAcesso
            Dim strSQL As New StringBuilder

            strSQL.Append(" select *, ")
            strSQL.Append(" case CA07_SOMENTE_LEITURA when 0 then 'ESCRITA' else 'LEITURA' end ESCRITA_LEITURA ")
            strSQL.Append(" from CA07_DETALHE_PERFIL")
            strSQL.Append(" left join CA03_PERFIL on FKCA07CA03_COD_PERFIL = CA03_COD_PERFIL ")
            strSQL.Append(" left join CA02_FUNCIONALIDADE on FKCA07CA02_COD_FUNCIONALIDADE = CA02_COD_FUNCIONALIDADE ")
            strSQL.Append(" where FKCA07CA02_COD_FUNCIONALIDADE is not null")

            If Aplicacao > 0 Then
                strSQL.Append(" and FKCA03CA01_COD_APLICACAO = " & Aplicacao)
            End If

            If Funcionalidade > 0 Then
                strSQL.Append(" and FKCA07CA02_COD_FUNCIONALIDADE = " & Funcionalidade)
            End If

            If Perfil > 0 Then
                strSQL.Append(" and FKCA07CA03_COD_PERFIL = " & Perfil)
            End If

            If ApenasSomenteLeitura Then
                strSQL.Append(" and CA07_SOMENTE_LEITURA = 1")
            End If

            strSQL.Append(" Order By " & IIf(Sort = "", "CA03_DES_PERFIL, CA02_DES_FUNCIONALIDADE", Sort))

            Return cnn.AbrirDataTable(strSQL.ToString)
        End Function

        Public Function ObterFuncionalidades(Optional ByVal Sort As String = "", Optional ByVal Aplicacao As Integer = 0, Optional ByVal Usuario As Integer = 0) As DataTable
            Dim cnn As New ConexaoAcesso
            Dim strSQL As New StringBuilder
            Dim objUsuario As New Usuario(Usuario)

            strSQL.Append(" Select CA02_COD_FUNCIONALIDADE, ")

            If Not objUsuario.Programador Then
                strSQL.Append(" Case CA07_SOMENTE_LEITURA When 0 Then 'ESCRITA' else 'LEITURA' end ESCRITA_LEITURA  ")
            Else
                strSQL.Append(" 'ESCRITA' as ESCRITA_LEITURA  ")
            End If

            strSQL.Append(" From CA02_FUNCIONALIDADE ")

            If Not objUsuario.Programador Then
                strSQL.Append(" Left Join CA07_DETALHE_PERFIL on FKCA07CA02_COD_FUNCIONALIDADE = CA02_COD_FUNCIONALIDADE ")
                strSQL.Append(" Left Join CA08_PERFIL_USUARIO on FKCA08CA03_COD_PERFIL = FKCA07CA03_COD_PERFIL ")
            End If

            strSQL.Append(" where CA02_COD_FUNCIONALIDADE is not null ")

            If Aplicacao > 0 Then
                strSQL.Append(" and FKCA02CA01_COD_APLICACAO = " & Aplicacao)
            End If

            If Usuario > 0 And Not objUsuario.Programador Then
                strSQL.Append(" And FKCA08CA04_COD_USUARIO = " & Usuario)
            End If

            strSQL.Append(" Order By " & IIf(Sort = "", "CA02_COD_FUNCIONALIDADE", Sort))

            Return cnn.AbrirDataTable(strSQL.ToString)
        End Function

        Public Function ObterTabela() As DataTable
            Dim cnn As New ConexaoAcesso
            Dim dt As DataTable
            Dim strSQL As New StringBuilder

            strSQL.Append(" select FKCA07CA02_COD_FUNCIONALIDADE as CODIGO, ")
            strSQL.Append(" from CA07_DETALHE_PERFIL")
            strSQL.Append(" order by 2 ")

            dt = cnn.AbrirDataTable(strSQL.ToString)

            '
            cnn = Nothing

            Return dt
        End Function

        Public Function ObterUltimo() As Integer
            Dim cnn As New ConexaoAcesso
            Dim strSQL As New StringBuilder
            Dim CodigoUltimo As Integer

            strSQL.Append(" select max(FKCA07CA02_COD_FUNCIONALIDADE) from CA07_DETALHE_PERFIL")
            strSQL.Append(" select max(FKCA07CA03_COD_PERFIL) from CA07_DETALHE_PERFIL")

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

        Public Function Excluir(ByVal Perfil As Integer, ByVal Funcionalidade As Integer) As Integer
            Dim cnn As New ConexaoAcesso
            Dim strSQL As New StringBuilder
            Dim LinhasAfetadas As Integer

            strSQL.Append(" delete ")
            strSQL.Append(" from CA07_DETALHE_PERFIL")
            strSQL.Append(" where FKCA07CA02_COD_FUNCIONALIDADE = " & Funcionalidade)
            strSQL.Append(" and FKCA07CA03_COD_PERFIL = " & Perfil)

            LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

            '
            cnn = Nothing

            Return LinhasAfetadas
        End Function

    End Class
End Class

'******************************************************************************
'*                                 19/08/2011                                 *
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

