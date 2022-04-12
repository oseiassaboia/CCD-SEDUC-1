Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Text

Public Class HorarioTurma

    Private DE13_ID_HORARIO_TURMA As Integer
    Private DE16_ID_HORARIO As Integer
    Private DE10_ID_TURMA As Integer
    Private RH80_ID_ALOCACAO_CARGA_HORARIA As Integer
    Private DE09_ID_DISCIPLINA As Integer
    Private CA04_ID_USUARIO As Integer
    Private DE13_NU_DIA_SEMANA As String
    Private DE13_DH_CADASTRO As String
    Private DE13_DH_ALTERACAO As String
    Private CA04_ID_USUARIO_ALT As Integer
    Private CA04_ID_USUARIO_EXC As Integer
    Private DE13_IN_HORARIO_AULA_PRATICA As String
    Private DE13_DH_EXCLUSAO As String

    Public Property Codigo() As Integer
        Get
            Return DE13_ID_HORARIO_TURMA
        End Get
        Set(ByVal Value As Integer)
            DE13_ID_HORARIO_TURMA = Value
        End Set
    End Property
    Public Property Horario() As Integer
        Get
            Return DE16_ID_HORARIO
        End Get
        Set(ByVal Value As Integer)
            DE16_ID_HORARIO = Value
        End Set
    End Property
    Public Property Turma() As Integer
        Get
            Return DE10_ID_TURMA
        End Get
        Set(ByVal Value As Integer)
            DE10_ID_TURMA = Value
        End Set
    End Property
    Public Property AlocacaoCargaHoraria() As Integer
        Get
            Return RH80_ID_ALOCACAO_CARGA_HORARIA
        End Get
        Set(ByVal Value As Integer)
            RH80_ID_ALOCACAO_CARGA_HORARIA = Value
        End Set
    End Property
    Public Property Disciplina() As Integer
        Get
            Return DE09_ID_DISCIPLINA
        End Get
        Set(ByVal Value As Integer)
            DE09_ID_DISCIPLINA = Value
        End Set
    End Property
    Public Property Usuario() As Integer
        Get
            Return CA04_ID_USUARIO
        End Get
        Set(ByVal Value As Integer)
            CA04_ID_USUARIO = Value
        End Set
    End Property
    Public Property DiaSemana() As String
        Get
            Return DE13_NU_DIA_SEMANA
        End Get
        Set(ByVal Value As String)
            DE13_NU_DIA_SEMANA = Value
        End Set
    End Property
    Public Property DataHoraCadastro() As String
        Get
            Return DE13_DH_CADASTRO
        End Get
        Set(ByVal Value As String)
            DE13_DH_CADASTRO = Value
        End Set
    End Property
    Public Property DataAlteracao() As String
        Get
            Return DE13_DH_ALTERACAO
        End Get
        Set(ByVal Value As String)
            DE13_DH_ALTERACAO = Value
        End Set
    End Property
    Public Property CodigoUsuarioAlteracao() As Integer
        Get
            Return CA04_ID_USUARIO_ALT
        End Get
        Set(ByVal Value As Integer)
            CA04_ID_USUARIO_ALT = Value
        End Set
    End Property

    Public Property CodigoUsuarioExclusao() As Integer
        Get
            Return CA04_ID_USUARIO_EXC
        End Get
        Set(ByVal Value As Integer)
            CA04_ID_USUARIO_EXC = Value
        End Set
    End Property
    Public Property AulaPratica() As String
        Get
            Return DE13_IN_HORARIO_AULA_PRATICA
        End Get
        Set(ByVal Value As String)
            DE13_IN_HORARIO_AULA_PRATICA = Value
        End Set
    End Property
    Public Property DataHoraExclusao() As String
        Get
            Return DE13_DH_EXCLUSAO
        End Get
        Set(ByVal Value As String)
            DE13_DH_EXCLUSAO = Value
        End Set
    End Property

    Public Sub New(Optional ByVal Codigo As Integer = 0)
        If Codigo > 0 Then
            Obter(Codigo)
        End If
    End Sub

    Public Sub Salvar()
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from dbdiario..DE13_HORARIO_TURMA")
        strSQL.Append(" where DE13_ID_HORARIO_TURMA = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("DE16_ID_HORARIO") = ProBanco(DE16_ID_HORARIO, eTipoValor.CHAVE)
        dr("DE10_ID_TURMA") = ProBanco(DE10_ID_TURMA, eTipoValor.CHAVE)
        dr("RH80_ID_ALOCACAO_CARGA_HORARIA") = ProBanco(RH80_ID_ALOCACAO_CARGA_HORARIA, eTipoValor.CHAVE)
        dr("DE09_ID_DISCIPLINA") = ProBanco(DE09_ID_DISCIPLINA, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("DE13_NU_DIA_SEMANA") = ProBanco(DE13_NU_DIA_SEMANA, eTipoValor.TEXTO)
        dr("DE13_DH_CADASTRO") = ProBanco(DE13_DH_CADASTRO, eTipoValor.DATA_COMPLETA)
        dr("DE13_DH_ALTERACAO") = ProBanco(DE13_DH_ALTERACAO, eTipoValor.DATA_COMPLETA)
        dr("CA04_ID_USUARIO_ALT") = ProBanco(CA04_ID_USUARIO_ALT, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO_EXC") = ProBanco(CA04_ID_USUARIO_EXC, eTipoValor.CHAVE)
        dr("DE13_IN_HORARIO_AULA_PRATICA") = ProBanco(DE13_IN_HORARIO_AULA_PRATICA, eTipoValor.BOOLEANO)
        dr("DE13_DH_EXCLUSAO") = ProBanco(DE13_DH_EXCLUSAO, eTipoValor.DATA_COMPLETA)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing


        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal Codigo As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from dbdiario..DBDIARIO..DE13_HORARIO_TURMA")
        strSQL.Append(" where DE13_ID_HORARIO_TURMA = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            DE13_ID_HORARIO_TURMA = DoBanco(dr("DE13_ID_HORARIO_TURMA"), eTipoValor.CHAVE)
            DE16_ID_HORARIO = DoBanco(dr("DE16_ID_HORARIO"), eTipoValor.CHAVE)
            DE10_ID_TURMA = DoBanco(dr("DE10_ID_TURMA"), eTipoValor.CHAVE)
            RH80_ID_ALOCACAO_CARGA_HORARIA = DoBanco(dr("RH80_ID_ALOCACAO_CARGA_HORARIA"), eTipoValor.CHAVE)
            DE09_ID_DISCIPLINA = DoBanco(dr("DE09_ID_DISCIPLINA"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            DE13_NU_DIA_SEMANA = DoBanco(dr("DE13_NU_DIA_SEMANA"), eTipoValor.TEXTO)
            DE13_DH_CADASTRO = DoBanco(dr("DE13_DH_CADASTRO"), eTipoValor.DATA_COMPLETA)
            DE13_DH_ALTERACAO = DoBanco(dr("DE13_DH_ALTERACAO"), eTipoValor.DATA_COMPLETA)
            CA04_ID_USUARIO_ALT = DoBanco(dr("CA04_ID_USUARIO_ALT"), eTipoValor.CHAVE)
            CA04_ID_USUARIO_EXC = DoBanco(dr("CA04_ID_USUARIO_EXC"), eTipoValor.CHAVE)
            DE13_IN_HORARIO_AULA_PRATICA = DoBanco(dr("DE13_IN_HORARIO_AULA_PRATICA"), eTipoValor.BOOLEANO)
            DE13_DH_EXCLUSAO = DoBanco(dr("DE13_DH_EXCLUSAO"), eTipoValor.DATA_COMPLETA)
        End If


        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional Horario As Integer = 0 _
                              , Optional Turma As Integer = 0, Optional Professor As Integer = 0, Optional Usuario As Integer = 0 _
                              , Optional DiaSemana As String = "", Optional DataHoraCadastro As String = "" _
                              , Optional DataAlteracao As String = "", Optional CodigoUsuarioAlteracao As Integer = 0, Optional _
                                 ByVal Disciplina As Integer = 0, Optional ByVal Periodo As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from dbdiario..DE13_HORARIO_TURMA de13 ")
        strSQL.Append(" INNER JOIN dbdiario..DE10_TURMA		DE10	ON	DE13.DE10_ID_TURMA = DE10.DE10_ID_TURMA ")
        strSQL.Append(" INNER JOIN dbdiario..DE04_PERIODO	DE04	ON	DE04.DE04_ID_PERIODO = DE10.DE04_ID_PERIODO ")
        strSQL.Append(" where DE13_ID_HORARIO_TURMA is not null and DE13_DH_EXCLUSAO is null ")

        If Codigo > 0 Then
            strSQL.Append(" and DE13_ID_HORARIO_TURMA = " & Codigo)
        End If

        If Horario > 0 Then
            strSQL.Append(" and DE16_ID_HORARIO = " & Horario)
        End If

        If Turma > 0 Then
            strSQL.Append(" and DE10_ID_TURMA = " & Turma)
        End If

        If Disciplina > 0 Then
            strSQL.Append(" and DE09_ID_DISCIPLINA = " & Disciplina)
        End If

        If Professor > 0 Then
            strSQL.Append(" and RH80_ID_ALOCACAO_CARGA_HORARIA = " & Professor)
        End If

        If Usuario > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO = " & Usuario)
        End If

        If Periodo <> "" Then
            strSQL.Append(" and DE13_NU_DIA_SEMANA ='" & Periodo & "'")
        End If

        If DiaSemana <> "" Then
            strSQL.Append(" and upper(DE13_NU_DIA_SEMANA) like '%" & DiaSemana.ToUpper & "%'")
        End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" and DE13_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)")
        End If

        If IsDate(DataAlteracao) Then
            strSQL.Append(" and DE13_DH_ALTERACAO = Convert(DateTime, '" & DataAlteracao & "', 103)")
        End If

        If CodigoUsuarioAlteracao > 0 Then
            strSQL.Append(" and CA04_ID_USUARIO_ALT = " & CodigoUsuarioAlteracao)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "DE13_ID_HORARIO_TURMA", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarProfessorCargaHoraria(Optional ByVal Sort As String = "",
                                                   Optional Codigo As Integer = 0,
                                                   Optional Turma As Integer = 0,
                                                   Optional Lotacao As Integer = 0,
                                                   Optional DiaDaSemana As Integer = 0,
                                                   Optional Horario As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select DE13.DE13_ID_HORARIO_TURMA,  DE13.DE10_ID_TURMA, DE16.DE16_ID_HORARIO, RH02_CD_MATRICULA + ' - ' + RH01_NM_PESSOA as PROFESSOR, DE09_NM_DISCIPLINA, RH79_NM_TIPO_CARGA_HORARIA ")
        strSQL.Append(" from DBDIARIO..DE13_HORARIO_TURMA as DE13 ")
        strSQL.Append(" left join DBDIARIO..DE16_HORARIO as DE16 on DE16.DE16_ID_HORARIO = DE13.DE16_ID_HORARIO ")
        strSQL.Append(" Left Join DBDIARIO..DE10_TURMA as DE10 on DE10.DE10_ID_TURMA = DE13.DE10_ID_TURMA ")
        strSQL.Append(" Left Join DBDIARIO..DE09_DISCIPLINA as DE09 on DE09.DE09_ID_DISCIPLINA = DE13.DE09_ID_DISCIPLINA ")
        strSQL.Append(" Left Join RH80_ALOCACAO_CARGA_HORARIA as RH80 on RH80.RH80_ID_ALOCACAO_CARGA_HORARIA = DE13.RH80_ID_ALOCACAO_CARGA_HORARIA ")
        strSQL.Append(" Left Join RH78_SERVIDOR_CARGA_HORARIA as RH78 on RH78.RH78_ID_SERVIDOR_CARGA_HORARIA = RH80.RH78_ID_SERVIDOR_CARGA_HORARIA ")
        strSQL.Append(" Left Join RH77_CARGA_HORARIA as RH77 on RH77.RH77_ID_CARGA_HORARIA = RH78.RH77_ID_CARGA_HORARIA ")
        strSQL.Append(" Left Join RH79_TIPO_CARGA_HORARIA as RH79 on RH79.RH79_ID_TIPO_CARGA_HORARIA = RH77.RH79_ID_TIPO_CARGA_HORARIA ")
        strSQL.Append(" Left Join RH14_LOTACAO_SERVIDOR as RH14 on RH14.RH14_ID_LOTACAO_SERVIDOR = RH80.RH14_ID_LOTACAO_SERVIDOR ")
        strSQL.Append(" Left Join RH02_SERVIDOR as RH02 on RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR ")
        strSQL.Append(" Left Join RH01_PESSOA as RH01 on RH01.RH01_ID_PESSOA = RH02.RH01_ID_PESSOA ")
        strSQL.Append(" where DE13.DE13_ID_HORARIO_TURMA is not null ")
        strSQL.Append(" and DE16.RH36_ID_LOTACAO = " & Lotacao)
        strSQL.Append(" And DE10.DE10_ID_TURMA = " & Turma)
        strSQL.Append(" And DE13.DE13_DH_EXCLUSAO IS NULL ")


        If Codigo > 0 Then
            strSQL.Append(" And DE13_ID_HORARIO_TURMA = " & Codigo)
        End If

        If DiaDaSemana > 0 Then
            strSQL.Append(" And  DE13.DE13_NU_DIA_SEMANA = " & DiaDaSemana)
        End If

        If Horario > 0 Then
            strSQL.Append(" And  DE16.DE16_ID_HORARIO = " & Horario)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "DE13_ID_HORARIO_TURMA", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarQtdProfessorDisciplina(ByVal Turma As Integer, ByVal Disciplina As Integer) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select DE38_ID_TURMA_DISCIPLINA, DE09_ID_DISCIPLINA, DE38_ST_DISCIPLINA  ")
        strSQL.Append(",   (select count(RH80_ID_ALOCACAO_CARGA_HORARIA)  ")
        strSQL.Append("     from DBDIARIO..DE13_HORARIO_TURMA as DE13")
        strSQL.Append("     where DE13.DE10_ID_TURMA = DE38.DE10_ID_TURMA ")
        strSQL.Append("     and DE13.DE09_ID_DISCIPLINA = DE38.DE09_ID_DISCIPLINA) as QTD")
        strSQL.Append(" from DBDIARIO..DE38_TURMA_DISCIPLINA as DE38 ")
        strSQL.Append("  where DE38_ID_TURMA_DISCIPLINA is not null ")
        strSQL.Append(" And DE10_ID_TURMA = " & Turma)
        strSQL.Append(" And DE09_ID_DISCIPLINA = " & Disciplina)

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function CarregarHorario(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional Turma As Integer = 0, Optional Lotacao As Integer = 0, Optional Turno As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select  DE16.DE16_ID_HORARIO, Convert(varchar, DE16_NR_HORA_INICIO) + ':' +  right('00' + convert(varchar,DE16_NR_MINUTO_INICIO),2) + ' às '  ")
        strSQL.Append(" + Convert(varchar, DE16_NR_HORA_TERMINO) + ':' +  right('00' + convert(varchar,DE16_NR_MINUTO_TERMINO),2) as HORARIO,  ")
        'DOMINGO
        strSQL.Append("     (SELECT top 1 DE13_ID_HORARIO_TURMA ")
        strSQL.Append("     From dbdiario..DE13_HORARIO_TURMA As DE13 ")
        strSQL.Append("     WHERE DE13.DE10_ID_TURMA = " & Turma)
        strSQL.Append("     and DE13.DE16_ID_HORARIO = DE16.DE16_ID_HORARIO")
        strSQL.Append("     And DE13.DE13_NU_DIA_SEMANA = 1 ")
        strSQL.Append("     And DE13.DE13_DH_EXCLUSAO IS NULL) As COD_DOM, ")

        strSQL.Append("(SELECT STRING_AGG('<B>' + DE09_NM_DISCIPLINA + '</B> <BR /> ' + isnull(RH01_NM_PESSOA, '<b style=""color:red;"">SEM PROFESSOR</b>'), '<BR/>') ")
        strSQL.Append("     From dbdiario..DE13_HORARIO_TURMA As DE13      ")
        strSQL.Append("     Left Join dbdiario..DE09_DISCIPLINA AS DE09 ON DE09.DE09_ID_DISCIPLINA = DE13.DE09_ID_DISCIPLINA   ")
        strSQL.Append("     Left Join  RH80_ALOCACAO_CARGA_HORARIA AS RH80 ON RH80.RH80_ID_ALOCACAO_CARGA_HORARIA = DE13.RH80_ID_ALOCACAO_CARGA_HORARIA ")
        strSQL.Append("     Left Join  RH14_LOTACAO_SERVIDOR AS RH14 ON RH14.RH14_ID_LOTACAO_SERVIDOR = RH80.RH14_ID_LOTACAO_SERVIDOR ")
        strSQL.Append("     Left Join  RH02_SERVIDOR AS RH02 ON RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR ")
        strSQL.Append("     Left Join  RH01_PESSOA AS RH01 ON RH01.RH01_ID_PESSOA = RH02.RH01_ID_PESSOA ")
        strSQL.Append("     WHERE DE13.DE10_ID_TURMA = " & Turma)
        strSQL.Append("     And DE13.DE16_ID_HORARIO = DE16.DE16_ID_HORARIO")
        strSQL.Append("     And DE13.DE13_NU_DIA_SEMANA = 1 ")
        strSQL.Append("     And DE13.DE13_DH_EXCLUSAO IS NULL) As DOM, ")
        'SEGUNDA
        strSQL.Append("     (Select top 1 DE13_ID_HORARIO_TURMA ")
        strSQL.Append("     From dbdiario..DE13_HORARIO_TURMA As DE13 ")
        strSQL.Append("     WHERE DE13.DE10_ID_TURMA = " & Turma)
        strSQL.Append("     And DE13.DE16_ID_HORARIO = DE16.DE16_ID_HORARIO")
        strSQL.Append("     And DE13.DE13_NU_DIA_SEMANA = 2 ")
        strSQL.Append("     And DE13.DE13_DH_EXCLUSAO IS NULL) As COD_SEG, ")

        strSQL.Append("     (Select STRING_AGG('<B>' + DE09_NM_DISCIPLINA + '</B> <BR /> ' + isnull(RH01_NM_PESSOA, '<b style=""color:red;"">SEM PROFESSOR</b>'), '<BR/>') ")
        strSQL.Append("     From dbdiario..DE13_HORARIO_TURMA As DE13 ")
        strSQL.Append("     Left Join dbdiario..DE09_DISCIPLINA AS DE09 ON DE09.DE09_ID_DISCIPLINA = DE13.DE09_ID_DISCIPLINA   ")
        strSQL.Append("     Left Join  RH80_ALOCACAO_CARGA_HORARIA AS RH80 ON RH80.RH80_ID_ALOCACAO_CARGA_HORARIA = DE13.RH80_ID_ALOCACAO_CARGA_HORARIA ")
        strSQL.Append("     Left Join  RH14_LOTACAO_SERVIDOR AS RH14 ON RH14.RH14_ID_LOTACAO_SERVIDOR = RH80.RH14_ID_LOTACAO_SERVIDOR ")
        strSQL.Append("     Left Join  RH02_SERVIDOR AS RH02 ON RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR ")
        strSQL.Append("     Left Join  RH01_PESSOA AS RH01 ON RH01.RH01_ID_PESSOA = RH02.RH01_ID_PESSOA ")
        strSQL.Append("     WHERE DE13.DE10_ID_TURMA = " & Turma)
        strSQL.Append("     and DE13.DE16_ID_HORARIO = DE16.DE16_ID_HORARIO")
        strSQL.Append("     And DE13.DE13_NU_DIA_SEMANA = 2 ")
        strSQL.Append("     And DE13.DE13_DH_EXCLUSAO IS NULL) As SEG, ")
        'TERÇA
        strSQL.Append("     (SELECT top 1 DE13_ID_HORARIO_TURMA ")
        strSQL.Append("     From dbdiario..DE13_HORARIO_TURMA As DE13 ")
        strSQL.Append("     WHERE DE13.DE10_ID_TURMA = " & Turma)
        strSQL.Append("     and DE13.DE16_ID_HORARIO = DE16.DE16_ID_HORARIO")
        strSQL.Append("     And DE13.DE13_NU_DIA_SEMANA = 3 ")
        strSQL.Append("     And DE13.DE13_DH_EXCLUSAO IS NULL) As COD_TER, ")

        strSQL.Append("     (SELECT STRING_AGG('<B>' + DE09_NM_DISCIPLINA + '</B> <BR /> ' + isnull(RH01_NM_PESSOA, '<b style=""color:red;"">SEM PROFESSOR</b>'), '<BR/>') ")
        strSQL.Append("     From dbdiario..DE13_HORARIO_TURMA As DE13 ")
        strSQL.Append("     Left Join dbdiario..DE09_DISCIPLINA AS DE09 ON DE09.DE09_ID_DISCIPLINA = DE13.DE09_ID_DISCIPLINA   ")
        strSQL.Append("     Left Join   RH80_ALOCACAO_CARGA_HORARIA AS RH80 ON RH80.RH80_ID_ALOCACAO_CARGA_HORARIA = DE13.RH80_ID_ALOCACAO_CARGA_HORARIA ")
        strSQL.Append("     Left Join   RH14_LOTACAO_SERVIDOR AS RH14 ON RH14.RH14_ID_LOTACAO_SERVIDOR = RH80.RH14_ID_LOTACAO_SERVIDOR ")
        strSQL.Append("     Left Join   RH02_SERVIDOR AS RH02 ON RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR ")
        strSQL.Append("     Left Join   RH01_PESSOA AS RH01 ON RH01.RH01_ID_PESSOA = RH02.RH01_ID_PESSOA ")
        strSQL.Append("     WHERE DE13.DE10_ID_TURMA = " & Turma)
        strSQL.Append("     and DE13.DE16_ID_HORARIO = DE16.DE16_ID_HORARIO")
        strSQL.Append("     And DE13.DE13_NU_DIA_SEMANA = 3 ")
        strSQL.Append("     And DE13.DE13_DH_EXCLUSAO IS NULL) As TER, ")
        'QUARTA
        strSQL.Append("     (SELECT top 1 DE13_ID_HORARIO_TURMA ")
        strSQL.Append("     From dbdiario..DE13_HORARIO_TURMA As DE13 ")
        strSQL.Append("     WHERE DE13.DE10_ID_TURMA = " & Turma)
        strSQL.Append("     and DE13.DE16_ID_HORARIO = DE16.DE16_ID_HORARIO")
        strSQL.Append("     And DE13.DE13_NU_DIA_SEMANA = 4 ")
        strSQL.Append("     And DE13.DE13_DH_EXCLUSAO IS NULL) As COD_QUA, ")

        strSQL.Append("     (SELECT STRING_AGG('<B>' + DE09_NM_DISCIPLINA + '</B> <BR /> ' + isnull(RH01_NM_PESSOA, '<b style=""color:red;"">SEM PROFESSOR</b>'), '<BR/>') ")
        strSQL.Append("     From dbdiario..DE13_HORARIO_TURMA As DE13 ")
        strSQL.Append("     Left Join dbdiario..DE09_DISCIPLINA AS DE09 ON DE09.DE09_ID_DISCIPLINA = DE13.DE09_ID_DISCIPLINA   ")
        strSQL.Append("     Left Join  RH80_ALOCACAO_CARGA_HORARIA AS RH80 ON RH80.RH80_ID_ALOCACAO_CARGA_HORARIA = DE13.RH80_ID_ALOCACAO_CARGA_HORARIA ")
        strSQL.Append("     Left Join  RH14_LOTACAO_SERVIDOR AS RH14 ON RH14.RH14_ID_LOTACAO_SERVIDOR = RH80.RH14_ID_LOTACAO_SERVIDOR ")
        strSQL.Append("     Left Join  RH02_SERVIDOR AS RH02 ON RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR ")
        strSQL.Append("     Left Join  RH01_PESSOA AS RH01 ON RH01.RH01_ID_PESSOA = RH02.RH01_ID_PESSOA ")
        strSQL.Append("     WHERE DE13.DE10_ID_TURMA = " & Turma)
        strSQL.Append("     and DE13.DE16_ID_HORARIO = DE16.DE16_ID_HORARIO")
        strSQL.Append("     And DE13.DE13_NU_DIA_SEMANA = 4 ")
        strSQL.Append("     And DE13.DE13_DH_EXCLUSAO IS NULL) As QUA, ")
        'QUINTA
        strSQL.Append("     (SELECT top 1 DE13_ID_HORARIO_TURMA ")
        strSQL.Append("     From dbdiario..DE13_HORARIO_TURMA As DE13 ")
        strSQL.Append("     WHERE DE13.DE10_ID_TURMA = " & Turma)
        strSQL.Append("     and DE13.DE16_ID_HORARIO = DE16.DE16_ID_HORARIO")
        strSQL.Append("     And DE13.DE13_NU_DIA_SEMANA = 5 ")
        strSQL.Append("     And DE13.DE13_DH_EXCLUSAO IS NULL) As COD_QUI, ")

        strSQL.Append("     (SELECT STRING_AGG('<B>' + DE09_NM_DISCIPLINA + '</B> <BR /> ' + isnull(RH01_NM_PESSOA, '<b style=""color:red;"">SEM PROFESSOR</b>'), '<BR/>') ")
        strSQL.Append("     From dbdiario..DE13_HORARIO_TURMA As DE13 ")
        strSQL.Append("     Left Join dbdiario..DE09_DISCIPLINA AS DE09 ON DE09.DE09_ID_DISCIPLINA = DE13.DE09_ID_DISCIPLINA   ")
        strSQL.Append("     Left Join  RH80_ALOCACAO_CARGA_HORARIA AS RH80 ON RH80.RH80_ID_ALOCACAO_CARGA_HORARIA = DE13.RH80_ID_ALOCACAO_CARGA_HORARIA ")
        strSQL.Append("     Left Join  RH14_LOTACAO_SERVIDOR AS RH14 ON RH14.RH14_ID_LOTACAO_SERVIDOR = RH80.RH14_ID_LOTACAO_SERVIDOR ")
        strSQL.Append("     Left Join  RH02_SERVIDOR AS RH02 ON RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR ")
        strSQL.Append("     Left Join  RH01_PESSOA AS RH01 ON RH01.RH01_ID_PESSOA = RH02.RH01_ID_PESSOA ")
        strSQL.Append("     WHERE DE13.DE10_ID_TURMA = " & Turma)
        strSQL.Append("     and DE13.DE16_ID_HORARIO = DE16.DE16_ID_HORARIO")
        strSQL.Append("     And DE13.DE13_NU_DIA_SEMANA = 5 ")
        strSQL.Append("     And DE13.DE13_DH_EXCLUSAO IS NULL) As QUI, ")
        'SEXTA
        strSQL.Append("     (SELECT top 1 DE13_ID_HORARIO_TURMA ")
        strSQL.Append("     From dbdiario..DE13_HORARIO_TURMA As DE13 ")
        strSQL.Append("     WHERE DE13.DE10_ID_TURMA = " & Turma)
        strSQL.Append("     and DE13.DE16_ID_HORARIO = DE16.DE16_ID_HORARIO")
        strSQL.Append("     And DE13.DE13_NU_DIA_SEMANA = 6 ")
        strSQL.Append("     And DE13.DE13_DH_EXCLUSAO IS NULL) As COD_SEX, ")

        strSQL.Append("     (SELECT STRING_AGG('<B>' + DE09_NM_DISCIPLINA + '</B> <BR /> ' + isnull(RH01_NM_PESSOA, '<b style=""color:red;"">SEM PROFESSOR</b>'), '<BR/>') ")
        strSQL.Append("     From dbdiario..DE13_HORARIO_TURMA As DE13 ")
        strSQL.Append("     Left Join dbdiario..DE09_DISCIPLINA AS DE09 ON DE09.DE09_ID_DISCIPLINA = DE13.DE09_ID_DISCIPLINA   ")
        strSQL.Append("     Left Join RH80_ALOCACAO_CARGA_HORARIA AS RH80 ON RH80.RH80_ID_ALOCACAO_CARGA_HORARIA = DE13.RH80_ID_ALOCACAO_CARGA_HORARIA ")
        strSQL.Append("     Left Join RH14_LOTACAO_SERVIDOR AS RH14 ON RH14.RH14_ID_LOTACAO_SERVIDOR = RH80.RH14_ID_LOTACAO_SERVIDOR ")
        strSQL.Append("     Left Join RH02_SERVIDOR AS RH02 ON RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR ")
        strSQL.Append("     Left Join RH01_PESSOA AS RH01 ON RH01.RH01_ID_PESSOA = RH02.RH01_ID_PESSOA ")
        strSQL.Append("     WHERE DE13.DE10_ID_TURMA = " & Turma)
        strSQL.Append("     and DE13.DE16_ID_HORARIO = DE16.DE16_ID_HORARIO")
        strSQL.Append("     And DE13.DE13_NU_DIA_SEMANA = 6 ")
        strSQL.Append("     And DE13.DE13_DH_EXCLUSAO IS NULL) As SEX, ")
        'SABADO
        strSQL.Append("     (SELECT top 1 DE13_ID_HORARIO_TURMA ")
        strSQL.Append("     From dbdiario..DE13_HORARIO_TURMA As DE13 ")
        strSQL.Append("     WHERE DE13.DE10_ID_TURMA = " & Turma)
        strSQL.Append("     and DE13.DE16_ID_HORARIO = DE16.DE16_ID_HORARIO")
        strSQL.Append("     And DE13.DE13_NU_DIA_SEMANA = 7 ")
        strSQL.Append("     And DE13.DE13_DH_EXCLUSAO IS NULL) As COD_SAB, ")

        strSQL.Append("     (SELECT STRING_AGG('<B>' + DE09_NM_DISCIPLINA + '</B> <BR /> ' + isnull(RH01_NM_PESSOA, '<b style=""color:red;"">SEM PROFESSOR</b>'), '<BR/>') ")
        strSQL.Append("     From dbdiario..DE13_HORARIO_TURMA As DE13 ")
        strSQL.Append("     Left Join dbdiario..DE09_DISCIPLINA AS DE09 ON DE09.DE09_ID_DISCIPLINA = DE13.DE09_ID_DISCIPLINA   ")
        strSQL.Append("     Left Join  RH80_ALOCACAO_CARGA_HORARIA AS RH80 ON RH80.RH80_ID_ALOCACAO_CARGA_HORARIA = DE13.RH80_ID_ALOCACAO_CARGA_HORARIA ")
        strSQL.Append("     Left Join  RH14_LOTACAO_SERVIDOR AS RH14 ON RH14.RH14_ID_LOTACAO_SERVIDOR = RH80.RH14_ID_LOTACAO_SERVIDOR ")
        strSQL.Append("     Left Join  RH02_SERVIDOR AS RH02 ON RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR ")
        strSQL.Append("     Left Join  RH01_PESSOA AS RH01 ON RH01.RH01_ID_PESSOA = RH02.RH01_ID_PESSOA ")
        strSQL.Append("     WHERE DE13.DE10_ID_TURMA = " & Turma)
        strSQL.Append("     and DE13.DE16_ID_HORARIO = DE16.DE16_ID_HORARIO")
        strSQL.Append("     And DE13.DE13_NU_DIA_SEMANA = 7 ")
        strSQL.Append("     And DE13.DE13_DH_EXCLUSAO IS NULL) As SAB")

        strSQL.Append(" From dbdiario..DE16_HORARIO As DE16 ")
        strSQL.Append(" Where DE16.RH36_ID_LOTACAO = " & Lotacao)
        strSQL.Append(" And DE16.DE16_DH_DESATIVACAO is null ")

        If Turno > 0 Then
            strSQL.Append(" And DE16.TG06_ID_TURNO = " & Turno)
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "DE16.DE16_NR_HORA_INICIO, DE16.DE16_NR_MINUTO_INICIO, DE16.DE16_ID_HORARIO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela(Optional Turma As Integer = 0, Optional Pessoa As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select DE13_ID_HORARIO_TURMA As CODIGO,  ")
        strSQL.Append("     Case DE13_NU_DIA_SEMANA ")
        strSQL.Append("     WHEN 1 THEN 'DOMINGO'  ")
        strSQL.Append("     WHEN 2 THEN 'SEGUNDA' ")
        strSQL.Append("     WHEN 3 THEN 'TERCA' ")
        strSQL.Append("     WHEN 4 THEN 'QUARTA' ")
        strSQL.Append("     WHEN 5 THEN 'QUINTA' ")
        strSQL.Append("     WHEN 6 THEN 'SEXTA' ")
        strSQL.Append("     WHEN 7 THEN 'SABADO' ")
        strSQL.Append("     End ")
        strSQL.Append("     + ' das ' ")
        strSQL.Append("     + Convert(varchar, DE16_NR_HORA_INICIO) + ':' +  right('00' + convert(varchar,DE16_NR_MINUTO_INICIO),2)  ")
        strSQL.Append("     + ' às '   ")
        strSQL.Append("     + Convert(varchar, DE16_NR_HORA_TERMINO) + ':' +  right('00' + convert(varchar,DE16_NR_MINUTO_TERMINO),2) + ' - ' + DE09.DE09_NM_DISCIPLINA + ' - (' + ")
        strSQL.Append("     Convert(varchar, ISNULL(RH02_CD_MATRICULA,'')) + ') ' + ISNULL(RH01_NM_PESSOA,'') as DESCRICAO ")
        strSQL.Append(" From DE13_HORARIO_TURMA As DE13 ")
        strSQL.Append(" Left Join DE09_DISCIPLINA As DE09 On DE09.DE09_ID_DISCIPLINA  = DE13.DE09_ID_DISCIPLINA ")
        strSQL.Append(" Left Join DE16_HORARIO As DE16 On DE16.DE16_ID_HORARIO = DE13.DE16_ID_HORARIO ")
        strSQL.Append(" Left Join  RH80_ALOCACAO_CARGA_HORARIA As RH80 On RH80.RH80_ID_ALOCACAO_CARGA_HORARIA = DE13.RH80_ID_ALOCACAO_CARGA_HORARIA ")
        strSQL.Append(" Left Join  RH14_LOTACAO_SERVIDOR As RH14 On RH14.RH14_ID_LOTACAO_SERVIDOR = RH80.RH14_ID_LOTACAO_SERVIDOR ")
        strSQL.Append(" Left Join  RH02_SERVIDOR As RH02 On RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR ")
        strSQL.Append(" Left Join  RH01_PESSOA As RH01 On RH01.RH01_ID_PESSOA = RH02.RH01_ID_PESSOA ")
        strSQL.Append(" where DE13_ID_HORARIO_TURMA Is Not null ")

        If Pessoa > 0 Then
            strSQL.Append(" and DE13_ID_HORARIO_TURMA in ( ")
            strSQL.Append("     Select DISTINCT DE13_ID_HORARIO_TURMA ")
            strSQL.Append("     From DE13_HORARIO_TURMA AS DE13 ")
            strSQL.Append("     Left Join  RH80_ALOCACAO_CARGA_HORARIA AS RH80 ON RH80.RH80_ID_ALOCACAO_CARGA_HORARIA = DE13.RH80_ID_ALOCACAO_CARGA_HORARIA ")
            strSQL.Append("     Left Join  RH14_LOTACAO_SERVIDOR AS RH14 ON RH14.RH14_ID_LOTACAO_SERVIDOR = RH80.RH14_ID_LOTACAO_SERVIDOR ")
            strSQL.Append("     Left Join  RH02_SERVIDOR AS RH02 ON RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR ")
            strSQL.Append("     where RH01_ID_PESSOA = " & Pessoa)
            strSQL.Append("     and DE10_ID_TURMA = " & Turma & ") ")
        End If

        If Turma > 0 Then
            strSQL.Append(" And DE13.DE10_ID_TURMA = " & Turma)
        End If

        strSQL.Append(" order by DE13_NU_DIA_SEMANA, DE16_NR_HORA_INICIO ")

        dt = cnn.AbrirDataTable(strSQL.ToString)


        cnn = Nothing

        Return dt
    End Function

    Public Function ObterHorasNaoAlocadas(ByVal Turma As Integer, ByVal Disciplina As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim HorasNaoAlocadas As Integer

        strSQL.Append(" Select Convert(Integer, DE38_QT_CARGA_HORARIA / DE38_QT_HORA_BASE_SEMANAL) As LIMITE_HORAS_SEMANAIS, ")
        strSQL.Append("     (Select count(DE13_ID_HORARIO_TURMA) ")
        strSQL.Append("     From DE13_HORARIO_TURMA As DE13 ")
        strSQL.Append("     Where DE13.DE09_ID_DISCIPLINA = DE38.DE09_ID_DISCIPLINA ")
        strSQL.Append("     And DE13.DE10_ID_TURMA = DE38.DE10_ID_TURMA ")
        strSQL.Append("     And DE13.DE13_DH_EXCLUSAO is null) As HORAS_ALOCADAS ")
        strSQL.Append(" From DE38_TURMA_DISCIPLINA As DE38 ")
        strSQL.Append(" Where DE38.DE09_ID_DISCIPLINA = " & Disciplina)
        strSQL.Append(" And DE38.DE10_ID_TURMA = " & Turma)

        With cnn.AbrirDataTable(strSQL.ToString)
            If .Rows.Count > 0 Then
                If Not IsDBNull(.Rows(0)(0)) And Not IsDBNull(.Rows(0)(1)) Then
                    HorasNaoAlocadas = .Rows(0)(0) - .Rows(0)(1)
                Else
                    HorasNaoAlocadas = -1
                End If
            Else
                HorasNaoAlocadas = -1
            End If
        End With


        cnn = Nothing

        Return HorasNaoAlocadas
    End Function

    Public Function ObterProfessorHorarioCoincidente(ByVal Pessoa As Integer, ByVal DiaSemana As Integer, ByVal HoraInicio As String, ByVal HoraTermino As String) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim dt As DataTable

        strSQL.Append(" Select DE13.DE13_ID_HORARIO_TURMA, RH48_NM_TIPO_ESCOLA + ' ' + RH36_NM_LOTACAO as ESCOLA, DE10_NU_TURMA as TURMA ")
        'strSQL.Append(" Select DE13.*,DE16_NR_HORA_INICIO, DE16_NR_MINUTO_INICIO, ")
        'strSQL.Append(" DE16_NR_HORA_TERMINO, DE16_NR_MINUTO_TERMINO, Convert(time, Convert(varchar, DE16_NR_HORA_INICIO) + ':' + convert(varchar, DE16_NR_MINUTO_INICIO)), ")
        'strSQL.Append(" Convert(time, Convert(varchar, DE16_NR_HORA_TERMINO) + ':' + convert(varchar, DE16_NR_MINUTO_TERMINO)) ")
        strSQL.Append(" From DE13_HORARIO_TURMA As DE13  ")
        strSQL.Append(" Left Join DE10_TURMA as DE10 on DE10.DE10_ID_TURMA = DE13.DE10_ID_TURMA ")
        strSQL.Append(" Left Join  RH80_ALOCACAO_CARGA_HORARIA AS RH80 ON RH80.RH80_ID_ALOCACAO_CARGA_HORARIA = DE13.RH80_ID_ALOCACAO_CARGA_HORARIA  ")
        strSQL.Append(" Left Join  RH14_LOTACAO_SERVIDOR AS RH14 ON RH14.RH14_ID_LOTACAO_SERVIDOR = RH80.RH14_ID_LOTACAO_SERVIDOR  ")
        strSQL.Append(" Left Join  RH36_LOTACAO AS RH36 ON RH36.RH36_ID_LOTACAO = RH14.RH36_ID_LOTACAO  ")
        strSQL.Append(" Left Join  RH48_TIPO_ESCOLA AS RH48 ON RH48.RH48_ID_TIPO_ESCOLA = RH36.RH48_ID_TIPO_ESCOLA  ")
        strSQL.Append(" Left Join  RH02_SERVIDOR AS RH02 ON RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR  ")
        strSQL.Append(" Left Join  RH01_PESSOA AS RH01 ON RH01.RH01_ID_PESSOA = RH02.RH01_ID_PESSOA  ")
        strSQL.Append(" Left Join DE16_HORARIO as DE16 on DE16.DE16_ID_HORARIO = DE13.DE16_ID_HORARIO ")
        strSQL.Append(" where RH01.RH01_ID_PESSOA = " & Pessoa)
        strSQL.Append(" And DE13_NU_DIA_SEMANA = " & DiaSemana)
        strSQL.Append(" And DE13_DH_EXCLUSAO is null ")
        strSQL.Append(" And ((   ")
        strSQL.Append("     (Convert(time, Convert(varchar, DE16_NR_HORA_INICIO) + ':' + convert(varchar, DE16_NR_MINUTO_INICIO))  < convert(time,'" & HoraInicio & "')   ")
        strSQL.Append("     And Convert(time, Convert(varchar, DE16_NR_HORA_TERMINO) + ':' + convert(varchar, DE16_NR_MINUTO_TERMINO))  > convert(time,'" & HoraInicio & "'))   ")
        strSQL.Append(" Or ")
        strSQL.Append("     (convert(time,convert(varchar, DE16_NR_HORA_INICIO) + ':' + convert(varchar, DE16_NR_MINUTO_INICIO))  < convert(time,'" & HoraTermino & "')   ")
        strSQL.Append("     And Convert(time, Convert(varchar, DE16_NR_HORA_TERMINO) + ':' + convert(varchar, DE16_NR_MINUTO_TERMINO))  > convert(time,'" & HoraTermino & "')) ")
        strSQL.Append(" )  ")
        strSQL.Append(" Or ")
        strSQL.Append("(")
        strSQL.Append("     Convert(time, Convert(varchar, DE16_NR_HORA_INICIO) + ':' + convert(varchar, DE16_NR_MINUTO_INICIO))  = convert(time,'" & HoraInicio & "')   ")
        strSQL.Append("     And Convert(time, Convert(varchar, DE16_NR_HORA_TERMINO) + ':' + convert(varchar, DE16_NR_MINUTO_TERMINO))  = convert(time,'" & HoraTermino & "') ")
        strSQL.Append(" )) ")

        dt = cnn.AbrirDataTable(strSQL.ToString)


        cnn = Nothing

        Return dt
    End Function

    Public Function ObterHorarioTurma(ByVal Turma As Integer, ByVal Horario As Integer, ByVal DiaSemana As Integer) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select DE13.DE13_ID_HORARIO_TURMA, DE13.DE10_ID_TURMA, DE13.RH80_ID_ALOCACAO_CARGA_HORARIA, DE16.DE16_ID_HORARIO ")
        strSQL.Append(" , DE13.DE13_NU_DIA_SEMANA, DE16.DE16_ID_HORARIO, DE09.DE09_ID_DISCIPLINA, DE09.DE09_NM_DISCIPLINA, DE09.DE09_IN_CONDICAO_ESPECIAL ")
        strSQL.Append(" from  DBDIARIO..DE13_HORARIO_TURMA As DE13 ")
        strSQL.Append(" left join DBDIARIO..DE16_HORARIO AS DE16 ON DE16.DE16_ID_HORARIO = DE13.DE16_ID_HORARIO")
        strSQL.Append(" left join DBDIARIO..DE09_DISCIPLINA AS DE09 ON DE09.DE09_ID_DISCIPLINA = DE13.DE09_ID_DISCIPLINA")
        strSQL.Append("  WHERE DE13.DE10_ID_TURMA = " & Turma)
        strSQL.Append("  and DE16.DE16_ID_HORARIO = " & Horario)
        strSQL.Append("  and DE13.DE13_NU_DIA_SEMANA = " & DiaSemana)
        strSQL.Append(" AND DE13.DE13_DH_EXCLUSAO IS NULL ")
        strSQL.Append(" AND DE16.DE16_DH_DESATIVACAO IS NULL ")

        dt = cnn.AbrirDataTable(strSQL.ToString)


        cnn = Nothing

        Return dt
    End Function

    Public Function ObterHorarioProfessorPerfil(ByVal Pessoa As Integer, ByVal AnoPeriodo As Integer) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select RH48.RH48_NM_TIPO_ESCOLA + ' ' + RH36.RH36_NM_LOTACAO AS LOTACAO ")
        strSQL.Append(", Case DE13.DE13_NU_DIA_SEMANA ")
        strSQL.Append("        WHEN 1 THEN 'DOMINGO'  ")
        strSQL.Append("        WHEN 2 THEN 'SEGUNDA' ")
        strSQL.Append("        WHEN 3 THEN 'TERCA' ")
        strSQL.Append("        WHEN 4 THEN 'QUARTA' ")
        strSQL.Append("        WHEN 5 THEN 'QUINTA' ")
        strSQL.Append("        WHEN 6 THEN 'SEXTA' ")
        strSQL.Append("        WHEN 7 THEN 'SABADO' ")
        strSQL.Append("        End As DIA_SEMANA ")
        strSQL.Append(", Convert(varchar, DE16_NR_HORA_INICIO) + ':' +  right('00' + convert(varchar,DE16_NR_MINUTO_INICIO),2) + ' às '  ")
        strSQL.Append(" + Convert(varchar, DE16_NR_HORA_TERMINO) + ':' +  right('00' + convert(varchar,DE16_NR_MINUTO_TERMINO),2) as HORARIO ")
        strSQL.Append(", DE10.DE10_NU_TURMA, TG06.TG06_NM_TURNO, DE09.DE09_NM_DISCIPLINA ")
        strSQL.Append(" From DBDIARIO..DE16_HORARIO As DE16 ")
        strSQL.Append(" Left Join DBGERAL.DBO.TG06_TURNO AS TG06 ON TG06.TG06_ID_TURNO = DE16.TG06_ID_TURNO ")
        strSQL.Append(" Left Join DBDIARIO..DE13_HORARIO_TURMA AS DE13 ON DE13.DE16_ID_HORARIO =  DE16.DE16_ID_HORARIO ")
        strSQL.Append(" Left Join DBDIARIO..DE09_DISCIPLINA AS DE09 ON DE09.DE09_ID_DISCIPLINA = DE13.DE09_ID_DISCIPLINA ")
        strSQL.Append(" Left Join DBDIARIO..DE10_TURMA AS DE10 ON  DE10.DE10_ID_TURMA = DE13.DE10_ID_TURMA ")
        strSQL.Append(" Left Join DBDIARIO..DE04_PERIODO AS DE04 ON DE04.DE04_ID_PERIODO = DE10.DE04_ID_PERIODO ")
        strSQL.Append(" Left Join RH80_ALOCACAO_CARGA_HORARIA AS RH80 ON RH80.RH80_ID_ALOCACAO_CARGA_HORARIA = DE13.RH80_ID_ALOCACAO_CARGA_HORARIA ")
        strSQL.Append(" Left Join RH14_LOTACAO_SERVIDOR AS RH14 ON RH14.RH14_ID_LOTACAO_SERVIDOR = RH80.RH14_ID_LOTACAO_SERVIDOR  ")
        strSQL.Append(" Left Join RH02_SERVIDOR AS RH02 ON RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR ")
        strSQL.Append(" Left Join RH36_LOTACAO AS RH36 ON RH36.RH36_ID_LOTACAO = RH14.RH36_ID_LOTACAO ")
        strSQL.Append(" Left Join RH48_TIPO_ESCOLA as RH48 on RH48.RH48_ID_TIPO_ESCOLA = RH36.RH48_ID_TIPO_ESCOLA ")
        strSQL.Append(" WHERE RH02.RH01_ID_PESSOA = " & Pessoa)
        strSQL.Append(" and Left(DE04.DE04_NM_PERIODO, 4) = " & AnoPeriodo)
        strSQL.Append(" ORDER BY RH36.RH36_NM_LOTACAO, DE13.DE13_NU_DIA_SEMANA, DE16.DE16_NR_HORA_INICIO, DE16.DE16_NR_MINUTO_INICIO, DE16.DE16_ID_HORARIO, TG06.TG06_ID_TURNO ")

        dt = cnn.AbrirDataTable(strSQL.ToString)


        cnn = Nothing

        Return dt
    End Function

    Public Function ObterUltimo() As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" Select max(DE13_ID_HORARIO_TURMA) from DBDIARIO..DE13_HORARIO_TURMA")

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

    Public Sub AtualizarDataHoraExclusao(ByVal Turma As Integer, ByVal Disciplina As Integer, ByVal Usuario As Integer)
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder


        strSQL.Append(" UPDATE DE13_HORARIO_TURMA  ")
        strSQL.Append(" SET CA04_ID_USUARIO_EXC = " & Usuario)
        strSQL.Append(" , DE13_DH_EXCLUSAO = GETDATE() ")
        strSQL.Append(" WHERE DE10_ID_TURMA = " & Turma)
        strSQL.Append(" And DE09_ID_DISCIPLINA = " & Disciplina)
        strSQL.Append(" And ( Select COUNT(DE40_ID_AULA) FROM DE40_AULA ")
        strSQL.Append(" WHERE DE40_AULA.DE13_ID_HORARIO_TURMA = DE13_HORARIO_TURMA.DE13_ID_HORARIO_TURMA) > 0 ")

        cnn.EditarDataTable(strSQL.ToString)


        cnn = Nothing

    End Sub

    Public Sub ExcluirHorarioTurma(ByVal Turma As Integer, ByVal Disciplina As Integer)
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder


        strSQL.Append(" DELETE DE13_HORARIO_TURMA  ")
        strSQL.Append(" WHERE DE10_ID_TURMA = " & Turma)
        strSQL.Append(" And DE09_ID_DISCIPLINA =  " & Disciplina)
        strSQL.Append(" And ( Select COUNT(DE40_ID_AULA) FROM DE40_AULA ")
        strSQL.Append(" WHERE DE40_AULA.DE13_ID_HORARIO_TURMA = DE13_HORARIO_TURMA.DE13_ID_HORARIO_TURMA) = 0 ")

        cnn.EditarDataTable(strSQL.ToString)


        cnn = Nothing

    End Sub
    Public Function Excluir(ByVal Codigo As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from DE13_HORARIO_TURMA")
        strSQL.Append(" where DE13_ID_HORARIO_TURMA = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)


        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class


