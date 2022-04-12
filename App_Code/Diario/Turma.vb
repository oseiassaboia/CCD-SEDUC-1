Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Net.Configuration

Public Class Turma

    Private DE10_ID_TURMA As Integer
    Private RH36_ID_LOTACAO As Integer
    Private DE37_ID_TIPO_TURMA As Integer
    Private DE07_ID_ETAPA As Integer
    Private DE75_ID_SUBAREA_ELETIVA As Integer
    Private DE04_ID_PERIODO As Integer
    Private TG06_ID_TURNO As Integer
    Private TG61_ID_SALA As Integer
    Private DE29_ID_TIPO_MEDIACAO As Integer
    Private DE30_ID_TIPO_ATENDIMENTO As Integer
    Private CA04_ID_USUARIO As Integer
    Private DE10_NU_TURMA As String
    Private DE10_NM_TURMA As String
    Private DE10_QT_VAGAS As String
    Private DE10_QT_EXCEDENTES As String
    Private DE10_IN_PROFESSOR_UNICO As String
    Private DE10_DT_INICIO_TURMA As String
    Private DE10_DT_TERMINO_TURMA As String
    Private DE10_DH_CADASTRO As String
    Private ED37_DS_TURMA As String

    Public Property Codigo() As Integer
        Get
            Return DE10_ID_TURMA
        End Get
        Set(ByVal Value As Integer)
            DE10_ID_TURMA = Value
        End Set
    End Property
    Public Property Lotacao() As Integer
        Get
            Return RH36_ID_LOTACAO
        End Get
        Set(ByVal Value As Integer)
            RH36_ID_LOTACAO = Value
        End Set
    End Property
    Public Property TipoTurma() As Integer
        Get
            Return DE37_ID_TIPO_TURMA
        End Get
        Set(ByVal Value As Integer)
            DE37_ID_TIPO_TURMA = Value
        End Set
    End Property
    Public Property Etapa() As Integer
        Get
            Return DE07_ID_ETAPA
        End Get
        Set(ByVal Value As Integer)
            DE07_ID_ETAPA = Value
        End Set
    End Property
    Public Property SubareaEletiva() As Integer
        Get
            Return DE75_ID_SUBAREA_ELETIVA
        End Get
        Set(ByVal Value As Integer)
            DE75_ID_SUBAREA_ELETIVA = Value
        End Set
    End Property
    Public Property Periodo() As Integer
        Get
            Return DE04_ID_PERIODO
        End Get
        Set(ByVal Value As Integer)
            DE04_ID_PERIODO = Value
        End Set
    End Property
    Public Property Turno() As Integer
        Get
            Return TG06_ID_TURNO
        End Get
        Set(ByVal Value As Integer)
            TG06_ID_TURNO = Value
        End Set
    End Property
    Public Property Sala() As Integer
        Get
            Return TG61_ID_SALA
        End Get
        Set(ByVal Value As Integer)
            TG61_ID_SALA = Value
        End Set
    End Property
    Public Property TipoMediacao() As Integer
        Get
            Return DE29_ID_TIPO_MEDIACAO
        End Get
        Set(ByVal Value As Integer)
            DE29_ID_TIPO_MEDIACAO = Value
        End Set
    End Property
    Public Property TipoAtendimento() As Integer
        Get
            Return DE30_ID_TIPO_ATENDIMENTO
        End Get
        Set(ByVal Value As Integer)
            DE30_ID_TIPO_ATENDIMENTO = Value
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
    Public Property Numero() As String
        Get
            Return DE10_NU_TURMA
        End Get
        Set(ByVal Value As String)
            DE10_NU_TURMA = Value
        End Set
    End Property
    Public Property Nome() As String
        Get
            Return DE10_NM_TURMA
        End Get
        Set(ByVal Value As String)
            DE10_NM_TURMA = Value
        End Set
    End Property
    Public Property Vagas() As String
        Get
            Return DE10_QT_VAGAS
        End Get
        Set(ByVal Value As String)
            DE10_QT_VAGAS = Value
        End Set
    End Property
    Public Property Excedentes() As String
        Get
            Return DE10_QT_EXCEDENTES
        End Get
        Set(ByVal Value As String)
            DE10_QT_EXCEDENTES = Value
        End Set
    End Property
    Public Property ProfessorUnico() As String
        Get
            Return DE10_IN_PROFESSOR_UNICO
        End Get
        Set(ByVal Value As String)
            DE10_IN_PROFESSOR_UNICO = Value
        End Set
    End Property
    Public Property DataInicioTurma() As String
        Get
            Return DE10_DT_INICIO_TURMA
        End Get
        Set(ByVal Value As String)
            DE10_DT_INICIO_TURMA = Value
        End Set
    End Property
    Public Property DataTerminoTurma() As String
        Get
            Return DE10_DT_TERMINO_TURMA
        End Get
        Set(ByVal Value As String)
            DE10_DT_TERMINO_TURMA = Value
        End Set
    End Property
    Public Property DataHoraCadastro() As String
        Get
            Return DE10_DH_CADASTRO
        End Get
        Set(ByVal Value As String)
            DE10_DH_CADASTRO = Value
        End Set
    End Property
    Public Property DescricaoAntiga() As String
        Get
            Return ED37_DS_TURMA
        End Get
        Set(ByVal Value As String)
            ED37_DS_TURMA = Value
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
        strSQL.Append(" from DBDIARIO..DE10_TURMA")
        strSQL.Append(" where DE10_ID_TURMA = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("RH36_ID_LOTACAO") = ProBanco(RH36_ID_LOTACAO, eTipoValor.CHAVE)
        dr("DE37_ID_TIPO_TURMA") = ProBanco(DE37_ID_TIPO_TURMA, eTipoValor.CHAVE)
        dr("DE07_ID_ETAPA") = ProBanco(DE07_ID_ETAPA, eTipoValor.CHAVE)
        dr("DE75_ID_SUBAREA_ELETIVA") = ProBanco(DE75_ID_SUBAREA_ELETIVA, eTipoValor.CHAVE)
        dr("DE04_ID_PERIODO") = ProBanco(DE04_ID_PERIODO, eTipoValor.CHAVE)
        dr("TG06_ID_TURNO") = ProBanco(TG06_ID_TURNO, eTipoValor.CHAVE)
        dr("TG61_ID_SALA") = ProBanco(TG61_ID_SALA, eTipoValor.CHAVE)
        dr("DE29_ID_TIPO_MEDIACAO") = ProBanco(DE29_ID_TIPO_MEDIACAO, eTipoValor.CHAVE)
        dr("DE30_ID_TIPO_ATENDIMENTO") = ProBanco(DE30_ID_TIPO_ATENDIMENTO, eTipoValor.CHAVE)
        dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
        dr("DE10_NU_TURMA") = ProBanco(DE10_NU_TURMA, eTipoValor.TEXTO)
        dr("DE10_NM_TURMA") = ProBanco(DE10_NM_TURMA, eTipoValor.TEXTO)
        dr("DE10_QT_VAGAS") = ProBanco(DE10_QT_VAGAS, eTipoValor.NUMERO_INTEIRO)
        dr("DE10_QT_EXCEDENTES") = ProBanco(DE10_QT_EXCEDENTES, eTipoValor.NUMERO_INTEIRO)
        dr("DE10_IN_PROFESSOR_UNICO") = ProBanco(DE10_IN_PROFESSOR_UNICO, eTipoValor.TEXTO)
        dr("DE10_DT_INICIO_TURMA") = ProBanco(DE10_DT_INICIO_TURMA, eTipoValor.DATA)
        dr("DE10_DT_TERMINO_TURMA") = ProBanco(DE10_DT_TERMINO_TURMA, eTipoValor.DATA)
        dr("DE10_DH_CADASTRO") = ProBanco(DE10_DH_CADASTRO, eTipoValor.DATA_COMPLETA)
        dr("ED37_DS_TURMA") = ProBanco(ED37_DS_TURMA, eTipoValor.TEXTO)

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
        strSQL.Append(" from DBDIARIO..DE10_TURMA")
        strSQL.Append(" where DE10_ID_TURMA = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            DE10_ID_TURMA = DoBanco(dr("DE10_ID_TURMA"), eTipoValor.CHAVE)
            RH36_ID_LOTACAO = DoBanco(dr("RH36_ID_LOTACAO"), eTipoValor.CHAVE)
            DE37_ID_TIPO_TURMA = DoBanco(dr("DE37_ID_TIPO_TURMA"), eTipoValor.CHAVE)
            DE07_ID_ETAPA = DoBanco(dr("DE07_ID_ETAPA"), eTipoValor.CHAVE)
            DE75_ID_SUBAREA_ELETIVA = DoBanco(dr("DE75_ID_SUBAREA_ELETIVA"), eTipoValor.CHAVE)
            DE04_ID_PERIODO = DoBanco(dr("DE04_ID_PERIODO"), eTipoValor.CHAVE)
            TG06_ID_TURNO = DoBanco(dr("TG06_ID_TURNO"), eTipoValor.CHAVE)
            TG61_ID_SALA = DoBanco(dr("TG61_ID_SALA"), eTipoValor.CHAVE)
            DE29_ID_TIPO_MEDIACAO = DoBanco(dr("DE29_ID_TIPO_MEDIACAO"), eTipoValor.CHAVE)
            DE30_ID_TIPO_ATENDIMENTO = DoBanco(dr("DE30_ID_TIPO_ATENDIMENTO"), eTipoValor.CHAVE)
            CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
            DE10_NU_TURMA = DoBanco(dr("DE10_NU_TURMA"), eTipoValor.TEXTO)
            DE10_NM_TURMA = DoBanco(dr("DE10_NM_TURMA"), eTipoValor.TEXTO)
            DE10_QT_VAGAS = DoBanco(dr("DE10_QT_VAGAS"), eTipoValor.NUMERO_INTEIRO)
            DE10_QT_EXCEDENTES = DoBanco(dr("DE10_QT_EXCEDENTES"), eTipoValor.NUMERO_INTEIRO)
            DE10_IN_PROFESSOR_UNICO = DoBanco(dr("DE10_IN_PROFESSOR_UNICO"), eTipoValor.TEXTO)
            DE10_DT_INICIO_TURMA = DoBanco(dr("DE10_DT_INICIO_TURMA"), eTipoValor.DATA)
            DE10_DT_TERMINO_TURMA = DoBanco(dr("DE10_DT_TERMINO_TURMA"), eTipoValor.DATA)
            DE10_DH_CADASTRO = DoBanco(dr("DE10_DH_CADASTRO"), eTipoValor.DATA_COMPLETA)
            ED37_DS_TURMA = DoBanco(dr("ED37_DS_TURMA"), eTipoValor.TEXTO)
        End If

    
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional Lotacao As Integer = 0, Optional TipoTurma As Integer = 0, Optional Etapa As Integer = 0, Optional Periodo As Integer = 0, Optional Turno As Integer = 0, Optional Sala As Integer = 0, Optional TipoMediacao As Integer = 0, Optional TipoAtendimento As Integer = 0, Optional Usuario As Integer = 0, Optional Nome As String = "", Optional Vagas As String = "", Optional Excedentes As String = "", Optional ProfessorUnico As String = "", Optional DataInicioTurma As String = "", Optional DataTerminoTurma As String = "", Optional DataHoraCadastro As String = "", Optional Aluno As Integer = 0, Optional AnoPeriodo As Integer = 0, Optional Pessoa As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select DE10.DE10_ID_TURMA, RH36.RH36_NM_LOTACAO, DE04.DE04_NM_PERIODO,  DE05.DE05_NM_MODALIDADE + ' - ' + DE06.DE06_NM_NIVEL + ' - ' + DE07.DE07_NM_ETAPA + isnull(' - ' + DE12_NM_AREA + ' - ' + DE15_NM_CURSO,'')  as DE07_NM_ETAPA,  DE37.DE37_NM_TIPO_TURMA, TG06.TG06_NM_TURNO, TG61.TG61_NM_SALA, DE10.DE10_NU_TURMA, iif(DE10_DT_TERMINO_TURMA is null, 'ABERTO', 'ENCERRADO') AS SITUACAO " & Chr(13))
        strSQL.Append(" from DBDIARIO..DE10_TURMA as DE10" & Chr(13))
        strSQL.Append(" left join DBDIARIO..DE37_TIPO_TURMA as DE37 on DE37.DE37_ID_TIPO_TURMA = DE10.DE37_ID_TIPO_TURMA " & Chr(13))
        strSQL.Append(" left join DBDIARIO..DE07_ETAPA as DE07 on DE07.DE07_ID_ETAPA = DE10.DE07_ID_ETAPA " & Chr(13))
        strSQL.Append(" left join DBDIARIO..DE06_NIVEL as DE06 on DE06.DE06_ID_NIVEL = DE07.DE06_ID_NIVEL " & Chr(13))
        strSQL.Append(" left join DBDIARIO..DE05_MODALIDADE as DE05 on DE05.DE05_ID_MODALIDADE = DE06.DE05_ID_MODALIDADE " & Chr(13))
        strSQL.Append(" left join DBDIARIO..DE15_CURSO as DE15 on DE15.DE15_ID_CURSO = DE07.DE15_ID_CURSO " & Chr(13))
        strSQL.Append(" left join DBDIARIO..DE12_AREA as DE12 on DE12.DE12_ID_AREA = DE15.DE12_ID_AREA " & Chr(13))
        strSQL.Append(" left join DBGERAL.DBO.TG06_TURNO as TG06 on TG06.TG06_ID_TURNO = DE10.TG06_ID_TURNO " & Chr(13))
        strSQL.Append(" left join DBGERAL.DBO.TG61_SALA as TG61 on TG61.TG61_ID_SALA = DE10.TG61_ID_SALA " & Chr(13))
        strSQL.Append(" Left Join DBDIARIO..DE04_PERIODO as DE04 on DE04.DE04_ID_PERIODO = DE10.DE04_ID_PERIODO " & Chr(13))
        strSQL.Append(" Left Join RH36_LOTACAO as RH36 on RH36.RH36_ID_LOTACAO = DE10.RH36_ID_LOTACAO " & Chr(13))
        strSQL.Append(" where DE10_ID_TURMA Is Not null" & Chr(13))

        If Codigo > 0 Then
            strSQL.Append(" And DE10.DE10_ID_TURMA = " & Codigo & Chr(13))
        End If

        If Lotacao > 0 Then
            strSQL.Append(" And DE10.RH36_ID_LOTACAO = " & Lotacao & Chr(13))
        End If

        If TipoTurma > 0 Then
            strSQL.Append(" And DE10.DE37_ID_TIPO_TURMA = " & TipoTurma & Chr(13))
        End If

        If Etapa > 0 Then
            strSQL.Append(" And DE10.DE07_ID_ETAPA = " & Etapa & Chr(13))
        End If

        If Periodo > 0 Then
            'strSQL.Append(" And DE10.DE04_ID_PERIODO = " & Periodo & Chr(13))
            strSQL.Append(" And Left(DE04_NM_PERIODO, 4) = " & Periodo & Chr(13))
        End If

        If Turno > 0 Then
            strSQL.Append(" And DE10.TG06_ID_TURNO = " & Turno & Chr(13))
        End If

        If Sala > 0 Then
            strSQL.Append(" And DE10.TG61_ID_SALA = " & Sala & Chr(13))
        End If

        If TipoMediacao > 0 Then
            strSQL.Append(" And DE10.DE29_ID_TIPO_MEDIACAO = " & TipoMediacao & Chr(13))
        End If

        If TipoAtendimento > 0 Then
            strSQL.Append(" And DE10.DE30_ID_TIPO_ATENDIMENTO = " & TipoAtendimento & Chr(13))
        End If

        If Usuario > 0 Then
            strSQL.Append(" And DE10.CA04_ID_USUARIO = " & Usuario & Chr(13))
        End If

        If Nome <> "" Then
            strSQL.Append(" And upper(DE10.DE10_NU_TURMA) Like '%" & Nome.ToUpper & "%'" & Chr(13))
        End If

        If Vagas <> "" Then
            strSQL.Append(" and upper(DE10.DE10_QT_VAGAS) like '%" & Vagas.ToUpper & "%'" & Chr(13))
        End If

        If Excedentes <> "" Then
            strSQL.Append(" and upper(DE10.DE10_QT_EXCEDENTES) like '%" & Excedentes.ToUpper & "%'" & Chr(13))
        End If

        If ProfessorUnico <> "" Then
            strSQL.Append(" and upper(DE10.DE10_IN_PROFESSOR_UNICO) like '%" & ProfessorUnico.ToUpper & "%'" & Chr(13))
        End If

        If DataInicioTurma <> "" Then
            strSQL.Append(" and upper(DE10.DE10_DT_INICIO_TURMA) like '%" & DataInicioTurma.ToUpper & "%'" & Chr(13))
        End If

        If DataTerminoTurma <> "" Then
            strSQL.Append(" and upper(DE10.DE10_DT_TERMINO_TURMA) like '%" & DataTerminoTurma.ToUpper & "%'" & Chr(13))
        End If

        If IsDate(DataHoraCadastro) Then
            strSQL.Append(" and DE10.DE10_DH_CADASTRO = Convert(DateTime, '" & DataHoraCadastro & "', 103)" & Chr(13))
        End If

        If AnoPeriodo > 0 Then
            strSQL.Append(" and Left(DE04.DE04_NM_PERIODO, 4) = " & AnoPeriodo & Chr(13))
        End If

        'Utilizado no frmAlunoTurmas para exibir as turmas que o aluno esta matriculado
        If Aluno > 0 And AnoPeriodo > 0 Then
            strSQL.Append(" And DE10.DE10_ID_TURMA In (Select distinct DE10.DE10_ID_TURMA  " & Chr(13))
            strSQL.Append("     From DBDIARIO..DE35_MATRICULA As DE35  " & Chr(13))
            strSQL.Append("     Left Join DBDIARIO..DE08_MATRICULA_TURMA As DE08 On DE08.DE35_ID_MATRICULA = DE35.DE35_ID_MATRICULA  " & Chr(13))
            strSQL.Append("     Left Join DBDIARIO..DE10_TURMA As DE10 On DE10.DE10_ID_TURMA = DE08.DE10_ID_TURMA  " & Chr(13))
            strSQL.Append("     Left Join DBDIARIO..DE04_PERIODO As DE04 On DE04.DE04_ID_PERIODO = DE10.DE04_ID_PERIODO  " & Chr(13))
            strSQL.Append("     Left Join DBDIARIO..DE36_SITUACAO_MATRICULA As DE36 On DE36.DE36_ID_SITUACAO_MATRICULA = DE08.DE36_ID_SITUACAO_MATRICULA  " & Chr(13))
            strSQL.Append("     where Left(DE04.DE04_NM_PERIODO, 4) = " & AnoPeriodo & Chr(13))
            strSQL.Append("     And DE35.DE01_ID_ALUNO = " & Aluno & Chr(13))
            strSQL.Append("     And DE36.DE36_IN_CONSIDERA_ENTURMADO = 1) " & Chr(13))
        End If

        If Pessoa > 0 Then
            strSQL.Append(" And DE10.DE10_ID_TURMA In (Select DE10_ID_TURMA " & Chr(13))
            strSQL.Append("     From DBDIARIO..DE13_HORARIO_TURMA As DE13 " & Chr(13))
            strSQL.Append("     Left Join  RH80_ALOCACAO_CARGA_HORARIA As RH80 On RH80.RH80_ID_ALOCACAO_CARGA_HORARIA = DE13.RH80_ID_ALOCACAO_CARGA_HORARIA " & Chr(13))
            strSQL.Append("     Left Join  RH14_LOTACAO_SERVIDOR As RH14 On RH14.RH14_ID_LOTACAO_SERVIDOR = RH80.RH14_ID_LOTACAO_SERVIDOR " & Chr(13))
            strSQL.Append("     Left Join  RH02_SERVIDOR As RH02 On RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR " & Chr(13))
            strSQL.Append("     where RH01_ID_PESSOA =  " & Pessoa & Chr(13))
            strSQL.Append("     and DE13_DH_EXCLUSAO is null)" & Chr(13))
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "DE10.DE10_NU_TURMA", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function QuantidadeVagas(ByVal Turma As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select isnull(DE10.DE10_QT_VAGAS, 0) + isnull(DE10.DE10_QT_EXCEDENTES, 0) - ")
        strSQL.Append("     (Select count(DE08.DE08_ID_MATRICULA_TURMA)       ")
        strSQL.Append("     From DBDIARIO..DE08_MATRICULA_TURMA As DE08 ")
        strSQL.Append("     Left Join DBDIARIO..DE36_SITUACAO_MATRICULA As DE36 On DE36.DE36_ID_SITUACAO_MATRICULA = DE08.DE36_ID_SITUACAO_MATRICULA   ")
        strSQL.Append("     where DE36.DE36_IN_CONSIDERA_ENTURMADO = 1 ")
        strSQL.Append("     And DE08.DE10_ID_TURMA = DE10.DE10_ID_TURMA ")
        strSQL.Append(" And DE36.DE36_IN_CONSIDERA_ENTURMADO = 1) As VAGAS   ")
        strSQL.Append(" From DE10_TURMA As DE10 ")
        strSQL.Append(" Where DE10.DE10_ID_TURMA Is Not null ")
        strSQL.Append(" And DE10.DE10_ID_TURMA = " & Turma)

        Return cnn.AbrirDataTable(strSQL.ToString).Rows(0)("VAGAS")
    End Function

    Public Function NumerarTurma(ByVal Turma As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" update DE08_MATRICULA_TURMA Set DE08_NR_CHAMADA = ORDEM ")
        strSQL.Append(" from( ")
        strSQL.Append("     Select ROW_NUMBER() OVER(ORDER BY DE01.DE01_NM_ALUNO ASC) As ORDEM, DE08.DE08_ID_MATRICULA_TURMA As COD_MATRICULA_TURMA ")
        strSQL.Append("     From DBDIARIO..DE08_MATRICULA_TURMA As DE08 ")
        strSQL.Append("     Left Join DBDIARIO..DE36_SITUACAO_MATRICULA As DE36 On DE36.DE36_ID_SITUACAO_MATRICULA = DE08.DE36_ID_SITUACAO_MATRICULA  ")
        strSQL.Append("     Left Join DBDIARIO..DE35_MATRICULA As DE35 On DE08.DE35_ID_MATRICULA = DE35.DE35_ID_MATRICULA ")
        strSQL.Append("     Left Join DBDIARIO..DE01_ALUNO As DE01 On DE35.DE01_ID_ALUNO = DE01.DE01_ID_ALUNO ")
        strSQL.Append("     where DE08.DE10_ID_TURMA = " & Turma)
        ' strSQL.Append("     And DE36.DE36_IN_CONSIDERA_ENTURMADO = 1 ")
        strSQL.Append("     ) As ORDENACAO ")
        strSQL.Append(" where COD_MATRICULA_TURMA = DE08_ID_MATRICULA_TURMA ")

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn = Nothing

        Return LinhasAfetadas
    End Function

    Public Function LimparNumeracao(ByVal Turma As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" update DE08_MATRICULA_TURMA ")
        strSQL.Append(" Set DBDIARIO..DE08_NR_CHAMADA = null ")
        strSQL.Append(" where DE10_ID_TURMA = " & Turma)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)


        cnn = Nothing

        Return LinhasAfetadas
    End Function


    Public Function TurmaNumerada(ByVal Turma As Integer) As Boolean
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim Numerada As Boolean = False

        strSQL.Append(" Select DE08_ID_MATRICULA_TURMA ")
        strSQL.Append(" From DBDIARIO..DE08_MATRICULA_TURMA ")
        strSQL.Append(" Where DE08_NR_CHAMADA Is Not null ")
        strSQL.Append(" And DE10_ID_TURMA = " & Turma)

        If cnn.AbrirDataTable(strSQL.ToString).Rows.Count > 0 Then
            Numerada = True
        End If


        cnn = Nothing

        Return Numerada
    End Function

    Public Function UltimoNumeroChamada(ByVal Turma As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim dt As DataTable
        Dim UltimoNumero As Integer = 0

        strSQL.Append(" Select top 1 DE08_NR_CHAMADA ")
        strSQL.Append(" From DBDIARIO..DE08_MATRICULA_TURMA ")
        strSQL.Append(" Where DE08_NR_CHAMADA Is Not null ")
        strSQL.Append(" And DE10_ID_TURMA = " & Turma)
        strSQL.Append(" And DE08_NR_CHAMADA > 0")
        strSQL.Append(" order by DE08_NR_CHAMADA desc ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            UltimoNumero = dt.Rows(0)("DE08_NR_CHAMADA")
        End If



        cnn = Nothing

        Return UltimoNumero
    End Function

    Public Function ObterTabela(Optional Lotacao As Integer = 0, Optional AnoPeriodo As Integer = 0, Optional Etapa As Integer = 0, Optional Pessoa As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select DE10.DE10_ID_TURMA As CODIGO, isnull(DE10.DE10_NU_TURMA + ' - ' + DE05.DE05_NM_MODALIDADE + ' - ' + DE06.DE06_NM_NIVEL + ' - ' + DE07.DE07_NM_ETAPA + ' - ' + TG06.TG06_NM_TURNO + isnull(' - ' + DE12_NM_AREA + ' - ' + DE15_NM_CURSO,'') + ' - ' + DE04_NM_PERIODO, DE10.DE10_NM_TURMA)  as DESCRICAO ")
        strSQL.Append(" from DBDIARIO..DE10_TURMA as DE10")
        strSQL.Append(" left join DBDIARIO..DE37_TIPO_TURMA as DE37 on DE37.DE37_ID_TIPO_TURMA = DE10.DE37_ID_TIPO_TURMA ")
        strSQL.Append(" left join DBDIARIO..DE07_ETAPA as DE07 on DE07.DE07_ID_ETAPA = DE10.DE07_ID_ETAPA ")
        strSQL.Append(" left join DBDIARIO..DE06_NIVEL as DE06 on DE06.DE06_ID_NIVEL = DE07.DE06_ID_NIVEL ")
        strSQL.Append(" left join DBDIARIO..DE05_MODALIDADE as DE05 on DE05.DE05_ID_MODALIDADE = DE06.DE05_ID_MODALIDADE ")
        strSQL.Append(" left join DBDIARIO..DE15_CURSO as DE15 on DE15.DE15_ID_CURSO = DE07.DE15_ID_CURSO ")
        strSQL.Append(" left join DBDIARIO..DE12_AREA as DE12 on DE12.DE12_ID_AREA = DE15.DE12_ID_AREA ")
        strSQL.Append(" left join DBGERAL.DBO.TG06_TURNO as TG06 on TG06.TG06_ID_TURNO = DE10.TG06_ID_TURNO ")
        strSQL.Append(" left join DBGERAL.DBO.TG61_SALA as TG61 on TG61.TG61_ID_SALA = DE10.TG61_ID_SALA ")
        strSQL.Append(" Left Join DE04_PERIODO as DE04 on DE04.DE04_ID_PERIODO = DE10.DE04_ID_PERIODO ")
        strSQL.Append(" where DE10_ID_TURMA Is Not null")

        If Lotacao > 0 Then
            strSQL.Append(" And DE10.RH36_ID_LOTACAO = " & Lotacao)
        End If

        If AnoPeriodo > 0 Then
            strSQL.Append(" and Left(DE04.DE04_NM_PERIODO, 4) = " & AnoPeriodo)
        End If

        If Etapa > 0 Then
            strSQL.Append(" And (DE10.DE07_ID_ETAPA = " & Etapa)
            strSQL.Append(" or DE10.DE07_ID_ETAPA is null ) ")
        End If

        If Pessoa > 0 Then
            strSQL.Append(" and DE10.DE10_ID_TURMA in (Select DE10_ID_TURMA ")
            strSQL.Append("     From DBDIARIO..DE13_HORARIO_TURMA AS DE13 ")
            strSQL.Append("     Left Join  RH80_ALOCACAO_CARGA_HORARIA AS RH80 ON RH80.RH80_ID_ALOCACAO_CARGA_HORARIA = DE13.RH80_ID_ALOCACAO_CARGA_HORARIA ")
            strSQL.Append("     Left Join  RH14_LOTACAO_SERVIDOR AS RH14 ON RH14.RH14_ID_LOTACAO_SERVIDOR = RH80.RH14_ID_LOTACAO_SERVIDOR ")
            strSQL.Append("     Left Join  RH02_SERVIDOR AS RH02 ON RH02.RH02_ID_SERVIDOR = RH14.RH02_ID_SERVIDOR ")
            strSQL.Append("     where RH01_ID_PESSOA =  " & Pessoa & ") ")

            strSQL.Append(" And (DE10_DT_TERMINO_TURMA >= getdate() or DE10_DT_TERMINO_TURMA is null) ") 'VERIFICA SE A TURMA JA ESTA FECHADA 
        End If

        strSQL.Append(" order by 2 ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn = Nothing

        Return dt
    End Function

    Public Function teste(ByVal CodigoEscola As Integer) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder
         
        strSQL.Append(" ")
        strSQL.Append(" ")
        strSQL.Append(" ")
        strSQL.Append(" ")


        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn = Nothing

        Return dt

    End Function
    
    Public Function ObterTabelaSecundaria(Optional Lotacao As Integer = 0, Optional AnoPeriodo As Integer = 0) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" Select DE10.DE10_ID_TURMA As CODIGO, isnull(DE10.DE10_NU_TURMA,'') + ' - ' + isnull(DE10.ED37_DS_TURMA,'') as DESCRICAO ")
        strSQL.Append(" from DBDIARIO..DE10_TURMA as DE10")
        strSQL.Append(" Left Join DBDIARIO..DE04_PERIODO as DE04 on DE04.DE04_ID_PERIODO = DE10.DE04_ID_PERIODO ")
        strSQL.Append(" where DE10_ID_TURMA Is Not null")

        If Lotacao > 0 Then
            strSQL.Append(" And DE10.RH36_ID_LOTACAO = " & Lotacao)
        End If

        If AnoPeriodo > 0 Then
            strSQL.Append(" and Left(DE04.DE04_NM_PERIODO, 4) = " & AnoPeriodo)
        End If

        strSQL.Append(" order by 2 ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn = Nothing

        Return dt
    End Function

    Public Function ObterTabelaNomeTurma(ByVal Turma As Integer) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select DE10.DE10_ID_TURMA as CODIGO,  DE10.DE10_NU_TURMA + ' - ' + isnull(DE05.DE05_NM_MODALIDADE + ' - ' + DE06.DE06_NM_NIVEL + ' - ' + DE07.DE07_NM_ETAPA, DE10.DE10_NM_TURMA) + ' - ' + TG06.TG06_NM_TURNO + ' - ' + DE04_NM_PERIODO + ' <b>(' + ")
        strSQL.Append(" case when DE10_DT_TERMINO_TURMA is null then 'ABERTA' else 'ENCERRADA' end + ') </b>'  as DESCRICAO ")
        strSQL.Append(" from DBDIARIO..DE10_TURMA as DE10")
        strSQL.Append(" left join DBDIARIO..DE37_TIPO_TURMA as DE37 on DE37.DE37_ID_TIPO_TURMA = DE10.DE37_ID_TIPO_TURMA ")
        strSQL.Append(" left join DBDIARIO..DE07_ETAPA as DE07 on DE07.DE07_ID_ETAPA = DE10.DE07_ID_ETAPA ")
        strSQL.Append(" left join DBDIARIO..DE06_NIVEL as DE06 on DE06.DE06_ID_NIVEL = DE07.DE06_ID_NIVEL ")
        strSQL.Append(" left join DBDIARIO..DE05_MODALIDADE as DE05 on DE05.DE05_ID_MODALIDADE = DE06.DE05_ID_MODALIDADE ")
        strSQL.Append(" left join DBDIARIO..DBGERAL.DBO.TG06_TURNO as TG06 on TG06.TG06_ID_TURNO = DE10.TG06_ID_TURNO ")
        strSQL.Append(" left join DBDIARIO..DBGERAL.DBO.TG61_SALA as TG61 on TG61.TG61_ID_SALA = DE10.TG61_ID_SALA ")
        strSQL.Append(" Left Join DBDIARIO..DE04_PERIODO as DE04 on DE04.DE04_ID_PERIODO = DE10.DE04_ID_PERIODO ")
        strSQL.Append(" where DE10_ID_TURMA Is Not null")

        strSQL.Append(" And DE10.DE10_ID_TURMA = " & Turma)

        strSQL.Append(" order by 2 ")

        dt = cnn.AbrirDataTable(strSQL.ToString)


        cnn = Nothing

        Return dt
    End Function

    Public Function EncerrarTurma(ByVal Usuario As Integer, ByVal Turma As Integer) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" Declare @P_ID_TURMA                     INT ")
        strSQL.Append(" Declare @P_ID_USUARIO					INT ")
        strSQL.Append(" Declare @P_RET_CD_ERRO					INT ")
        strSQL.Append(" Declare @P_RET_DS_MSG_ERRO_SQLSERVER	VARCHAR(MAX) ")
        strSQL.Append(" Declare @P_RET_DS_MSG_ERRO_CONCEITUAL	VARCHAR(MAX) ")
        strSQL.Append(" Declare @P_RET_NR_LIN_ERRO				INT ")

        strSQL.Append(" Set @P_RET_CD_ERRO = 0  ")
        strSQL.Append(" Set @P_RET_DS_MSG_ERRO_SQLSERVER = ''  ")
        strSQL.Append(" Set @P_RET_DS_MSG_ERRO_CONCEITUAL = ''  ")
        strSQL.Append(" Set @P_RET_NR_LIN_ERRO = 0  ")

        strSQL.Append(" Set @P_ID_TURMA = " & Turma)
        strSQL.Append(" Set @P_ID_USUARIO = " & Usuario)

        strSQL.Append(" exec SP_ENCERRAR_TURMA @P_ID_TURMA, @P_ID_USUARIO, @P_RET_CD_ERRO OUT, @P_RET_DS_MSG_ERRO_SQLSERVER OUT, @P_RET_DS_MSG_ERRO_CONCEITUAL OUT, @P_RET_NR_LIN_ERRO OUT ")

        strSQL.Append(" Select @P_RET_CD_ERRO As CD_ERRO, @P_RET_DS_MSG_ERRO_CONCEITUAL as DS_MSG_ERRO_CONCEITUAL ")

        dt = cnn.AbrirDataTable(strSQL.ToString)


        cnn = Nothing

        Return dt
    End Function

    Public Function ExcluirTurma(ByVal Usuario As Integer, ByVal Turma As Integer) As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" Declare @P_ID_TURMA                     INT ")
        strSQL.Append(" Declare @P_ID_USUARIO					INT ")
        strSQL.Append(" Declare @P_RET_CD_ERRO					INT ")
        strSQL.Append(" Declare @P_RET_DS_MSG_ERRO_SQLSERVER	VARCHAR(MAX) ")
        strSQL.Append(" Declare @P_RET_DS_MSG_ERRO_CONCEITUAL	VARCHAR(MAX) ")
        strSQL.Append(" Declare @P_RET_NR_LIN_ERRO				INT ")

        strSQL.Append(" Set @P_RET_CD_ERRO = 0  ")
        strSQL.Append(" Set @P_RET_DS_MSG_ERRO_SQLSERVER = ''  ")
        strSQL.Append(" Set @P_RET_DS_MSG_ERRO_CONCEITUAL = ''  ")
        strSQL.Append(" Set @P_RET_NR_LIN_ERRO = 0  ")

        strSQL.Append(" Set @P_ID_TURMA = " & Turma)
        strSQL.Append(" Set @P_ID_USUARIO = " & Usuario)

        strSQL.Append(" exec SP_DELETE_DE10_TURMA @P_ID_TURMA, @P_ID_USUARIO, @P_RET_CD_ERRO OUT, @P_RET_DS_MSG_ERRO_SQLSERVER OUT, @P_RET_DS_MSG_ERRO_CONCEITUAL OUT, @P_RET_NR_LIN_ERRO OUT ")

        strSQL.Append(" Select @P_RET_CD_ERRO As CD_ERRO, @P_RET_DS_MSG_ERRO_CONCEITUAL as DS_MSG_ERRO_CONCEITUAL ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn = Nothing

        Return dt
    End Function

    Public Function GerarNumeroTurma(ByVal Lotacao As Integer, ByVal TipoTurma As Integer, ByVal Periodo As Integer, ByVal Turno As Integer, ByVal Etapa As Integer, ByVal Eletiva As Integer, ByVal NumeroTurma As String) As String
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" declare @NUMERO_TURMA varchar(10)  ")


        strSQL.Append(" set @NUMERO_TURMA = dbo.FN_NOVO_CD_TURMA (" & Lotacao & ", " & TipoTurma & ", " & Periodo & ", " & Turno & ", " & IIf(Etapa = 0, "NULL", Etapa) & ", " & IIf(Eletiva = 0, "NULL", Eletiva) & ", " & IIf(NumeroTurma = "", "NULL", "'" & NumeroTurma & "'") & ")")

        strSQL.Append(" Select @NUMERO_TURMA As NUMERO_TURMA ")

        Return cnn.AbrirDataTable(strSQL.ToString).Rows(0)("NUMERO_TURMA")
    End Function

    Public Function VerificarHorarioTurmaSalvo(ByVal Turma As Integer) As Boolean
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim Encontrado As Boolean = False

        strSQL.Append(" select count(DE13_ID_HORARIO_TURMA) from DE13_HORARIO_TURMA ")
        strSQL.Append(" where DBDIARIO..DE10_ID_TURMA = " & Turma)
        strSQL.Append("and DE13_DH_EXCLUSAO is null")

        With cnn.AbrirDataTable(strSQL.ToString)
            If .Rows(0)(0) > 0 Then
                Encontrado = True
            End If
        End With


        cnn = Nothing

        Return Encontrado

    End Function

    Public Function ObterUltimo() As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim CodigoUltimo As Integer

        strSQL.Append(" Select max(DE10_ID_TURMA) from DE10_TURMA")

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

    Public Sub TrocarTurma(ByVal Turma01 As Integer, ByVal Turma02 As Integer)
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" exec SP_TROCA_NU_TURMA " & Turma01 & ", " & Turma02 & "  ")

        cnn.AbrirDataTable(strSQL.ToString)


        cnn = Nothing

    End Sub


    Public Function MigrarTurma(ByVal LotacaoOrigem As Integer, ByVal LotacaoDestino As Integer, ByVal Turmas As String, ByVal Usuario As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer = 0

        strSQL.Append(" BEGIN TRANSACTION; " & Chr(13))
        strSQL.Append(" BEGIN TRY       " & Chr(13))

        strSQL.Append("     update DE13_HORARIO_TURMA Set CA04_ID_USUARIO_EXC = " & Usuario & ", DE13_DH_EXCLUSAO = getdate() " & Chr(13))
        strSQL.Append("     From DE10_TURMA As DE10 " & Chr(13))
        strSQL.Append(" 	Where DE10.DE10_ID_TURMA = DE13_HORARIO_TURMA.DE10_ID_TURMA " & Chr(13))
        strSQL.Append("     And RH36_ID_LOTACAO = " & LotacaoOrigem & Chr(13))
        strSQL.Append(" 	And DE10.DE10_ID_TURMA In (" & Turmas & ") " & Chr(13))

        strSQL.Append("     update DE35_MATRICULA set DE35_MATRICULA.RH36_ID_LOTACAO  = " & LotacaoDestino & Chr(13))
        strSQL.Append("     From DE08_MATRICULA_TURMA   As DE08 " & Chr(13))
        strSQL.Append("     Left Join DE10_TURMA		as DE10 on DE10.DE10_ID_TURMA = DE08.DE10_ID_TURMA" & Chr(13))
        strSQL.Append("     where DE35_MATRICULA.DE35_ID_MATRICULA = DE08.DE35_ID_MATRICULA" & Chr(13))
        strSQL.Append("     And DE10.RH36_ID_LOTACAO = " & LotacaoOrigem & Chr(13))
        strSQL.Append("     And DE10.DE10_ID_TURMA In (" & Turmas & ") " & Chr(13))

        strSQL.Append(" 	update DE10_TURMA Set RH36_ID_LOTACAO = " & LotacaoDestino & ", DE10_DH_ALTERACAO = getdate(), CA04_ID_USUARIO_ALT = " & Usuario & Chr(13))
        strSQL.Append("     where RH36_ID_LOTACAO = " & LotacaoOrigem & Chr(13))
        strSQL.Append(" 	And DE10_ID_TURMA In (" & Turmas & ") " & Chr(13))

        strSQL.Append("         End Try " & Chr(13))
        strSQL.Append("         BEGIN Catch   " & Chr(13))
        strSQL.Append("    Select  " & Chr(13))
        strSQL.Append("         ERROR_NUMBER() As ErrorNumber  " & Chr(13))
        strSQL.Append("         , ERROR_SEVERITY() As ErrorSeverity   " & Chr(13))
        strSQL.Append("         ,ERROR_STATE() As ErrorState   " & Chr(13))
        strSQL.Append("         , ERROR_PROCEDURE() As ErrorProcedure   " & Chr(13))
        strSQL.Append("         ,ERROR_LINE() As ErrorLine   " & Chr(13))
        strSQL.Append("         , ERROR_MESSAGE() As ErrorMessage;   " & Chr(13))
        strSQL.Append(" 	If @@TRANCOUNT > 0   " & Chr(13))
        strSQL.Append("         ROLLBACK TRANSACTION;  " & Chr(13))
        strSQL.Append("         Return " & Chr(13))
        strSQL.Append(" End Catch; " & Chr(13))

        strSQL.Append(" If @@TRANCOUNT > 0   " & Chr(13))
        strSQL.Append(" 	COMMIT TRANSACTION;  " & Chr(13))
        strSQL.Append(" 	Select 0  As ErrorNumber; " & Chr(13))

        With cnn.AbrirDataTable(strSQL.ToString)
            If .Rows.Count > 0 Then
                If .Rows(0)("ErrorNumber") = 0 Then
                    LinhasAfetadas = 1
                End If
            End If
        End With


        cnn = Nothing

        Return LinhasAfetadas
    End Function

    Public Function Excluir(ByVal Codigo As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer = 0

        strSQL.Append(" BEGIN TRANSACTION; ")
        strSQL.Append(" BEGIN Try       ")
        strSQL.Append(" 	delete DE38_TURMA_DISCIPLINA   ")
        strSQL.Append("     where DE10_ID_TURMA = " & Codigo & "; ")

        strSQL.Append("     delete ")
        strSQL.Append("         From DE10_TURMA ")
        strSQL.Append("         Where DE10_ID_TURMA = " & Codigo & "; ")
        strSQL.Append("         End Try ")
        strSQL.Append("         BEGIN Catch   ")
        strSQL.Append("    Select  ")
        strSQL.Append("         ERROR_NUMBER() As ErrorNumber  ")
        strSQL.Append("         , ERROR_SEVERITY() As ErrorSeverity   ")
        strSQL.Append("         ,ERROR_STATE() As ErrorState   ")
        strSQL.Append("         , ERROR_PROCEDURE() As ErrorProcedure   ")
        strSQL.Append("         ,ERROR_LINE() As ErrorLine   ")
        strSQL.Append("         , ERROR_MESSAGE() As ErrorMessage;   ")
        strSQL.Append(" 	If @@TRANCOUNT > 0   ")
        strSQL.Append("         ROLLBACK TRANSACTION;  ")
        strSQL.Append("         Return ")
        strSQL.Append(" End Catch; ")

        strSQL.Append(" If @@TRANCOUNT > 0   ")
        strSQL.Append(" 	COMMIT TRANSACTION;  ")
        strSQL.Append(" 	Select 0  As ErrorNumber; ")

        With cnn.AbrirDataTable(strSQL.ToString)
            If .Rows.Count > 0 Then
                If .Rows(0)("ErrorNumber") = 0 Then
                    LinhasAfetadas = 1
                End If
            End If
        End With


        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class


