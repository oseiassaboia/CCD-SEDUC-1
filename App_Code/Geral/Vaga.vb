Imports Microsoft.VisualBasic
Imports System.Data

Public Class Vaga
	Private SE06_COD_VAGA as Integer
	Private FKSE06SE03_COD_CONCURSO as Integer
	Private FKSE06CG01_COD_MUNICIPIO as Integer
    Private FKSE06SE15_COD_CARGO_CONCURSO As Integer
    Private FKSE06CG03_COD_SEXO As Integer
    Private FKSE06SE23_COD_DISCIPLINA As Integer
    Private FKSE06SE17_COD_SUBFASE_ORDENACAO As Integer
    Private FKSE06SE28_COD_POVOADO As Integer
	Private FKSE06SS01_COD_IMOVEL As Integer
    Private SE06_QUANTIDADE as String

	Public Property Codigo() as Integer
		Get
			Return SE06_COD_VAGA
		End Get
		Set(ByVal Value As Integer)
			SE06_COD_VAGA = Value
		End Set
	End Property
	Public Property Concurso() as Integer
		Get
			Return FKSE06SE03_COD_CONCURSO
		End Get
		Set(ByVal Value As Integer)
			FKSE06SE03_COD_CONCURSO = Value
		End Set
	End Property
	Public Property Municipio() as Integer
		Get
			Return FKSE06CG01_COD_MUNICIPIO
		End Get
		Set(ByVal Value As Integer)
			FKSE06CG01_COD_MUNICIPIO = Value
		End Set
	End Property
    Public Property CargoConcurso() As Integer
        Get
            Return FKSE06SE15_COD_CARGO_CONCURSO
        End Get
        Set(ByVal Value As Integer)
            FKSE06SE15_COD_CARGO_CONCURSO = Value
        End Set
    End Property
    Public Property Sexo() As Integer
        Get
            Return FKSE06CG03_COD_SEXO
        End Get
        Set(ByVal Value As Integer)
            FKSE06CG03_COD_SEXO = Value
        End Set
    End Property
    Public Property Disciplina() As Integer
        Get
            Return FKSE06SE23_COD_DISCIPLINA
        End Get
        Set(ByVal Value As Integer)
            FKSE06SE23_COD_DISCIPLINA = Value
        End Set
    End Property
    Public Property SubFaseOrdenacao() As Integer
        Get
            Return FKSE06SE17_COD_SUBFASE_ORDENACAO
        End Get
        Set(ByVal Value As Integer)
            FKSE06SE17_COD_SUBFASE_ORDENACAO = Value
        End Set
    End Property
    Public Property Povoado() As Integer
        Get
            Return FKSE06SE28_COD_POVOADO
        End Get
        Set(ByVal Value As Integer)
            FKSE06SE28_COD_POVOADO = Value
        End Set
    End Property
	Public Property Escola() As Integer
        Get
            Return FKSE06SS01_COD_IMOVEL
        End Get
        Set(value As Integer)
            FKSE06SS01_COD_IMOVEL = value
        End Set
    End Property
    Public Property Quantidade() As String
        Get
            Return SE06_QUANTIDADE
        End Get
        Set(ByVal Value As String)
            SE06_QUANTIDADE = Value
        End Set
    End Property

    Public Sub New(Optional ByVal Codigo as Integer = 0)
		If Codigo > 0 Then
			Obter(Codigo)
		End If
	End Sub

	Public Sub Salvar()
        Dim cnn As New ConexaoSeletivo
        Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from SE06_VAGA")
		strSQL.Append(" where SE06_COD_VAGA = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("FKSE06SE03_COD_CONCURSO") = ProBanco(FKSE06SE03_COD_CONCURSO, eTipoValor.CHAVE)
		dr("FKSE06CG01_COD_MUNICIPIO") = ProBanco(FKSE06CG01_COD_MUNICIPIO, eTipoValor.CHAVE)
        dr("FKSE06SE15_COD_CARGO_CONCURSO") = ProBanco(FKSE06SE15_COD_CARGO_CONCURSO, eTipoValor.CHAVE)
        dr("FKSE06CG03_COD_SEXO") = ProBanco(FKSE06CG03_COD_SEXO, eTipoValor.CHAVE)
        dr("FKSE06SE23_COD_DISCIPLINA") = ProBanco(FKSE06SE23_COD_DISCIPLINA, eTipoValor.CHAVE)
        dr("FKSE06SE17_COD_SUBFASE_ORDENACAO") = ProBanco(FKSE06SE17_COD_SUBFASE_ORDENACAO, eTipoValor.CHAVE)
        dr("FKSE06SE28_COD_POVOADO") = ProBanco(FKSE06SE28_COD_POVOADO, eTipoValor.CHAVE)
		dr("FKSE06SS01_COD_IMOVEL") = ProBanco(FKSE06SS01_COD_IMOVEL, eTipoValor.CHAVE)
        dr("SE06_QUANTIDADE") = ProBanco(SE06_QUANTIDADE, eTipoValor.NUMERO_INTEIRO)

        cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

        cnn = Nothing
    End Sub

	Public Sub Obter(ByVal Codigo as Integer)
        Dim cnn As New ConexaoSeletivo
        Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
        strSQL.Append(" from SE06_VAGA")
        strSQL.Append(" where SE06_COD_VAGA = " & Codigo)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			SE06_COD_VAGA = DoBanco(dr("SE06_COD_VAGA"), eTipoValor.CHAVE)
			FKSE06SE03_COD_CONCURSO = DoBanco(dr("FKSE06SE03_COD_CONCURSO"), eTipoValor.CHAVE)
			FKSE06CG01_COD_MUNICIPIO = DoBanco(dr("FKSE06CG01_COD_MUNICIPIO"), eTipoValor.CHAVE)
            FKSE06SE15_COD_CARGO_CONCURSO = DoBanco(dr("FKSE06SE15_COD_CARGO_CONCURSO"), eTipoValor.CHAVE)
            FKSE06CG03_COD_SEXO = DoBanco(dr("FKSE06CG03_COD_SEXO"), eTipoValor.CHAVE)
            FKSE06SE23_COD_DISCIPLINA = DoBanco(dr("FKSE06SE23_COD_DISCIPLINA"), eTipoValor.CHAVE)
            FKSE06SE17_COD_SUBFASE_ORDENACAO = DoBanco(dr("FKSE06SE17_COD_SUBFASE_ORDENACAO"), eTipoValor.CHAVE)
            FKSE06SE28_COD_POVOADO = DoBanco(dr("FKSE06SE28_COD_POVOADO"), eTipoValor.CHAVE)
			FKSE06SS01_COD_IMOVEL = DoBanco(dr("FKSE06SS01_COD_IMOVEL"), eTipoValor.CHAVE)
            SE06_QUANTIDADE = DoBanco(dr("SE06_QUANTIDADE"), eTipoValor.NUMERO_INTEIRO)
        End If

        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional Concurso As Integer = 0, Optional Municipio As Integer = 0, Optional CargoConcurso As Integer = 0, Optional Quantidade As String = "", Optional Candidato As Integer = 0, Optional LocalizarVaga As String = "", optional Disciplina as integer=0, Optional Escola As Integer = 0) As DataTable
        Dim cnn As New ConexaoSeletivo
        Dim strSQL As New StringBuilder

        strSQL.Append(" select *,  VW07_SS01_IMOVEL.DESCRICAO  as ESCOLA, case ISNULL(CG03_COD_SEXO,0) when 0 then 'AMBOS OS SEXOS' else CG03_DESCRICAO END AS SEXO, ")
        strSQL.Append(" SE04_DESCRICAO + ' - ' + CG01_NOME + ")
        strSQL.Append(" case ISNULL(SE28_COD_POVOADO,0) when 0 then '' else ' - POVOADO: ' + SE28_DES_POVOADO END + ")
        strSQL.Append(" case ISNULL(SE23_COD_DISCIPLINA,0) when 0 then '' else ' - ' + SE23_DESCRICAO END + ")
        strSQL.Append(" case ISNULL(CG03_COD_SEXO,0) when 0 then ' - AMBOS OS SEXOS' else ' - ' + CG03_DESCRICAO END AS VAGA ")


        strSQL.Append(" ,(select count(*) ")
        strSQL.Append(" From SE02_INSCRICAO ")
        strSQL.Append(" Where FKSE02SE06_COD_VAGA = SE06_COD_VAGA ")
        strSQL.Append(" And SE02_COD_INSCRICAO In (  ")
        strSQL.Append("     Select FKSE14SE02_COD_INSCRICAO ")
        strSQL.Append("     From SE14_PONTO_INSCRICAO ")
        strSQL.Append("     Where SE14_VALIDA = 1) ")
        strSQL.Append(" and SE02_DATA_CANCELAMENTO is null ")
        strSQL.Append(" ) As QUANTIDADE ")

        If Candidato > 0 Then
            strSQL.Append(", Right('000000000000' + convert(varchar,( select SE02_COD_INSCRICAO ")
            strSQL.Append("     From SE02_INSCRICAO ")
            strSQL.Append("     Where FKSE02SE06_COD_VAGA = SE06_COD_VAGA ")
            strSQL.Append(" and SE02_DATA_CANCELAMENTO is null ")
            strSQL.Append("     and FKSE02SE01_COD_CANDIDATO = " & Candidato & ")),10)as INSCRICAO ")
        End If

        strSQL.Append(" , (SELECT COUNT(SE02_COD_INSCRICAO)  ")
        strSQL.Append(" FROM VW05_INSCRICAO_ULTIMA_CLASSIFICACAO ")
        strSQL.Append(" LEFT JOIN SE24_CLASSIFICACAO ON CLASSIFICACAO = SE24_COD_CLASSIFICACAO ")
        strSQL.Append(" WHERE FKSE24SE12_COD_STATUS_CANDIDATO_FASE = 2 ")
        strSQL.Append(" And SE02_COD_INSCRICAO not in ( ")
        strSQL.Append("     Select FKSE13SE02_COD_INSCRICAO ")
        strSQL.Append("     from SE13_STATUS_INSCRICAO ")
        strSQL.Append("     where FKSE13SE12_COD_STATUS in (4,5) ")
        strSQL.Append("     and SE13_DATA_EXCLUSAO is null ) ")
        strSQL.Append(" and SE02_DATA_CANCELAMENTO is null ")
        strSQL.Append(" AND FKSE02SE06_COD_VAGA =  SE06_COD_VAGA) AS QUANTIDADE_CLASSIFICADOS ")

        strSQL.Append(" from SE06_VAGA ")
        strSQL.Append(" Left Join CG01_MUNICIPIO On FKSE06CG01_COD_MUNICIPIO = CG01_COD_MUNICIPIO ")
        strSQL.Append(" Left Join SE15_CARGO_CONCURSO On FKSE06SE15_COD_CARGO_CONCURSO = SE15_COD_CARGO_CONCURSO ")
        strSQL.Append(" Left Join SE04_CARGO On FKSE15SE04_COD_CARGO = SE04_COD_CARGO ")
        strSQL.Append(" Left Join CG03_SEXO On FKSE06CG03_COD_SEXO = CG03_COD_SEXO ")
        strSQL.Append(" left join SE23_DISCIPLINA On FKSE06SE23_COD_DISCIPLINA = SE23_COD_DISCIPLINA ")
        strSQL.Append(" left join SE28_POVOADO on FKSE06SE28_COD_POVOADO = SE28_COD_POVOADO ")
        strSQL.Append("left join VW07_SS01_IMOVEL on CODIGO = FKSE06SS01_COD_IMOVEL")
        strSQL.Append(" where SE06_COD_VAGA Is Not null")

        If Codigo > 0 Then
            strSQL.Append(" And SE06_COD_VAGA = " & Codigo)
        End If

        If Concurso > 0 Then
            strSQL.Append(" And FKSE06SE03_COD_CONCURSO = " & Concurso)
        End If

        If Municipio > 0 Then
            strSQL.Append(" And FKSE06CG01_COD_MUNICIPIO = " & Municipio)
        End If
		
		If Escola > 0 Then
            strSQL.Append(" And FKSE06SS01_COD_IMOVEL = " & Escola)
        End If

        If CargoConcurso > 0 Then
            strSQL.Append(" And FKSE06SE15_COD_CARGO_CONCURSO = " & CargoConcurso)
        End If

        If Quantidade <> "" Then
            strSQL.Append(" And upper(SE06_QUANTIDADE) Like '%" & Quantidade.ToUpper & "%'")
        End If

        If LocalizarVaga <> "" Then
            strSQL.Append(" And upper(SE04_DESCRICAO + ISNULL(' ' + SE23_DESCRICAO,'') + ' - ' + CG01_NOME + ' - ' +  case ISNULL(CG03_COD_SEXO,0) when 0 then 'AMBOS OS SEXOS' else CG03_DESCRICAO END) Like '%" & LocalizarVaga.ToUpper & "%'")
        End If
		
		  If Disciplina > 0 Then
            strSQL.Append(" And FKSE06SE23_COD_DISCIPLINA = " & Disciplina)
        End If

	strSQL.Append(" And SE06_COD_VAGA not in (4226,4194,4272, 4310, 4208, 4251, 4293, 4175, 4319, 4242, 4297)")

        strSQL.Append(" Order By " & IIf(Sort = "", "SE06_COD_VAGA", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela(ByVal CONCURSO As Integer) As DataTable
        Dim cnn As New ConexaoSeletivo
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select SE06_COD_VAGA as CODIGO, CG01_NOME + ' - ' + SE04_DESCRICAO + ISNULL(' ' + SE23_DESCRICAO,'') + ' - ' + case ISNULL(CG03_COD_SEXO,0) when 0 then 'AMBOS OS SEXOS' else CG03_DESCRICAO END as DESCRICAO")
        strSQL.Append(" from SE06_VAGA")
        strSQL.Append(" left join SE23_DISCIPLINA on FKSE06SE23_COD_DISCIPLINA = SE23_COD_DISCIPLINA ")
        strSQL.Append(" Left Join CG01_MUNICIPIO on FKSE06CG01_COD_MUNICIPIO = CG01_COD_MUNICIPIO ")
        strSQL.Append(" Left Join SE15_CARGO_CONCURSO on FKSE06SE15_COD_CARGO_CONCURSO = SE15_COD_CARGO_CONCURSO ")
        strSQL.Append(" Left Join SE04_CARGO on FKSE15SE04_COD_CARGO = SE04_COD_CARGO ")
        strSQL.Append(" Left Join CG03_SEXO on FKSE06CG03_COD_SEXO = CG03_COD_SEXO ")
        strSQL.Append(" WHERE FKSE06SE03_COD_CONCURSO = " & CONCURSO)
        strSQL.Append(" order by 1 ")

        dt = cnn.AbrirDataTable(strSQL.ToString)

        cnn = Nothing

        Return dt
    End Function

    Public Function ObterUltimo() as Integer
        Dim cnn As New ConexaoSeletivo
        Dim strSQL As New StringBuilder
		Dim CodigoUltimo As Integer
		
		strSQL.Append(" Select max(SE06_COD_VAGA) from SE06_VAGA")

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
	Public Function Excluir(ByVal Codigo as Integer) As Integer
        Dim cnn As New ConexaoSeletivo
        Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from SE06_VAGA")
		strSQL.Append(" where SE06_COD_VAGA = " & Codigo)

		LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)


        cnn = Nothing

		Return LinhasAfetadas
	End Function

End Class

'******************************************************************************
'*                                 08/03/2016                                 *
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

