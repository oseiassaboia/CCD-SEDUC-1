Imports Microsoft.VisualBasic
Imports System.Data

Public Class Bairro
    Private TG04_ID_BAIRRO As Integer
    Private TG03_ID_MUNICIPIO As Integer
    Private ME02_ID_ZONA As Integer
    Private TG04_NM_BAIRRO As String

    Public Property Codigo() As Integer
        Get
            Return TG04_ID_BAIRRO
        End Get
        Set(ByVal Value As Integer)
            TG04_ID_BAIRRO = Value
        End Set
    End Property
    Public Property Municipio() As Integer
        Get
            Return TG03_ID_MUNICIPIO
        End Get
        Set(ByVal Value As Integer)
            TG03_ID_MUNICIPIO = Value
        End Set
    End Property
    Public Property Zona() As Integer
        Get
            Return ME02_ID_ZONA
        End Get
        Set(ByVal Value As Integer)
            ME02_ID_ZONA = Value
        End Set
    End Property
    Public Property Nome() As String
        Get
            Return TG04_NM_BAIRRO
        End Get
        Set(ByVal Value As String)
            TG04_NM_BAIRRO = Value
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
        strSQL.Append(" from DBGERAL.DBO.TG04_BAIRRO")
        strSQL.Append(" where TG04_ID_BAIRRO = " & Codigo)

        dt = cnn.EditarDataTable(strSQL.ToString)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
        Else
            dr = dt.Rows(0)
        End If

        dr("TG03_ID_MUNICIPIO") = ProBanco(TG03_ID_MUNICIPIO, eTipoValor.CHAVE)
        dr("ME02_ID_ZONA") = ProBanco(ME02_ID_ZONA, eTipoValor.CHAVE)
        dr("TG04_NM_BAIRRO") = ProBanco(TG04_NM_BAIRRO, eTipoValor.TEXTO)

        cnn.SalvarDataTable(dr)

        dt.Dispose()
        dt = Nothing

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Sub Obter(ByVal Codigo As Integer)
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim dr As DataRow
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG04_BAIRRO")
        strSQL.Append(" where TG04_ID_BAIRRO = " & Codigo)

        dt = cnn.AbrirDataTable(strSQL.ToString)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            TG04_ID_BAIRRO = DoBanco(dr("TG04_ID_BAIRRO"), eTipoValor.CHAVE)
            TG03_ID_MUNICIPIO = DoBanco(dr("TG03_ID_MUNICIPIO"), eTipoValor.CHAVE)
            ME02_ID_ZONA = DoBanco(dr("ME02_ID_ZONA"), eTipoValor.CHAVE)
            TG04_NM_BAIRRO = DoBanco(dr("TG04_NM_BAIRRO"), eTipoValor.TEXTO)
        End If

        cnn.FecharBanco()
        cnn = Nothing
    End Sub

    Public Function Pesquisar(Optional ByVal Sort As String = "", Optional Codigo As Integer = 0, Optional Municipio As Integer = 0, Optional Zona As Integer = 0, Optional Nome As String = "") As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select * ")
        strSQL.Append(" from DBGERAL.DBO.TG04_BAIRRO")
        'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
        strSQL.Append(" where TG04_ID_BAIRRO is not null")

        If Codigo > 0 Then
            strSQL.Append(" and TG04_ID_BAIRRO = " & Codigo)
        End If

        If Municipio > 0 Then
            strSQL.Append(" and TG03_ID_MUNICIPIO = " & Municipio)
        End If

        If Zona > 0 Then
            strSQL.Append(" and ME02_ID_ZONA = " & Zona)
        End If

        If Nome <> "" Then
            strSQL.Append(" and upper(TG04_NM_BAIRRO) like '%" & Nome.toUpper & "%'")
        End If

        strSQL.Append(" Order By " & IIf(Sort = "", "TG04_ID_BAIRRO", Sort))

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function PesquisarParaCep(ByVal Municipio As Integer, ByVal Nome As String) As DataTable
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder

        strSQL.Append(" select TG04_ID_BAIRRO, TG03_ID_MUNICIPIO, TG04_NM_BAIRRO ")
        strSQL.Append(" from DBGERAL.DBO.TG04_BAIRRO")
        strSQL.Append(" where TG04_ID_BAIRRO is not null")
        strSQL.Append(" and TG03_ID_MUNICIPIO = " & Municipio)
        strSQL.Append(" and TG04_NM_BAIRRO = '" & Nome.ToUpper & "'")

        Return cnn.AbrirDataTable(strSQL.ToString)
    End Function

    Public Function ObterTabela() As DataTable
        Dim cnn As New Conexao
        Dim dt As DataTable
        Dim strSQL As New StringBuilder

        strSQL.Append(" select TG04_ID_BAIRRO as CODIGO, TG03_ID_MUNICIPIO as DESCRICAO")
        strSQL.Append(" from DBGERAL.DBO.TG04_BAIRRO")
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

        strSQL.Append(" select max(TG04_ID_BAIRRO) from DBGERAL.DBO.TG04_BAIRRO")

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
    Public Function Excluir(ByVal Codigo As Integer) As Integer
        Dim cnn As New Conexao
        Dim strSQL As New StringBuilder
        Dim LinhasAfetadas As Integer

        strSQL.Append(" delete ")
        strSQL.Append(" from DBGERAL.DBO.TG04_BAIRRO")
        strSQL.Append(" where TG04_ID_BAIRRO = " & Codigo)

        LinhasAfetadas = cnn.ExecutarSQL(strSQL.ToString)

        cnn.FecharBanco()
        cnn = Nothing

        Return LinhasAfetadas
    End Function

End Class

