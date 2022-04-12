Imports Microsoft.VisualBasic
Imports System.Data

Public Class TipoCargaHoraria
    Implements IDisposable

	Private RH79_ID_TIPO_CARGA_HORARIA as Integer
	Private RH79_NM_TIPO_CARGA_HORARIA as String

	Public Property Id() as Integer
		Get
			Return RH79_ID_TIPO_CARGA_HORARIA
		End Get
		Set(ByVal Value As Integer)
			RH79_ID_TIPO_CARGA_HORARIA = Value
		End Set
	End Property
	Public Property Descricao() as String
		Get
			Return RH79_NM_TIPO_CARGA_HORARIA
		End Get
		Set(ByVal Value As String)
			RH79_NM_TIPO_CARGA_HORARIA = Value
		End Set
	End Property

	Public Sub New(Optional ByVal Id as integer = 0)
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
		strSQL.Append(" from RH79_TIPO_CARGA_HORARIA")
		strSQL.Append(" where RH79_ID_TIPO_CARGA_HORARIA = " & Id)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH79_NM_TIPO_CARGA_HORARIA") = ProBanco(RH79_NM_TIPO_CARGA_HORARIA, eTipoValor.TEXTO)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal Id as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH79_TIPO_CARGA_HORARIA")
		strSQL.Append(" where RH79_ID_TIPO_CARGA_HORARIA = " & Id)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH79_ID_TIPO_CARGA_HORARIA = DoBanco(dr("RH79_ID_TIPO_CARGA_HORARIA"), eTipoValor.CHAVE)
			RH79_NM_TIPO_CARGA_HORARIA = DoBanco(dr("RH79_NM_TIPO_CARGA_HORARIA"), eTipoValor.TEXTO)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional Id as Integer = 0, Optional Descricao as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH79_TIPO_CARGA_HORARIA")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH79_ID_TIPO_CARGA_HORARIA is not null")
		
		If Id > 0 then 
			strSQL.Append(" and RH79_ID_TIPO_CARGA_HORARIA = " & Id)
		End If
		
		If Descricao <> "" then 
			strSQL.Append(" and upper(RH79_NM_TIPO_CARGA_HORARIA) like '%" & Descricao.toUpper & "%'")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH79_ID_TIPO_CARGA_HORARIA", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH79_ID_TIPO_CARGA_HORARIA as CODIGO, RH79_NM_TIPO_CARGA_HORARIA as DESCRICAO")
		strSQL.Append(" from RH79_TIPO_CARGA_HORARIA")
		strSQL.Append(" order by 2 ")

		dt = cnn.AbrirDataTable(strSQL.ToString)

		cnn.FecharBanco()
		cnn = Nothing

		Return dt
	End Function

	Public Function ObterUltimo() as Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim CodigoUltimo As Integer
		
		strSQL.Append(" select max(RH79_ID_TIPO_CARGA_HORARIA) from RH79_TIPO_CARGA_HORARIA")

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
	Public Function Excluir(ByVal Id as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH79_TIPO_CARGA_HORARIA")
		strSQL.Append(" where RH79_ID_TIPO_CARGA_HORARIA = " & Id)

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

