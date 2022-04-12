Imports Microsoft.VisualBasic
Imports System.Data

Public Class LogFrequenciaServidor

    Implements  IDisposable 
	Private RH27_ID_LOG_FREQ_SERVIDOR as Integer
	Private RH18_ID_FREQ_SERVIDOR as Integer
	Private CA04_ID_USUARIO as Integer
	Private RH27_ST_FREQ_SERVIDOR as String
	Private RH27_DH_FREQ_SERVIDOR as String
	Private RH27_ST_FREQ_LOTACAO as String
	Private RH27_DH_FREQ_LOTACAO as String

	Public Property Codigo() as Integer
		Get
			Return RH27_ID_LOG_FREQ_SERVIDOR
		End Get
		Set(ByVal Value As Integer)
			RH27_ID_LOG_FREQ_SERVIDOR = Value
		End Set
	End Property
	Public Property IdFrequenciaServidor() as Integer
		Get
			Return RH18_ID_FREQ_SERVIDOR
		End Get
		Set(ByVal Value As Integer)
			RH18_ID_FREQ_SERVIDOR = Value
		End Set
	End Property
	Public Property idUsuario() as Integer
		Get
			Return CA04_ID_USUARIO
		End Get
		Set(ByVal Value As Integer)
			CA04_ID_USUARIO = Value
		End Set
	End Property
	Public Property SituacaoFrequencia() as String
		Get
			Return RH27_ST_FREQ_SERVIDOR
		End Get
		Set(ByVal Value As String)
			RH27_ST_FREQ_SERVIDOR = Value
		End Set
	End Property
	Public Property DataHoraFrequenciaServidor() as String
		Get
			Return RH27_DH_FREQ_SERVIDOR
		End Get
		Set(ByVal Value As String)
			RH27_DH_FREQ_SERVIDOR = Value
		End Set
	End Property
	Public Property SituacaoFrequenciaLotacao() as String
		Get
			Return RH27_ST_FREQ_LOTACAO
		End Get
		Set(ByVal Value As String)
			RH27_ST_FREQ_LOTACAO = Value
		End Set
	End Property
	Public Property DataHoraFrequenciaLotacao() as String
		Get
			Return RH27_DH_FREQ_LOTACAO
		End Get
		Set(ByVal Value As String)
			RH27_DH_FREQ_LOTACAO = Value
		End Set
	End Property

	Public Sub New(Optional ByVal Codigo as integer = 0)
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
		strSQL.Append(" from RH27_LOG_FREQ_SERVIDOR")
		strSQL.Append(" where RH27_ID_LOG_FREQ_SERVIDOR = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH18_ID_FREQ_SERVIDOR") = ProBanco(RH18_ID_FREQ_SERVIDOR, eTipoValor.CHAVE)
		dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)
		dr("RH27_ST_FREQ_SERVIDOR") = ProBanco(RH27_ST_FREQ_SERVIDOR, eTipoValor.TEXTO)
		dr("RH27_DH_FREQ_SERVIDOR") = ProBanco(RH27_DH_FREQ_SERVIDOR, eTipoValor.DATA)
		dr("RH27_ST_FREQ_LOTACAO") = ProBanco(RH27_ST_FREQ_LOTACAO, eTipoValor.NUMERO_DECIMAL)
		dr("RH27_DH_FREQ_LOTACAO") = ProBanco(RH27_DH_FREQ_LOTACAO, eTipoValor.DATA)

		cnn.SalvarDataTable(dr)

		dt.Dispose()
		dt = Nothing

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Sub Obter(ByVal Codigo as String)
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH27_LOG_FREQ_SERVIDOR")
		strSQL.Append(" where RH27_ID_LOG_FREQ_SERVIDOR = " & Codigo)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH27_ID_LOG_FREQ_SERVIDOR = DoBanco(dr("RH27_ID_LOG_FREQ_SERVIDOR"), eTipoValor.CHAVE)
			RH18_ID_FREQ_SERVIDOR = DoBanco(dr("RH18_ID_FREQ_SERVIDOR"), eTipoValor.CHAVE)
			CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
			RH27_ST_FREQ_SERVIDOR = DoBanco(dr("RH27_ST_FREQ_SERVIDOR"), eTipoValor.TEXTO)
			RH27_DH_FREQ_SERVIDOR = DoBanco(dr("RH27_DH_FREQ_SERVIDOR"), eTipoValor.DATA)
			RH27_ST_FREQ_LOTACAO = DoBanco(dr("RH27_ST_FREQ_LOTACAO"), eTipoValor.NUMERO_DECIMAL)
			RH27_DH_FREQ_LOTACAO = DoBanco(dr("RH27_DH_FREQ_LOTACAO"), eTipoValor.DATA)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional Codigo as Integer = 0, Optional IdFrequenciaServidor as Integer = 0, Optional idUsuario as Integer = 0, Optional SituacaoFrequencia as String = "", Optional DataHoraFrequenciaServidor as String = "", Optional SituacaoFrequenciaLotacao as String = "", Optional DataHoraFrequenciaLotacao as String = "") as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH27_LOG_FREQ_SERVIDOR")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH27_ID_LOG_FREQ_SERVIDOR is not null")
		
		If Codigo > 0 then 
			strSQL.Append(" and RH27_ID_LOG_FREQ_SERVIDOR = " & Codigo)
		End If
		
		If IdFrequenciaServidor > 0 then 
			strSQL.Append(" and RH18_ID_FREQ_SERVIDOR = " & IdFrequenciaServidor)
		End If
		
		If idUsuario > 0 then 
			strSQL.Append(" and CA04_ID_USUARIO = " & idUsuario)
		End If
		
		If SituacaoFrequencia <> "" then 
			strSQL.Append(" and upper(RH27_ST_FREQ_SERVIDOR) like '%" & SituacaoFrequencia.toUpper & "%'")
		End If
		
		If isDate(DataHoraFrequenciaServidor) then 
			strSQL.Append(" and RH27_DH_FREQ_SERVIDOR = Convert(DateTime, '" & DataHoraFrequenciaServidor & "', 103)")
		End If
		
		If IsNumeric(SituacaoFrequenciaLotacao.Replace(".", "")) then
			strSQL.Append(" and RH27_ST_FREQ_LOTACAO = " & SituacaoFrequenciaLotacao.Replace(".", "").Replace(",", "."))
		End If
		
		If isDate(DataHoraFrequenciaLotacao) then 
			strSQL.Append(" and RH27_DH_FREQ_LOTACAO = Convert(DateTime, '" & DataHoraFrequenciaLotacao & "', 103)")
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH27_ID_LOG_FREQ_SERVIDOR", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH27_ID_LOG_FREQ_SERVIDOR as CODIGO, RH18_ID_FREQ_SERVIDOR as DESCRICAO")
		strSQL.Append(" from RH27_LOG_FREQ_SERVIDOR")
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
		
		strSQL.Append(" select max(RH27_ID_LOG_FREQ_SERVIDOR) from RH27_LOG_FREQ_SERVIDOR")

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
	Public Function Excluir(ByVal Codigo as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH27_LOG_FREQ_SERVIDOR")
		strSQL.Append(" where RH27_ID_LOG_FREQ_SERVIDOR = " & Codigo)

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
         GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

'******************************************************************************
'*                                 20/05/2019                                 *
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

