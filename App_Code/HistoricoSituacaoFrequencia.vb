Imports Microsoft.VisualBasic
Imports System.Data

Public Class HistoricoSituacaoFrequencia

    Implements  IDisposable
	Private RH43_ID_HISTORICO_SITUACAO_FREQ as Integer
	Private RH24_ID_FREQ_LOTACAO as Integer
	Private RH43_ST_FREQ_LOTACAO as String
	Private RH43_DH_ST_FREQ_LOTACAO as String
	Private CA04_ID_USUARIO as Integer

	Public Property Codigo() as Integer
		Get
			Return RH43_ID_HISTORICO_SITUACAO_FREQ
		End Get
		Set(ByVal Value As Integer)
			RH43_ID_HISTORICO_SITUACAO_FREQ = Value
		End Set
	End Property
	Public Property IdFrequenciaLotacao() as Integer
		Get
			Return RH24_ID_FREQ_LOTACAO
		End Get
		Set(ByVal Value As Integer)
			RH24_ID_FREQ_LOTACAO = Value
		End Set
	End Property
	Public Property SituacaoFrequencia() as String
		Get
			Return RH43_ST_FREQ_LOTACAO
		End Get
		Set(ByVal Value As String)
			RH43_ST_FREQ_LOTACAO = Value
		End Set
	End Property
	Public Property DataHoraSituacaoFrequencia() as String
		Get
			Return RH43_DH_ST_FREQ_LOTACAO
		End Get
		Set(ByVal Value As String)
			RH43_DH_ST_FREQ_LOTACAO = Value
		End Set
	End Property
	Public Property IdUsuario() as Integer
		Get
			Return CA04_ID_USUARIO
		End Get
		Set(ByVal Value As Integer)
			CA04_ID_USUARIO = Value
		End Set
	End Property

	Public Sub New(Optional ByVal Codigo As Integer = 0)
		If  Codigo >  0 Then
			Obter(IdFrequenciaLotacao)
		End If
	End Sub

	Public Sub Salvar()
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim dr As DataRow
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH43_HISTORICO_SITUACAO_FREQ")
		strSQL.Append(" where RH43_ID_HISTORICO_SITUACAO_FREQ = " & Codigo)

		dt = cnn.EditarDataTable(strSQL.ToString)

		If dt.Rows.Count = 0 Then
			dr = dt.NewRow
		Else
			dr = dt.Rows(0)
		End If

		dr("RH24_ID_FREQ_LOTACAO") = ProBanco(RH24_ID_FREQ_LOTACAO, eTipoValor.CHAVE)
		dr("RH43_ST_FREQ_LOTACAO") = ProBanco(RH43_ST_FREQ_LOTACAO, eTipoValor.TEXTO)
		dr("RH43_DH_ST_FREQ_LOTACAO") = ProBanco(RH43_DH_ST_FREQ_LOTACAO, eTipoValor.DATA)
		dr("CA04_ID_USUARIO") = ProBanco(CA04_ID_USUARIO, eTipoValor.CHAVE)

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
		strSQL.Append(" from RH43_HISTORICO_SITUACAO_FREQ")
		strSQL.Append(" where RH43_ID_HISTORICO_SITUACAO_FREQ = " & Codigo)

		dt = cnn.AbrirDataTable(strSQL.ToString)

		If dt.Rows.Count > 0 Then
			dr = dt.Rows(0)
			
			RH43_ID_HISTORICO_SITUACAO_FREQ = DoBanco(dr("RH43_ID_HISTORICO_SITUACAO_FREQ"), eTipoValor.CHAVE)
			RH24_ID_FREQ_LOTACAO = DoBanco(dr("RH24_ID_FREQ_LOTACAO"), eTipoValor.CHAVE)
			RH43_ST_FREQ_LOTACAO = DoBanco(dr("RH43_ST_FREQ_LOTACAO"), eTipoValor.TEXTO)
			RH43_DH_ST_FREQ_LOTACAO = DoBanco(dr("RH43_DH_ST_FREQ_LOTACAO"), eTipoValor.DATA)
			CA04_ID_USUARIO = DoBanco(dr("CA04_ID_USUARIO"), eTipoValor.CHAVE)
		End If

		cnn.FecharBanco()
		cnn = Nothing
	End Sub

	Public Function Pesquisar(Optional ByVal Sort as String = "", Optional Codigo as Integer = 0, Optional IdFrequenciaLotacao as Integer = 0, Optional SituacaoFrequencia as String = "", Optional DataHoraSituacaoFrequencia as String = "", Optional IdUsuario as Integer = 0) as DataTable
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select * ")
		strSQL.Append(" from RH43_HISTORICO_SITUACAO_FREQ")
		'strSQL.Append(" left join tabela on coluna1 = coluna2 ")
		strSQL.Append(" where RH43_ID_HISTORICO_SITUACAO_FREQ is not null")
		
		If Codigo > 0 then 
			strSQL.Append(" and RH43_ID_HISTORICO_SITUACAO_FREQ = " & Codigo)
		End If
		
		If IdFrequenciaLotacao > 0 then 
			strSQL.Append(" and RH24_ID_FREQ_LOTACAO = " & IdFrequenciaLotacao)
		End If
		
		If SituacaoFrequencia <> "" then 
			strSQL.Append(" and upper(RH43_ST_FREQ_LOTACAO) like '%" & SituacaoFrequencia.toUpper & "%'")
		End If
		
		If isDate(DataHoraSituacaoFrequencia) then 
			strSQL.Append(" and RH43_DH_ST_FREQ_LOTACAO = Convert(DateTime, '" & DataHoraSituacaoFrequencia & "', 103)")
		End If
		
		If IdUsuario > 0 then 
			strSQL.Append(" and CA04_ID_USUARIO = " & IdUsuario)
		End If
		
		strSQL.Append(" Order By " & IIf(Sort = "", "RH24_ID_FREQ_LOTACAO", Sort))

		Return cnn.AbrirDataTable(strSQL.ToString)
	End Function

	Public Function ObterTabela() as DataTable
		Dim cnn As New Conexao
		Dim dt As DataTable
		Dim strSQL As New StringBuilder
		
		strSQL.Append(" select RH43_ID_HISTORICO_SITUACAO_FREQ as CODIGO, RH24_ID_FREQ_LOTACAO as DESCRICAO")
		strSQL.Append(" from RH43_HISTORICO_SITUACAO_FREQ")
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
		
		strSQL.Append(" select max(RH24_ID_FREQ_LOTACAO) from RH43_HISTORICO_SITUACAO_FREQ")

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
	Public Function Excluir( ByVal Codigo as String) As Integer
		Dim cnn As New Conexao
		Dim strSQL As New StringBuilder
		Dim LinhasAfetadas As Integer
		
		strSQL.Append(" delete ")
		strSQL.Append(" from RH43_HISTORICO_SITUACAO_FREQ")
		strSQL.Append(" and RH43_ID_HISTORICO_SITUACAO_FREQ = " & Codigo)

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

