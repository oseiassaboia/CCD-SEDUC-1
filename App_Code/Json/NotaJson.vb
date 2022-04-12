Imports Microsoft.VisualBasic

Public Class InsercaoNota

    Private _CodigoAvaliacao As Integer
    Private _CodigoUsuario As Integer
    Private _ListaNota As New List(Of NotaJson)

    Public Property CodigoUsuario As Integer
        Get
            Return _CodigoUsuario
        End Get
        Set(ByVal value As Integer)
            _CodigoUsuario = value
        End Set
    End Property
    Public Property CodigoAvaliacao As Integer
        Get
            Return _CodigoAvaliacao
        End Get
        Set(ByVal value As Integer)
            _CodigoAvaliacao = value
        End Set
    End Property

    Public Property ListaNota As List(Of NotaJson)
        Get
            Return _ListaNota
        End Get
        Set(ByVal value As List(Of NotaJson))
            _ListaNota = value
        End Set
    End Property

End Class

Public Class NotaJson

    Private _CodigoAluno As Integer
    Private _Nota As Decimal

    Public Property CodigoAluno As Integer
        Get
            Return _CodigoAluno
        End Get
        Set(ByVal value As Integer)
            _CodigoAluno = value
        End Set
    End Property

    Public Property Nota As Decimal
        Get
            Return _Nota
        End Get
        Set(ByVal value As Decimal)
            _Nota = value
        End Set
    End Property
End Class

'Public Class ListaNotas
'    Private _ListaNotas As New List(Of NotaJson)

'    Public Property ListaNotas As List(Of NotaJson)
'        Get
'            Return _ListaNotas
'        End Get
'        Set(ByVal value As List(Of NotaJson))
'            _ListaNotas = value
'        End Set
'    End Property
'End Class

