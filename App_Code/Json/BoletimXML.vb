Imports Microsoft.VisualBasic

Public Class BoletimXML
    Private _Disciplina As Integer
    Private _Periodo As Integer
    Private _Etapa As Integer
    Private _Usuario As Integer
    Private _Notas As New List(Of NotaBoletim)

    Public Property Disciplina As Integer
        Get
            Return _Disciplina
        End Get
        Set(ByVal value As Integer)
            _Disciplina = value
        End Set
    End Property
    Public Property Periodo As Integer
        Get
            Return _Periodo
        End Get
        Set(ByVal value As Integer)
            _Periodo = value
        End Set
    End Property
    Public Property Etapa As Integer
        Get
            Return _Etapa
        End Get
        Set(ByVal value As Integer)
            _Etapa = value
        End Set
    End Property
    Public Property Usuario As Integer
        Get
            Return _Usuario
        End Get
        Set(ByVal value As Integer)
            _Usuario = value
        End Set
    End Property

    Public Property Notas As List(Of NotaBoletim)
        Get
            Return _Notas
        End Get
        Set(ByVal value As List(Of NotaBoletim))
            _Notas = value
        End Set
    End Property

End Class

Public Class NotaBoletim

    Private _Aluno As Integer
    Private _CodigoMomento As Integer
    Private _Nota As String

    Public Property Aluno As Integer
        Get
            Return _Aluno
        End Get
        Set(ByVal value As Integer)
            _Aluno = value
        End Set
    End Property

    Public Property CodigoMomento As String
        Get
            Return _CodigoMomento
        End Get
        Set(ByVal value As String)
            _CodigoMomento = value
        End Set
    End Property

    Public Property Nota As String
        Get
            Return _Nota
        End Get
        Set(ByVal value As String)
            _Nota = value
        End Set
    End Property

End Class


