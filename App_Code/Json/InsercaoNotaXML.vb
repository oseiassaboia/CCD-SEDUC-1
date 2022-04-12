Imports Microsoft.VisualBasic

Public Class ROOT

    Private _AV As Integer
    Private _US As Integer
    Private _NTS As New List(Of NT)

    Public Property US As Integer
        Get
            Return _US
        End Get
        Set(ByVal value As Integer)
            _US = value
        End Set
    End Property
    Public Property AV As Integer
        Get
            Return _AV
        End Get
        Set(ByVal value As Integer)
            _AV = value
        End Set
    End Property

    Public Property NTS As List(Of NT)
        Get
            Return _NTS
        End Get
        Set(ByVal value As List(Of NT))
            _NTS = value
        End Set
    End Property

End Class

Public Class NT

    Private _AL As Integer
    Private _VL As String

    Public Property AL As Integer
        Get
            Return _AL
        End Get
        Set(ByVal value As Integer)
            _AL = value
        End Set
    End Property

    Public Property VL As String
        Get
            Return _VL
        End Get
        Set(ByVal value As String)
            _VL = value
        End Set
    End Property

End Class
