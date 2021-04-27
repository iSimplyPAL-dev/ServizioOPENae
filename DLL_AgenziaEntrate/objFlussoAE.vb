Option Strict On
Imports System
Imports System.Runtime.InteropServices

Namespace AgenziaEntrate
    <Serializable()> _
    Public Class objFlussoAE

        Private _nIdFlusso As Integer = -1
        Private _sCodiceISTAT As String = String.Empty
        Private _sAnno As String = String.Empty
        Private _sNomeFile As String = String.Empty
        Private _sDataEstra As String = String.Empty
        Private _nUtenti As Integer = -1
        Private _nArticoli As Integer = -1
        Private _nRecords As Integer = -1

        Public Property IdFlusso() As Integer
            Get
                Return _nIdFlusso
            End Get
            Set(ByVal Value As Integer)
                _nIdFlusso = Value
            End Set
        End Property

        Public Property CodiceISTAT() As String
            Get
                Return _sCodiceISTAT
            End Get
            Set(ByVal Value As String)
                _sCodiceISTAT = Value
            End Set
        End Property
        Public Property Anno() As String
            Get
                Return _sAnno
            End Get
            Set(ByVal Value As String)
                _sAnno = Value
            End Set
        End Property

        Public Property NomeFile() As String
            Get
                Return _sNomeFile
            End Get
            Set(ByVal Value As String)
                _sNomeFile = Value
            End Set
        End Property

        Public Property DataEstrazione() As String
            Get
                Return _sDataEstra
            End Get
            Set(ByVal Value As String)
                _sDataEstra = Value
            End Set
        End Property

        Public Property NumeroUtenti() As Integer
            Get
                Return _nUtenti
            End Get
            Set(ByVal Value As Integer)
                _nUtenti = Value
            End Set
        End Property

        Public Property NumeroRecords() As Integer
            Get
                Return _nRecords
            End Get
            Set(ByVal Value As Integer)
                _nRecords = Value
            End Set
        End Property

        Public Property NumeroArticoli() As Integer
            Get
                Return _nArticoli
            End Get
            Set(ByVal Value As Integer)
                _nArticoli = Value
            End Set
        End Property
    End Class
End Namespace