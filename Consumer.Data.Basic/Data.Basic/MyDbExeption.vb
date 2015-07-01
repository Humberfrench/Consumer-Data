Imports System.Runtime.Serialization
Imports System.Data

Namespace Data
    Public Class MyDbExeption
        Inherits DataException
        Implements ISerializable
#Region "Variaveis"
        Dim strMessage As String
        Dim intCode As Integer
#End Region

#Region "Constructor"
        Public Sub New()
            MyBase.New()
            RaiseException("Ocorreu um erro desconhecido acessando os dados.")
        End Sub
        Public Sub New(ByVal exReference As Exception)
            MyBase.New(exReference.Message.ToString())
            RaiseException(exReference.Message.ToString())
        End Sub
        Public Sub New(ByVal Message As String)
            MyBase.New(Message)
            RaiseException(Message)
        End Sub
        Public Sub New(ByVal Message As String, _
                       ByVal innerExeption As Exception)
            MyBase.New(Message, innerExeption)
            RaiseException(Message, InnerException)
        End Sub
        Public Sub New(ByVal Info As SerializationInfo, _
                       ByVal Context As StreamingContext)
            MyBase.New(Info, Context)
        End Sub
#End Region

#Region "Methods"
        Protected Sub RaiseException(ByVal Message As String, _
                                     ByVal innerExeption As Exception)
            strMessage = Message
        End Sub
        Protected Sub RaiseException(ByVal Message As String)
            strMessage = Message
        End Sub
#End Region

#Region "Properties"
        Public ReadOnly Property Code() As Integer
            Get
                Return intCode
            End Get
        End Property
        Public Overrides ReadOnly Property Message() As String
            Get
                If strMessage = "" Then
                    strMessage = MyBase.Message
                End If
                Return strMessage
            End Get
        End Property
#End Region


    End Class

End Namespace