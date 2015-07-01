#Region "Imports"
Imports System.Data
Imports System.ComponentModel
Imports System.Globalization
Imports System.Data.Common
#End Region

Namespace Data

    <Description("Classe padrão de Conexão ao banco")> _
    Public Class MyDbConnection
        Implements IDisposable

#Region "Variables"
        Private strConnection As String
        Private strProviderName As String
        Private strKeyName As String
        Private oFactory As DbProviderFactory
        Private oDbConnection As DbConnection
        Private oConnectionString As Security.ConnectionString = Nothing
        Private disposedValue As Boolean        ' To detect redundant calls
#End Region

#Region "Construtores"

        ''' <summary>
        ''' Construtor Básico
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()

            BuildProvider()

        End Sub

        ''' <summary>
        ''' Construtor Básico
        ''' </summary>
        ''' <param name="strMyKeyName">Nome da chave do config que contém a string de conexão.</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal strMyKeyName As String)

            KeyName = strMyKeyName
            BuildProvider()

        End Sub

        Public Sub New(ByVal oConnString As Security.ConnectionString)
            oConnectionString = oConnString
            KeyName = "_defaultByCommand"
            BuildProvider()
        End Sub

        Public Sub New(ByVal oConnection As DbConnection)
            oDbConnection = oConnection
        End Sub

        ''' <summary>
        ''' Construtor. Este método constroi todos os objetos e é chamado pelos construtores acima detalhados
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub BuildProvider()

            Try
                If strKeyName.ToString = "" Then
                    Throw New ArgumentNullException("Propriedade KeyName não está preenchida. Não é possível conectar")
                End If

                'obter conexão
                If KeyName <> "_defaultByCommand" Then
                    oConnectionString = New Security.ConnectionString(strKeyName)
                End If
                'obtendo conexao string
                strConnection = oConnectionString.GetConectionStringForConnect
                'obter provider name
                strProviderName = oConnectionString.ProviderName

                'validando o preenchimento
                If strConnection.ToString = "" Then
                    Throw New DataException("A string de conexão não foi iniciada corretamente")
                End If

                'validando o preenchimento
                If strProviderName.ToString = "" Then
                    Throw New DataException("O Provider name não foi preenchido")
                End If

                'validar provider name, em testes de performance, esta verificação não honerou tempo. 
                'a diferença de usar a verificação ou não para 50 mil linhas é de 27 milisegundos.(0,027 s)
                If Not ValidateProviderName(strProviderName) Then
                    Throw New DataException("O Provider name é inválido ou não está instalado na máquina")
                End If

                oFactory = DbProviderFactories.GetFactory(strProviderName)
                oDbConnection = oFactory.CreateConnection
                oDbConnection.ConnectionString = strConnection
                oDbConnection.Open()

                oConnectionString = Nothing
                disposedValue = False

            Catch dbEx As DataException
                oConnectionString = Nothing
                Throw New DataException(dbEx.Message.ToString())
            Catch NullEx As ArgumentNullException
                oConnectionString = Nothing
                Throw New ArgumentNullException(NullEx.Message.ToString())
            Catch ex As Exception
                oConnectionString = Nothing
                Throw New Exception(ex.Message.ToString())
            End Try
        End Sub

#End Region

#Region "Métodos Privates"

        ''' <summary>
        ''' Efetua a validação do provider solicitado para a conexão. E verifica se o mesmo está ou não instalado
        ''' </summary>
        ''' <param name="strProviderName">Nome do Provider de Dados</param>
        ''' <returns>Booleano</returns>
        ''' <remarks></remarks>
        Private Function ValidateProviderName(ByVal strProviderName As String) As Boolean

            Dim dtProviders As DataTable = Nothing
            Dim blnReturn As Boolean = False

            'carrega os providers instalados
            dtProviders = DbProviderFactories.GetFactoryClasses()
            blnReturn = False

            For Each drProviders As DataRow In dtProviders.Rows
                If drProviders.Item(2).ToString() = strProviderName Then
                    blnReturn = True
                    Exit For
                End If
            Next

            Return blnReturn

        End Function

#End Region

#Region "Métodos Publicos"

        ''' <summary>
        ''' Obtem um DataTable desconectado
        ''' </summary>
        ''' <param name="oCommand">Interface de Command que contém as informações para carregar o DataTable</param>
        ''' <returns>DataTable</returns>
        ''' <remarks></remarks>
        Public Function GetDataTable(ByVal oCommand As DbCommand) As DataTable

            Dim oReader As DbDataReader = Nothing
            Dim oDataTable As DataTable = Nothing

            Try

                oDataTable = New DataTable

                oReader = oCommand.ExecuteReader(CommandBehavior.CloseConnection)
                oDataTable.Load(oReader)
                oReader.Close()
                oReader = Nothing

            Catch ex As Exception
                oDataTable = Nothing
                oReader = Nothing
                Throw New DataException(ex.Message.ToString)
            End Try

            Return oDataTable

        End Function

        ''' <summary>
        ''' Obtem um DataSet desconectado
        ''' </summary>
        ''' <param name="oCommand">Interface de Command que contém as informações para carregar o DataSet</param>
        ''' <returns>DataSet</returns>
        ''' <remarks></remarks>
        Public Function GetDataSet(ByVal oCommand As DbCommand) As DataSet

            Dim oDataTable As DataTable = Nothing

            Try

                oDataTable = GetDataTable(oCommand)

            Catch ex As Exception
                oDataTable = Nothing
                Throw New DataException(ex.Message.ToString)
            End Try

            If IsNothing(oDataTable) Then
                Return Nothing
            Else
                Return oDataTable.DataSet
            End If

        End Function

        ''' <summary>
        ''' Obtém um command do tipo DbCommand
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetNewCommand() As DbCommand
            Return oDbConnection.CreateCommand
        End Function

        Public Function GetNewConnectcion() As DbConnection
            Return oDbConnection
        End Function

        ''' <summary>
        ''' Obtém um Parameter do tipo DbParameter
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetNewParameter() As DbParameter
            Return oFactory.CreateParameter
        End Function

        ''' <summary>
        ''' Obtem um Schema do Banco
        ''' </summary>
        ''' <param name="strCollectionName">O nome do Schema a retornar</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSchema(ByVal strCollectionName As String) As DataTable

            Dim oDataTable As DataTable = Nothing

            Try

                oDataTable = New DataTable
                oDataTable = oDbConnection.GetSchema(strCollectionName)

            Catch ex As Exception
                oDataTable = Nothing
                Throw New DataException(ex.Message.ToString)
            End Try

            Return oDataTable

        End Function

        ''' <summary>
        ''' Obtem um Schema do Banco
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSchema() As DataTable

            Dim oDataTable As DataTable = Nothing

            Try

                oDataTable = New DataTable
                oDataTable = oDbConnection.GetSchema()

            Catch ex As Exception
                oDataTable = Nothing
                Throw New DataException(ex.Message.ToString)
            End Try

            Return oDataTable

        End Function

        ''' <summary>
        ''' Fecha a Conexão
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Close()
            'not changed
            If Me.State <> ConnectionState.Broken AndAlso Me.State <> ConnectionState.Closed Then
                oDbConnection.Close()
            End If
        End Sub


#End Region

#Region "Propriedades"

        Public Property KeyName() As String
            Get
                Return strKeyName
            End Get
            Set(ByVal value As String)
                strKeyName = String.Intern(value)
            End Set
        End Property

        Public Property ConnectionString() As String
            Get
                Return strConnection
            End Get
            Set(ByVal value As String)
                strConnection = String.Intern(value)
            End Set
        End Property

        Public ReadOnly Property Database() As String
            Get
                Return oDbConnection.Database
            End Get
        End Property

        Public ReadOnly Property DataSource() As String
            Get
                Return oDbConnection.DataSource
            End Get
        End Property

        Public ReadOnly Property ServerVersion() As String
            Get
                Return oDbConnection.ServerVersion
            End Get
        End Property

        Public ReadOnly Property State() As System.Data.ConnectionState
            Get
                Return oDbConnection.State
            End Get
        End Property

#End Region

#Region " IDisposable Support "
        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    'free managed resources when explicitly called
                    'If Me.State <> ConnectionState.Broken AndAlso Me.State <> ConnectionState.Closed Then
                    'oDbConnection.Close()
                    'End If
                    'oDbConnection.Dispose()
                    oFactory = Nothing
                    'oDbConnection = Nothing
                End If
                'free shared unmanaged resources
                strConnection = Nothing
                strProviderName = Nothing
                strKeyName = Nothing

            End If
            Me.disposedValue = True
        End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

End Namespace