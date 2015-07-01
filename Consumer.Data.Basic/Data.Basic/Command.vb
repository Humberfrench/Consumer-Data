Imports System.Data
Imports System.Data.Common
Imports System.Collections.Generic
Imports System.Text
Imports Consumer.Data.Basic.Security

Namespace Data
    Public Class Command
        Implements ICommand

#Region "Variables"

        Private disposedValue As Boolean = False        ' To detect redundant calls
        Private stbCmdText As StringBuilder = Nothing
        Private oParameters As List(Of Parameter) = Nothing
        Private intCommandType As Integer = 0
        Private intCommandTimeout As Integer = 0
        Private blnPrepared As Boolean = False
        Private strCmdText As String = String.Empty
        Private strKeyName As String = String.Empty
        Private strConnectionString As String = String.Empty
        Private oConnStringData As ConnectionString = Nothing

#End Region

#Region "Construtores"

        ''' <summary>
        ''' Construtor do Command
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub New()

            disposedValue = False
            oParameters = New List(Of Parameter)
            stbCmdText = New StringBuilder
            strCmdText = ""
            strConnectionString = ""
            intCommandTimeout = 360
            intCommandType = CommandType.Text
            blnPrepared = False

        End Sub

        ''' <summary>
        ''' Construtor do Command
        ''' </summary>
        ''' <param name="strKey">Key da String de Conexão</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal strKey As String)

            Me.new()
            strKeyName = strKey
            blnPrepared = False

        End Sub

        ''' <summary>
        ''' Construtor do Command
        ''' </summary>
        ''' <param name="cmdText">Comando a ser executado</param>
        ''' <param name="cmdType">Tipo de Command, se executa uma Stored Procedure ou SQL Statment</param>
        ''' <param name="strKey">Chave do Config</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal cmdText As String, _
                       ByVal cmdType As CommandType, _
                       ByVal strKey As String)

            Me.new()
            strKeyName = strKey
            strCmdText = cmdText
            stbCmdText.Append(strCmdText)
            intCommandType = cmdType

        End Sub

        Public Sub New(ByVal oConnString As ConnectionString)
            Me.new()
            strKeyName = "_defaultByCommand"
            strConnectionString = oConnString.GetConectionStringForConnect
            oConnStringData = oConnString
        End Sub

        ''' <summary>
        ''' Construtor do Command - Generic.
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub BuildCommand()

            'disposedValue = False
            'oParameters = New List(Of Parameter)
            'stbCmdText = New StringBuilder
            'strCmdText = ""
            'strConnectionString = ""
            'intCommandTimeout = 360
            'intCommandType = CommandType.Text
            'blnPrepared = False

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Esta Sub executa a verificação dos parametros, se estão preenchidos<BR/>
        ''' Estes paramtros são preenchidos alimentando as propriedades abaixo.<BR/> 
        ''' Devem ser preenchidas TODAS:
        ''' <list type="bullet">
        '''<item>
        '''<description>Propriedade Key</description>
        '''</item>
        '''<item>
        '''<description>Propriedade CommandText</description>
        '''</item>
        '''<item>
        '''<description>Propriedade CommandType</description>
        '''</item>
        ''' </list>
        ''' <seealso cref="Command.Key" />
        ''' <seealso cref="Command.CommandText" />
        ''' <seealso cref="Command.CommandTimeout" />
        ''' <seealso cref="Command.CommandType" />
        ''' <seealso cref="Command.GetDataTable" />
        ''' <seealso cref="Command.GetDataSet" />
        ''' <seealso cref="Command.GetDataReader" />
        ''' <seealso cref="Command.Execute" />
        ''' <seealso cref="Command.ExecuteScalar" />
        ''' <seealso cref="Command.ExecuteNonQuery" />
        '''</summary>
        ''' <remarks>
        ''' Na verificação, caos não exista o dado preenchido é gerado um erro e este tratado nas funções
        ''' <list type="bullet">
        '''<item>
        '''<description>ReturnDataSet</description>
        '''</item>
        '''<item>
        '''<description>ReturnDataTable</description>
        '''</item>
        ''' </list>
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Protected Sub ValidateCommand()

            'esta função é importante antes de renderizar o command
            'caso falte algo importante, ele dedura e retorna ERRO

            If Me.Key = Nothing Then
                Throw New ArgumentNullException(NO_KEY_SET)
            ElseIf Me.Key = "_defaultByCommand" Then
                If oConnStringData Is Nothing Then
                    Throw New ArgumentNullException(NO_OBJCONN_SET)
                End If
            End If

            'propriedade CommandText
            If Me.CommandText = Nothing Then
                Throw New ArgumentNullException(NO_COMMANDTEXT_SET)
            End If

            'propriedade CommandType
            If Me.CommandType = Nothing Then
                Throw New ArgumentNullException(NO_COMMANDTYPE_SET)
            End If

        End Sub

#End Region

#Region "Métodos Públicos"

#Region "CreateParameter"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Cria um novo parametro.
        ''' Pode se criar um parametro de duas formas conforme exemplo
        ''' <example>
        ''' <code>
        ''' 'Como criar parametro - forma 1:
        ''' Dim parData as new Parameter
        ''' <BR/>
        ''' 'Como criar parametro - forma 2, identico a forma 1:
        ''' Dim parData as Parameter
        ''' parData = new Parameter
        ''' <BR/>
        ''' 'Como criar parametro - forma 3:
        ''' dim parData as Parameter
        ''' dim cmdData as new Command
        ''' parData = cmdData.CreateParameter
        ''' </code>
        ''' </example>
        ''' </summary>
        ''' <returns>Um objeto do tipo Parameter</returns>
        ''' <remarks>
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Function CreateParameter() As Parameter Implements ICommand.CreateParameter
            Return New Parameter
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Cria um novo parametro para ser usado e adiciona ao command atual
        ''' </summary>
        ''' <param name="Name">nome do Parametro. Tipo String</param>
        ''' <param name="Value">Valor que o parametro recebe. Tipo Objeto.</param>
        ''' <param name="Type">Tipo do Parametro. Tipo de dado, proveniente de enumerador</param>
        ''' <param name="Size">Tamanho do Parametro. Para Strings, o tamanho do char/varchar</param>
        ''' <remarks>
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Sub CreateParameter(ByVal Name As String, _
                                   ByVal Type As DbType, _
                                   ByVal Size As Integer, _
                                   ByVal Value As Object) Implements ICommand.CreateParameter

            Dim oParameter As Parameter
            oParameter = New Parameter(Name, Type, Size, Value)
            Me.Parameters.Add(oParameter)

        End Sub

#End Region

#Region "DataTable"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Retorna um datatable.<BR/>
        ''' Para este Método funcionar correramente é preciso que alguns parametros estejam preenchidos corretamente.<BR/>
        ''' Estes paramtros são preenchidos alimentando as propriedades abaixo. <BR/>
        ''' Devem ser preenchidas TODAS:
        ''' <list type="bullet">
        '''<item>
        ''' <term>Key</term>
        '''<description>Propriedade da Chave do web.config/app.config setado para a string de conexão</description>
        '''</item>
        '''<item>
        ''' <term>CommandText</term>
        '''<description>Propriedade da instrução SQL a ser rodada</description>
        '''</item>
        ''' </list>
        ''' <seealso cref="Command.Key" />
        ''' <seealso cref="Command.CommandText" />
        ''' <seealso cref="Command.CommandTimeout" />
        ''' <seealso cref="Command.CommandType" />
        ''' </summary>
        ''' <param name="strTableName">Nome da Tabela, opcional</param>
        '''<remarks>
        '''</remarks>
        '''<returns>DataTable</returns>
        ''' -----------------------------------------------------------------------------
        Public Function GetDataTable(Optional ByVal strTableName As String = "DataSet1") As DataTable Implements ICommand.GetDataTable

            Try
                ValidateCommand()
            Catch ex As Exception
                Throw New ArgumentNullException(ex.Message.ToString, "ReturnDataTable()")
            End Try

            Dim dtDados As New DataTable(strTableName)
            Dim parDados As DbParameter = Nothing
            Dim oCommand As DbCommand
            Dim parDadosRet As Parameter = Nothing
            Dim oConnectionFactory As MyDbConnection

            If strKeyName = "_defaultByCommand" Then
                oConnectionFactory = New MyDbConnection(oConnStringData)
            Else
                oConnectionFactory = New MyDbConnection(strKeyName)
            End If

            oCommand = oConnectionFactory.GetNewCommand

            Try
                'conformando a existencia ou nao de transacao
                ' oConnectionFactory.HasTransaction = HasTransact

                'setando o command text, time out, tipo e coletando os parametros.
                oCommand.CommandText = CommandText
                oCommand.CommandTimeout = CommandTimeout
                oCommand.CommandType = CommandType

                'colertando os parametros
                If oParameters.Count > 0 Then

                    For Each parItem As Parameter In oParameters
                        parDados = oConnectionFactory.GetNewParameter
                        parDados.DbType = parItem.DbType
                        parDados.Direction = parItem.Direction
                        parDados.ParameterName = parItem.ParameterName
                        parDados.Size = parItem.Size
                        parDados.SourceColumn = parItem.SourceColumn
                        parDados.Value = parItem.Value
                        oCommand.Parameters.Add(parDados)
                    Next
                End If

                If Prepared Then
                    oCommand.Prepare()
                End If

                dtDados = oConnectionFactory.GetDataTable(oCommand)

                If oParameters.Count > 0 Then
                    Parameters.Clear()

                    For Each parItem As DbParameter In oCommand.Parameters
                        parDadosRet = New Parameter
                        parDadosRet.DbType = parItem.DbType
                        parDadosRet.Direction = parItem.Direction
                        parDadosRet.ParameterName = parItem.ParameterName
                        parDadosRet.Size = parItem.Size
                        parDadosRet.SourceColumn = parItem.SourceColumn
                        parDadosRet.Value = parItem.Value
                        Parameters.Add(parDadosRet)
                    Next
                End If

                parDados = Nothing
                parDadosRet = Nothing

                If Not IsNothing(oCommand.Connection) Then
                    If oCommand.Connection.State = ConnectionState.Open Then
                        oCommand.Connection.Close()
                    End If
                    oCommand.Connection.Dispose()
                End If
                oCommand.Parameters.Clear()
                oCommand.Dispose()
                oCommand.Connection = Nothing
                oCommand = Nothing

                oConnectionFactory.Dispose()
                oConnectionFactory = Nothing

            Catch ex As Exception
                parDados = Nothing
                parDadosRet = Nothing

                'importante para fechar efetivamente a conexão com o oracle..
                If Not IsNothing(oCommand.Connection) Then
                    If oCommand.Connection.State = ConnectionState.Open Then
                        oCommand.Connection.Close()
                    End If
                    oCommand.Connection.Dispose()
                End If
                oCommand.Parameters.Clear()
                oCommand.Dispose()
                oCommand.Connection = Nothing
                oCommand = Nothing

                oConnectionFactory.Dispose()
                oConnectionFactory = Nothing

                dtDados = Nothing
                Throw New MyDbExeption(ex)
            End Try

            Return dtDados
        End Function

#End Region

#Region "DataSet"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Retorna um datatable.<BR/>
        ''' Para este Método funcionar correramente é preciso que alguns parametros estejam preenchidos corretamente.<BR/>
        ''' Estes paramtros são preenchidos alimentando as propriedades abaixo. <BR/>
        ''' Devem ser preenchidas TODAS:
        ''' <list type="bullet">
        '''<item>
        ''' <term>Key</term>
        '''<description>Propriedade da Chave do web.config/app.config setado para a string de conexão</description>
        '''</item>
        '''<item>
        ''' <term>CommandText</term>
        '''<description>Propriedade da instrução SQL a ser rodada</description>
        '''</item>
        ''' </list>
        ''' <seealso cref="Command.Key" />
        ''' <seealso cref="Command.CommandText" />
        ''' <seealso cref="Command.CommandTimeout" />
        ''' <seealso cref="Command.CommandType" />
        ''' </summary>
        '''<remarks>
        '''</remarks>
        ''' <returns>DataSet</returns>
        ''' -----------------------------------------------------------------------------
        Public Overridable Function GetDataSet(Optional ByVal strTableName As String = "DataSet1") As DataSet Implements ICommand.GetDataSet

            Return Me.GetDataTable(strTableName).DataSet

        End Function

#End Region

#Region "DataReader"

        ''' <summary>
        ''' Obtem um datareader, DESCONECTADO
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDataReader() As DbDataReader Implements ICommand.GetDataReader

            Try
                ValidateCommand()
            Catch ex As Exception
                Throw New ArgumentNullException(ex.Message.ToString, "ReturnDataTable()")
            End Try

            Dim parDados As DbParameter = Nothing
            Dim oCommand As DbCommand
            Dim parDadosRet As Parameter = Nothing
            Dim oDataReader As DbDataReader
            Dim oConnectionFactory As MyDbConnection

            If strKeyName = "_defaultByCommand" Then
                oConnectionFactory = New MyDbConnection(oConnStringData)
            Else
                oConnectionFactory = New MyDbConnection(strKeyName)
            End If

            oCommand = oConnectionFactory.GetNewCommand

            Try
                'conformando a existencia ou nao de transacao
                ' oConnectionFactory.HasTransaction = HasTransact

                'setando o command text, time out, tipo e coletando os parametros.
                oCommand.CommandText = CommandText
                oCommand.CommandTimeout = CommandTimeout
                oCommand.CommandType = CommandType

                'colertando os parametros
                If oParameters.Count > 0 Then

                    For Each parItem As Parameter In oParameters
                        parDados = oConnectionFactory.GetNewParameter
                        parDados.DbType = parItem.DbType
                        parDados.Direction = parItem.Direction
                        parDados.ParameterName = parItem.ParameterName
                        parDados.Size = parItem.Size
                        parDados.SourceColumn = parItem.SourceColumn
                        parDados.Value = parItem.Value
                        oCommand.Parameters.Add(parDados)
                    Next
                End If

                If Prepared Then
                    oCommand.Prepare()
                End If

                'deve ser devolvido desconectado
                oDataReader = oCommand.ExecuteReader(CommandBehavior.CloseConnection)

                If oParameters.Count > 0 Then
                    Parameters.Clear()

                    For Each parItem As DbParameter In oCommand.Parameters
                        parDadosRet = New Parameter
                        parDadosRet.DbType = parItem.DbType
                        parDadosRet.Direction = parItem.Direction
                        parDadosRet.ParameterName = parItem.ParameterName
                        parDadosRet.Size = parItem.Size
                        parDadosRet.SourceColumn = parItem.SourceColumn
                        parDadosRet.Value = parItem.Value
                        Parameters.Add(parDadosRet)
                    Next
                End If

                parDados = Nothing
                parDadosRet = Nothing

                If Not IsNothing(oCommand.Connection) Then
                    If oCommand.Connection.State = ConnectionState.Open Then
                        oCommand.Connection.Close()
                    End If
                    oCommand.Connection.Dispose()
                End If
                oCommand.Parameters.Clear()
                oCommand.Dispose()
                oCommand.Connection = Nothing
                oCommand = Nothing

                oConnectionFactory.Dispose()
                oConnectionFactory = Nothing
            Catch ex As Exception
                parDados = Nothing
                parDadosRet = Nothing

                'importante para fechar efetivamente a conexão com o oracle..
                If Not IsNothing(oCommand.Connection) Then
                    If oCommand.Connection.State = ConnectionState.Open Then
                        oCommand.Connection.Close()
                    End If
                    oCommand.Connection.Dispose()
                End If
                oCommand.Parameters.Clear()
                oCommand.Dispose()
                oCommand.Connection = Nothing
                oCommand = Nothing

                oConnectionFactory.Dispose()
                oConnectionFactory = Nothing

                oDataReader = Nothing
                Throw New MyDbExeption(ex)
            End Try

            Return oDataReader

        End Function

#End Region

#Region "Executes"

        ''' <summary>
        ''' Executa uma query
        ''' </summary>
        ''' <returns>Integer</returns>
        ''' <remarks></remarks>
        Public Function ExecuteNonQuery() As Integer Implements ICommand.ExecuteNonQuery
            Return Execute()
        End Function

        ''' <summary>
        ''' Executa uma query
        ''' </summary>
        ''' <returns>Integer</returns>
        ''' <remarks></remarks>
        Public Function Execute() As Integer Implements ICommand.Execute

            Try
                ValidateCommand()
            Catch ex As Exception
                Throw New ArgumentNullException(ex.Message.ToString, "Execute()")
            End Try

            Dim parDados As DbParameter = Nothing
            Dim oCommand As DbCommand
            Dim parDadosRet As Parameter = Nothing
            Dim intRetorno As Integer
            Dim oConnectionFactory As MyDbConnection

            If strKeyName = "_defaultByCommand" Then
                oConnectionFactory = New MyDbConnection(oConnStringData)
            Else
                oConnectionFactory = New MyDbConnection(strKeyName)
            End If

            oCommand = oConnectionFactory.GetNewCommand

            Try
                'conformando a existencia ou nao de transacao
                ' oConnectionFactory.HasTransaction = HasTransact

                'setando o command text, time out, tipo e coletando os parametros.
                oCommand.CommandText = CommandText
                oCommand.CommandTimeout = CommandTimeout
                oCommand.CommandType = CommandType

                'colertando os parametros
                If oParameters.Count > 0 Then

                    For Each parItem As Parameter In oParameters
                        parDados = oConnectionFactory.GetNewParameter
                        parDados.DbType = parItem.DbType
                        parDados.Direction = parItem.Direction
                        parDados.ParameterName = parItem.ParameterName
                        parDados.Size = parItem.Size
                        parDados.SourceColumn = parItem.SourceColumn
                        parDados.Value = parItem.Value
                        oCommand.Parameters.Add(parDados)
                    Next
                End If

                If Prepared Then
                    oCommand.Prepare()
                End If

                intRetorno = oCommand.ExecuteNonQuery()

                If oParameters.Count > 0 Then
                    Parameters.Clear()

                    For Each parItem As DbParameter In oCommand.Parameters
                        parDadosRet = New Parameter
                        parDadosRet.DbType = parItem.DbType
                        parDadosRet.Direction = parItem.Direction
                        parDadosRet.ParameterName = parItem.ParameterName
                        parDadosRet.Size = parItem.Size
                        parDadosRet.SourceColumn = parItem.SourceColumn
                        parDadosRet.Value = parItem.Value
                        Parameters.Add(parDadosRet)
                    Next
                End If

                parDados = Nothing
                parDadosRet = Nothing

                If Not IsNothing(oCommand.Connection) Then
                    If oCommand.Connection.State = ConnectionState.Open Then
                        oCommand.Connection.Close()
                    End If
                    oCommand.Connection.Dispose()
                End If
                oCommand.Parameters.Clear()
                oCommand.Dispose()
                oCommand.Connection = Nothing
                oCommand = Nothing

                oConnectionFactory.Dispose()
                oConnectionFactory = Nothing

            Catch ex As Exception
                parDados = Nothing
                parDadosRet = Nothing

                'importante para fechar efetivamente a conexão com o oracle..
                If Not IsNothing(oCommand.Connection) Then
                    If oCommand.Connection.State = ConnectionState.Open Then
                        oCommand.Connection.Close()
                    End If
                    oCommand.Connection.Dispose()
                End If
                oCommand.Parameters.Clear()
                oCommand.Dispose()
                oCommand.Connection = Nothing
                oCommand = Nothing

                oConnectionFactory.Dispose()
                oConnectionFactory = Nothing

                intRetorno = 0
                Throw New MyDbExeption(ex)
            End Try

            Return intRetorno
        End Function

        ''' <summary>
        ''' Executa uma query
        ''' </summary>
        ''' <returns>Integer</returns>
        ''' <remarks></remarks>
        Public Function ExecuteScalar() As Object Implements ICommand.ExecuteScalar

            Try
                ValidateCommand()
            Catch ex As Exception
                Throw New ArgumentNullException(ex.Message.ToString, "ReturnDataTable()")
            End Try

            Dim parDados As DbParameter = Nothing
            Dim oCommand As DbCommand
            Dim parDadosRet As Parameter = Nothing
            Dim oRetorno As Object
            Dim oConnectionFactory As MyDbConnection

            If strKeyName = "_defaultByCommand" Then
                oConnectionFactory = New MyDbConnection(oConnStringData)
            Else
                oConnectionFactory = New MyDbConnection(strKeyName)
            End If

            oCommand = oConnectionFactory.GetNewCommand

            Try
                'conformando a existencia ou nao de transacao
                ' oConnectionFactory.HasTransaction = HasTransact

                'setando o command text, time out, tipo e coletando os parametros.
                oCommand.CommandText = CommandText
                oCommand.CommandTimeout = CommandTimeout
                oCommand.CommandType = CommandType

                'colertando os parametros
                If oParameters.Count > 0 Then

                    For Each parItem As Parameter In oParameters
                        parDados = oConnectionFactory.GetNewParameter
                        parDados.DbType = parItem.DbType
                        parDados.Direction = parItem.Direction
                        parDados.ParameterName = parItem.ParameterName
                        parDados.Size = parItem.Size
                        parDados.SourceColumn = parItem.SourceColumn
                        parDados.Value = parItem.Value
                        oCommand.Parameters.Add(parDados)
                    Next
                End If

                If Prepared Then
                    oCommand.Prepare()
                End If

                oRetorno = oCommand.ExecuteScalar()

                If oParameters.Count > 0 Then
                    Parameters.Clear()

                    For Each parItem As DbParameter In oCommand.Parameters
                        parDadosRet = New Parameter
                        parDadosRet.DbType = parItem.DbType
                        parDadosRet.Direction = parItem.Direction
                        parDadosRet.ParameterName = parItem.ParameterName
                        parDadosRet.Size = parItem.Size
                        parDadosRet.SourceColumn = parItem.SourceColumn
                        parDadosRet.Value = parItem.Value
                        Parameters.Add(parDadosRet)
                    Next
                End If

                parDados = Nothing
                parDadosRet = Nothing

                If Not IsNothing(oCommand.Connection) Then
                    If oCommand.Connection.State = ConnectionState.Open Then
                        oCommand.Connection.Close()
                    End If
                    oCommand.Connection.Dispose()
                End If
                oCommand.Parameters.Clear()
                oCommand.Dispose()
                oCommand.Connection = Nothing
                oCommand = Nothing

                oConnectionFactory.Dispose()
                oConnectionFactory = Nothing

            Catch ex As Exception
                parDados = Nothing
                parDadosRet = Nothing

                'importante para fechar efetivamente a conexão com o oracle..
                If Not IsNothing(oCommand.Connection) Then
                    If oCommand.Connection.State = ConnectionState.Open Then
                        oCommand.Connection.Close()
                    End If
                    oCommand.Connection.Dispose()
                End If
                oCommand.Parameters.Clear()
                oCommand.Dispose()
                oCommand.Connection = Nothing
                oCommand = Nothing

                oConnectionFactory.Dispose()
                oConnectionFactory = Nothing

                oRetorno = 0
                Throw New MyDbExeption(ex)
            End Try

            Return oRetorno

        End Function

#End Region

#Region "Schema"

        ''' <summary>
        ''' Obtem um Schema do Banco
        ''' </summary>
        ''' <param name="strCollectionName">O nome do Schema a retornar</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSchema(ByVal strCollectionName As String) As DataTable Implements ICommand.GetSchema

            Dim oDataTable As DataTable = Nothing

            Try
                Using oConnectionFactory As MyDbConnection = New MyDbConnection(strKeyName)

                    oDataTable = New DataTable
                    oDataTable = oConnectionFactory.GetSchema(strCollectionName)

                End Using

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
        Public Function GetSchema() As DataTable Implements ICommand.GetSchema

            Dim oDataTable As DataTable = Nothing

            Try
                Using oConnectionFactory As MyDbConnection = New MyDbConnection(strKeyName)

                    oDataTable = New DataTable
                    oDataTable = oConnectionFactory.GetSchema()

                End Using

            Catch ex As Exception
                oDataTable = Nothing
                Throw New DataException(ex.Message.ToString)
            End Try

            Return oDataTable


        End Function

#End Region

#Region "Métodos Gerais"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Este comando ajusta o time out do command/conexão para o tempo de 30 segundos.
        ''' </summary>
        ''' <remarks>
        ''' O Valor inicial de um commandtimeout é de 360 segundos
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Sub ResetCommandTimeout() Implements ICommand.ResetCommandTimeout
            Me.CommandTimeout = 360
        End Sub

        Public Sub ReNewCommand() Implements ICommand.ReNewCommand
            stbCmdText = New StringBuilder
            Me.Parameters.Clear()
            Me.CommandText = ""
        End Sub
#End Region

#End Region

#Region "Propriedades"

        ''' <summary>
        ''' Torna ou não um command preparado
        ''' </summary>
        ''' <value>Boolean</value>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Property Prepared() As Boolean Implements ICommand.Prepared
            Get
                Return blnPrepared
            End Get
            Set(ByVal value As Boolean)
                blnPrepared = value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Key que ao ser preenchida permite buscar a string de conexão com o banco de dados.
        ''' </summary>
        ''' <value>String</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property Key() As String Implements ICommand.Key
            Get
                Return strKeyName
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Instrução de SQL para ser executada. <BR/>
        ''' Pode ser uma prodedure ou uma instrução normal de SQL
        ''' </summary>
        ''' <value>String</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property CommandText() As String Implements ICommand.CommandText
            Get
                strCmdText = stbCmdText.ToString
                Return strCmdText
            End Get
            Set(ByVal Value As String)
                stbCmdText = New StringBuilder
                stbCmdText.Append(String.Intern(Value))
                strCmdText = String.Intern(Value)
            End Set
        End Property

        ''' <summary>
        ''' Command Text, setável dentro de uma StringBuilder
        ''' </summary>
        ''' <value>StringBuilder</value>
        ''' <returns>StringBuilder</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property CommandTextObject() As StringBuilder Implements ICommand.CommandTextObject
            Get
                strCmdText = stbCmdText.ToString
                Return stbCmdText
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Tempo Limite de espera para execução de um comando. A unidade de tempo é em segundos.
        ''' </summary>
        ''' <value>Inteiro. Representa o tempo em segundos</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property CommandTimeout() As Integer Implements ICommand.CommandTimeout
            Get
                'propriedade não suportada para o Oracle
                Return intCommandTimeout
            End Get
            Set(ByVal Value As Integer)
                intCommandTimeout = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Enumerado especificando o tipo de commando
        ''' Para melhor informação veja em <seealso cref="System.Data.CommandType">System.Data.CommandType</seealso>
        ''' </summary>
        ''' <value>Tipo Inteiro, para o numerador System.Data.CommandType</value>
        ''' <remarks>
        ''' Renderiza um enumerador vindo de System.Data.CommandType. Qualquer alteração na origem afeta
        ''' esta classe, de forma transparente, sem ocasionar quebras de compatibilidade internas.
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property CommandType() As CommandType Implements ICommand.CommandType
            Get
                Return CType(intCommandType, System.Data.CommandType)
            End Get
            Set(ByVal Value As CommandType)
                intCommandType = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Coleção de objetos Parameter. Serve para guardar parametros para execução de uma query.
        ''' </summary>
        ''' <value>Parameter</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property Parameters() As List(Of Parameter) Implements ICommand.Parameters
            Get
                Return oParameters
            End Get
        End Property

#End Region

#Region " IDisposable Support "
        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    stbCmdText = Nothing
                    oParameters = Nothing
                End If

                intCommandType = 0
                intCommandTimeout = 0
                blnPrepared = Nothing
                strCmdText = Nothing
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

