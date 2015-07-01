#Region "Imports"
Imports System.Data
Imports System.ComponentModel
Imports System.Globalization
Imports System.Data.Common
#End Region

Namespace Data

    ''' <summary>
    ''' Classe que renderiza um parametro para o command executar as procedures, ou queries parametrizadas.
    ''' </summary>
    ''' <remarks>
    ''' <seealso cref="Command" />
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    <Serializable(), Description("Classe que Renderiza um parametro para a classe Command")> _
    Public Class Parameter
        Implements IDisposable
        Private disposedValue As Boolean = False        ' To detect redundant calls

#Region "Destruidores"

        Public Sub Dispose() Implements System.IDisposable.Dispose
            Dispose(True)
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Função dispose, que destroi o objeto.
        ''' </summary>
        ''' <param name="Disposable">Se true, o objeto é destruido.</param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Protected Sub Dispose(ByVal Disposable As Boolean)
            If Disposable Then
                MyBase.Finalize()
            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Destrutior do objeto. Chamada quando a instancia do mesmo é setada para True.
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Protected Overrides Sub Finalize()

            intDbType = Nothing
            intDirection = Nothing
            strName = Nothing
            objValue = Nothing
            intSize = Nothing
            strSourceColumn = Nothing
            blnIsNullable = Nothing
            blnIsKey = Nothing
            MyBase.Finalize()
        End Sub

#End Region

#Region "Variaveis"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Variável auxiliar do projeto.
        ''' Campo que Renderiza a propridedade dbType
        ''' <seealso cref="System.Data.DbType" />
        ''' <seealso cref="Parameter.DbType" />
        ''' </summary>
        ''' <remarks>
        ''' Tipo de Dados declarado como <b><i>"Protected"</i></b>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private intDbType As DbType

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Variável auxiliar do projeto.
        ''' Campo que Renderiza a propridedade Direction
        ''' <seealso cref="Parameter.Direction" />
        ''' </summary>
        ''' <remarks>
        ''' Tipo de Dados declarado como <b><i>"Protected"</i></b>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private intDirection As ParameterDirection

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Variável auxiliar do projeto.
        ''' Campo que Renderiza a propridedade Name
        ''' <seealso cref="Parameter.Name" />
        ''' </summary>
        ''' <remarks>
        ''' Tipo de Dados declarado como <b><i>"Protected"</i></b>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private strName As String

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Variável auxiliar do projeto.
        ''' Campo que Renderiza a propridedade Value
        ''' <seealso cref="Parameter.Value" />
        ''' </summary>
        ''' <remarks>
        ''' Tipo de Dados declarado como <b><i>"Protected"</i></b>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private objValue As Object

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Variável auxiliar do projeto.
        ''' Campo que Renderiza a propridedade CommandTimeout
        ''' </summary>
        ''' <remarks>
        ''' Tipo de Dados declarado como <b><i>"Protected"</i></b>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private intSize As Integer

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Variável auxiliar do projeto.
        ''' Campo que Renderiza a propridedade Size
        ''' <seealso cref="Parameter.Size" />
        ''' </summary>
        ''' <remarks>
        ''' Tipo de Dados declarado como <b><i>"Protected"</i></b>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private strSourceColumn As String

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Variável auxiliar do projeto.
        ''' Campo que Renderiza a propridedade IsNullable
        ''' <seealso cref="Parameter.IsNullable" />
        ''' </summary>
        ''' <remarks>
        ''' Tipo de Dados declarado como <b><i>"Protected"</i></b>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private blnIsNullable As Boolean = False

        ''' <summary>
        ''' Variável que verifica se o campo é ou não chave. New in 4.0
        ''' </summary>
        ''' <remarks></remarks>
        Private blnIsKey As Boolean = False

        ''' <summary>
        ''' Variável que indica que o campo é um contador - Propriedade específica para o class builder - New in 1.5
        ''' </summary>
        ''' <remarks></remarks>
        Private blnIsCounter As Boolean = False
#End Region

#Region "Constructors"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Construtor da Classe de Parâmetro. <BR/>
        ''' Cria um novo parametro para ser usado.
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            intDbType = DbType.String
            intDirection = ParameterDirection.Input
            objValue = New Object
            intSize = 0
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Construtor da Classe de Parâmetro. <BR/>
        ''' Cria um novo parametro para ser usado.
        ''' </summary>
        ''' <param name="Name">nome do Parametro. Tipo String</param>
        ''' <param name="Value">Valor que o parametro recebe. Tipo Objeto.</param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Name As String, _
                       ByVal Value As Object)

            intDbType = DbType.String
            intDirection = ParameterDirection.Input
            objValue = Value
            strName = Name
            intSize = 0
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Construtor da Classe de Parâmetro. <BR/>
        ''' Cria um novo parametro para ser usado.
        ''' </summary>
        ''' <param name="Name">nome do Parametro. Tipo String</param>
        ''' <param name="Value">Valor que o parametro recebe. Tipo Objeto.</param>
        ''' <param name="Type">Tipo do Parametro. Tipo de dado, proveniente de enumerador</param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Name As String, _
                       ByVal Type As DbType, _
                       ByVal Value As Object)

            intDbType = Type
            intDirection = ParameterDirection.Input
            objValue = Value
            strName = Name
            intSize = 0
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Construtor da Classe de Parâmetro. <BR/>
        ''' Cria um novo parametro para ser usado.
        ''' </summary>
        ''' <param name="Name">nome do Parametro. Tipo String</param>
        ''' <param name="Value">Valor que o parametro recebe. Tipo Objeto.</param>
        ''' <param name="Type">Tipo do Parametro. Tipo de dado, proveniente de enumerador</param>
        ''' <param name="Size">Tamanho do Parametro. Para Strings, o tamanho do char/varchar</param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Name As String, _
                       ByVal Type As DbType, _
                       ByVal Size As Integer, _
                       ByVal Value As Object)
            intDbType = Type
            intDirection = ParameterDirection.Input
            objValue = Value
            strName = Name
            intSize = 0
        End Sub

#End Region

#Region "Propriedades"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Propriedade Que contem o valor do tipo do parâmetro se é permitido nulo ou não.
        ''' </summary>
        ''' <value>Boolean</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property IsNullable() As Boolean
            Get
                Return blnIsNullable
            End Get
            Set(ByVal Value As Boolean)
                blnIsNullable = Value
            End Set
        End Property

        ''' <summary>
        ''' Propriedade Verificadora se o campo é chave primaria ou não
        ''' </summary>
        ''' <value>Boolean</value>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Property IsKey() As Boolean
            Get
                Return blnIsKey
            End Get
            Set(ByVal Value As Boolean)
                blnIsKey = Value
            End Set
        End Property

        ''' <summary>
        ''' Propriedade que seta se o campo é ou não contador
        ''' </summary>
        ''' <value>Boolean</value>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Property IsCounter() As Boolean
            Get
                Return blnIsCounter
            End Get
            Set(ByVal value As Boolean)
                blnIsCounter = value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Propriedade Que contem o nome da coluna do parametro
        ''' </summary>
        ''' <value>String</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property SourceColumn() As String
            Get
                Return strSourceColumn
            End Get
            Set(ByVal Value As String)
                strSourceColumn = String.Intern(Value)
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Tipo do Parametro, valor proveniente de um enumerador do banco de dados
        ''' </summary>
        ''' <value>Inteiro (dbType)</value>
        ''' <seealso cref="system.data.dbtype" />
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property DbType() As DbType
            Get
                Return intDbType
            End Get
            Set(ByVal Value As DbType)
                intDbType = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Parâmetro que indica a direção, pode ser parâmetro de entrada, de saída ou os dois.
        ''' </summary>
        ''' <value>ParameterDirection - Tipo Inteiro.</value>
        ''' <seealso cref="system.data.ParameterDirection" />
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Direction() As ParameterDirection
            Get
                Return intDirection
            End Get
            Set(ByVal Value As ParameterDirection)
                intDirection = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Nome do parâmetro.
        ''' </summary>
        ''' <value>String</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ParameterName() As String
            Get
                Return strName
            End Get
            Set(ByVal Value As String)
                strName = String.Intern(Value)
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Nome do parâmetro.
        ''' </summary>
        ''' <value>String</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Name() As String
            Get
                Return strName
            End Get
            Set(ByVal Value As String)
                strName = String.Intern(Value)
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Valor do parâmetro
        ''' </summary>
        ''' <value>Object, mas pode ser qualquer tipo de dado primitivo.</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Value() As Object
            Get
                Return objValue
            End Get
            Set(ByVal Value As Object)
                objValue = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Tamanho da String.<br/>
        ''' Válido somente para tipos que conhenham strings como:
        ''' <list type="bullet">
        ''' <item>
        ''' <description>Char</description>        
        ''' </item>
        ''' <item>
        ''' <description>VarChar</description>        
        ''' </item>
        ''' <item>
        ''' <description>NVarChar</description>
        ''' </item>
        ''' </list>
        ''' </summary>
        ''' <value>Tamkanho da String. Tipo Inteiro.</value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Size() As Integer
            Get
                Return intSize
            End Get
            Set(ByVal Value As Integer)
                intSize = Value
            End Set
        End Property

#End Region

#Region "Outras Funcoes"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Retorna o Parametro convertido para String. Na verdade retorna o nome do parametro
        ''' </summary>
        ''' <returns>String</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[humberto]	07/04/2006	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Overrides Function ToString() As String
            Return Me.Name
        End Function

#End Region

    End Class

End Namespace