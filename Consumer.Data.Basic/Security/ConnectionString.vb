Imports System.Configuration
Namespace Security
    Public Class ConnectionString

        Private strServer As String
        Private strUserName As String
        Private strPassword As String
        Private strDataBase As String
        Private strConection As String
        Private strProviderName As String

        ''' <summary>
        ''' Construtor padrão
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            strServer = ""
            strUserName = ""
            strPassword = ""
            strDataBase = ""
            strConection = ""
            strProviderName = ""
        End Sub

        ''' <summary>
        ''' Construtor para inicializar com valores os parâmetros do Banco.
        ''' </summary>
        ''' <param name="pStrServer">Nome do Servidor</param>
        ''' <param name="pStrUserName">Nome do Usuário</param>
        ''' <param name="pStrPassword">Senha. Por padrão deve vir Criptografada</param>
        ''' <param name="pStrDataBase">Nome do Banco</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal pStrServer As String, _
                       ByVal pStrUserName As String, _
                       ByVal pStrPassword As String, _
                       ByVal pStrDataBase As String)

            Me.New()
            strServer = pStrServer
            strUserName = pStrUserName
            strPassword = pStrPassword
            strDataBase = pStrDataBase

        End Sub

        ''' <summary>
        ''' Construtor para inicializar com valores os parâmetros do Banco. Com provider selecionado
        ''' </summary>
        ''' <param name="pStrServer">Nome do Servidor</param>
        ''' <param name="pStrUserName">Nome do Usuário</param>
        ''' <param name="pStrPassword">Senha. Por padrão deve vir Criptografada</param>
        ''' <param name="pStrDataBase">Nome do Banco</param>
        ''' <param name="pStrProvider">Provider Name</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal pStrServer As String, _
                       ByVal pStrUserName As String, _
                       ByVal pStrPassword As String, _
                       ByVal pStrDataBase As String, _
                       ByVal pStrProvider As String)

            Me.New()
            strServer = pStrServer
            strUserName = pStrUserName
            strPassword = pStrPassword
            strDataBase = pStrDataBase
            strProviderName = pStrProvider
        End Sub

        ''' <summary>
        ''' Construtor, que aceita o key do config como parâmetro, buscando já a string de conexão
        ''' </summary>
        ''' <param name="strKey">Key do Config</param>
        ''' <remarks>Este construtor já considera que a senha da string de conexão virá criptografada. Caso não venha, haverá erro.</remarks>
        Public Sub New(ByVal strKey As String)
            Me.New()
            strConection = GetConectionString(strKey)
            GetData(strConection, True)
        End Sub

        ''' <summary>
        ''' Construtor que recebe a string de conexão criptografada ou não.
        ''' </summary>
        ''' <param name="strConectionString">String de Conexão</param>
        ''' <param name="blnIsPasswordCrypt">Se é criptografada ou não</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal strConectionString As String, ByVal blnIsPasswordCrypt As Boolean)
            Me.New()
            GetData(strConectionString, blnIsPasswordCrypt)
        End Sub

        ''' <summary>
        ''' Obtem a string de Conexão de acordo com o Key
        ''' </summary>
        ''' <param name="strKey">Key, que deve estar na seção de appSettings ou connectionStrings</param>
        ''' <returns>String</returns>
        ''' <remarks></remarks>
        Public Function GetConectionString(ByVal strKey As String) As String
            Dim strConn As String

            If Not IsNothing(ConfigurationManager.ConnectionStrings(strKey)) Then
                strConn = ConfigurationManager.ConnectionStrings(strKey).ConnectionString.ToString()
                If strProviderName = String.Empty Then
                    strProviderName = ConfigurationManager.ConnectionStrings(strKey).ProviderName.ToString()
                End If
            Else
                'este uso será removido em breve
                If Not IsNothing(ConfigurationManager.AppSettings(strKey)) Then
                    strConn = ConfigurationManager.AppSettings(strKey).ToString()
                Else
                    Throw New ArgumentNullException("Não existem conexões definidas na Chave de <appSettings> ou <connectionStrings>!")
                End If
            End If
            Return strConn

        End Function

        ''' <summary>
        ''' Valida as propriedades se foram preenchidas corretamente - OLE DB
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub ValidatePropertiesOleDB()

            If strServer = "" Then
                Throw New ArgumentNullException("Não existem valores para conexões definidos.Argumento Faltante: Servidor")
            End If
            If strProviderName = "" Then
                Throw New ArgumentNullException("Não existem valores para conexões definidos.Argumento Faltante: Provedor de Dados")
            End If
            If strDataBase = "" Then
                Throw New ArgumentNullException("Não existem valores para conexões definidos.Argumento Faltante: Banco de Dados")
            End If

        End Sub
        ''' <summary>
        ''' Valida as propriedades se foram preenchidas corretamente - SQL
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub ValidatePropertiesSQL()

            If strServer = "" Then
                Throw New ArgumentNullException("Não existem valores para conexões definidos.Argumento Faltante: Servidor")
            End If
            If strUserName = "" Then
                Throw New ArgumentNullException("Não existem valores para conexões definidos.Argumento Faltante: Usuário")
            End If
            If strPassword = "" Then
                Throw New ArgumentNullException("Não existem valores para conexões definidos.Argumento Faltante: Senha")
            End If
            If strProviderName = "" Then
                Throw New ArgumentNullException("Não existem valores para conexões definidos.Argumento Faltante: Provedor de Dados")
            End If
            If strDataBase = "" Then
                Throw New ArgumentNullException("Não existem valores para conexões definidos.Argumento Faltante: Banco de Dados")
            End If

        End Sub

        ''' <summary>
        ''' Obtem a string de Conexão com a senha criptografada
        ''' Método importante para atualizar as chaves do config
        ''' </summary>
        ''' <returns>String</returns>
        ''' <remarks></remarks>
        Public Function GetConectionString() As String
            Dim strConn As String = String.Empty

            If strProviderName.Trim() = "" Or strProviderName = "System.Data.OleDb" Then
                strConn = "Provider=" + strServer + ";" + _
                          "uid=" + strUserName + ";" + _
                          "pwd=" + strPassword + ";" + _
                          "Data Source=" + strDataBase + ";"
            Else
                strConn = "server=" + strServer + ";" + _
                           "uid=" + strUserName + ";" + _
                           "pwd=" + Encript.EncryptString(strPassword) + ";" + _
                           "database=" + strDataBase + ";"
            End If

            Return strConn

        End Function

        ''' <summary>
        ''' Obtem a string de Conexão com a senha de-criptografada
        ''' Este método é o que se usa para obter a string de conexão do config com a senha criptografada, e 
        ''' ser usado para conectar a base de dados
        ''' </summary>
        ''' <returns>String</returns>
        ''' <remarks></remarks>
        Public Function GetConectionStringForConnect() As String

            Dim strConn As String = String.Empty

            If strProviderName.Trim() = "" Or strProviderName = "System.Data.OleDb" Then
                strConn = "Provider=" + strProviderName + ";" + _
                          "uid=" + strUserName + ";" + _
                          "pwd=" + strPassword + ";" + _
                          "Data Source=" + strDataBase + ";"
            Else
                strConn = "server=" + strServer + ";" + _
                          "uid=" + strUserName + ";" + _
                          "pwd=" + strPassword + ";" + _
                          "database=" + strDataBase + ";"
            End If

            Return strConn

        End Function

        ''' <summary>
        ''' Método que pega a string de coenxão e preenche as propriedades desta classe presente
        ''' </summary>
        ''' <param name="strConection">String de Conexão</param>
        ''' <param name="blnIsPasswordCrypt">Se a senha está ou não criptografada</param>
        ''' <remarks></remarks>
        Private Sub GetData(ByVal strConection As String, ByVal blnIsPasswordCrypt As Boolean)

            If strProviderName.Trim() = "" Or strProviderName = "System.Data.OleDb" Then
                'sem provider name preenchido. é caso de provavelmente ser OLEDB
                strProviderName = "System.Data.OleDb"
                GetDataOleDB(strConection)
                ValidatePropertiesOleDB()
            Else
                GetDataSQL(strConection, blnIsPasswordCrypt)
                ValidatePropertiesSQL()
            End If

        End Sub

        ''' <summary>
        ''' Método que pega a string de coenxão e 
        ''' preenche as propriedades desta classe presente, 
        ''' válido para Ole DB Providers
        ''' </summary>
        ''' <param name="strConection">String de Conexão</param>
        ''' <remarks></remarks>
        Private Sub GetDataOleDB(ByVal strConection As String)

            Dim strApoio() As String
            Dim strApoio1() As String

            'Provider=Microsoft.Jet.OLEDB.4.0;
            'Data Source=App_Data\zzzzz.mdb

            strApoio = strConection.Split(";"c)
            'terei algumas ocorrências, e não sei a ordem.
            For intCont As Integer = 0 To strApoio.Length - 1
                strApoio1 = strApoio(intCont).Split("="c)
                If strApoio1(0).ToLower() = "provider" Then
                    strServer = strApoio1(1)
                End If
                If strApoio1(0).ToLower() = "data source" Then
                    strDataBase = strApoio1(1)
                End If
                If strApoio1(0).ToLower() = "uid" Or strApoio1(0).ToLower() = "user id" Then
                    strUserName = strApoio1(1)
                End If
                If strApoio1(0).ToLower() = "pwd" Or strApoio1(0).ToLower() = "password" Then
                    strPassword = strApoio1(1)
                End If
            Next

            If strUserName = "" Then strUserName = "Admin"

        End Sub

        ''' <summary>
        ''' Método que pega a string de coenxão e 
        ''' preenche as propriedades desta classe presente, 
        ''' válido para SQLs Providers
        ''' </summary>
        ''' <param name="strConection">String de Conexão</param>
        ''' <param name="blnIsPasswordCrypt">Se a senha está ou não criptografada</param>
        ''' <remarks></remarks>
        Private Sub GetDataSQL(ByVal strConection As String, ByVal blnIsPasswordCrypt As Boolean)

            Dim strApoio() As String
            Dim strApoio1() As String

            strApoio = strConection.Split(";"c)
            'connectionString="Data Source=.\SQLExpress;
            'Initial Catalog=torchwood_dev;
            'Persist Security Info=True;
            'User ID=sa;
            'Password=DMHHSPNOUJ" 

            'terei quatro ocorrências, e não sei a ordem.
            For intCont As Integer = 0 To strApoio.Length - 1
                strApoio1 = strApoio(intCont).Split("="c)
                If strApoio1(0).ToLower() = "server" Or strApoio1(0).ToLower() = "data source" Then
                    strServer = strApoio1(1)
                End If
                If strApoio1(0).ToLower() = "uid" Or strApoio1(0).ToLower() = "user id" Then
                    strUserName = strApoio1(1)
                End If
                If strApoio1(0).ToLower() = "pwd" Or strApoio1(0).ToLower() = "password" Then
                    If blnIsPasswordCrypt Then
                        strPassword = Encript.DecryptQueryString(strApoio1(1))
                        If strPassword = "" And strApoio1(1) <> "" Then
                            strPassword = Encript.DecryptString(strApoio1(1))
                        End If
                    Else
                        strPassword = strApoio1(1)
                    End If
                End If
                If strApoio1(0).ToLower() = "database" Or strApoio1(0).ToLower() = "initial catalog" Then
                    strDataBase = strApoio1(1)
                End If
            Next
        End Sub

        ''' <summary>
        ''' Grava a string de conexão na chave do config
        ''' Como o padrão de desenvolvimento é o framework 2.0 este método exclusivamente grava na seção do
        ''' config reservada as strings de Conexão
        ''' </summary>
        ''' <param name="strKeyName">Nome da Key para ser gravado</param>
        ''' <remarks></remarks>
        Public Sub WriteToKey(ByVal strKeyName As String)
            Dim strConn As String
            If strProviderName = "System.Data.Oledb" Then
                Me.ValidatePropertiesOleDB()
            Else
                Me.ValidatePropertiesSQL()
            End If

            strConn = Me.GetConectionString
            ConfigurationManager.ConnectionStrings(strKeyName).ProviderName = strProviderName
            ConfigurationManager.ConnectionStrings(strKeyName).ConnectionString = strConn

        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ServerName() As String
            Get
                Return strServer
            End Get
            Set(ByVal Value As String)
                strServer = String.Intern(Value)
            End Set
        End Property

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property UserName() As String
            Get
                Return strUserName
            End Get
            Set(ByVal Value As String)
                strUserName = String.Intern(Value)
            End Set
        End Property

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Password() As String
            Get
                Return strPassword
            End Get
            Set(ByVal Value As String)
                strPassword = String.Intern(Value)
            End Set
        End Property

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DataBase() As String
            Get
                Return strDataBase
            End Get
            Set(ByVal Value As String)
                strDataBase = String.Intern(Value)
            End Set
        End Property

        ''' <summary>
        ''' Nome do Provider que será feita a conexão com o banco de dados.
        ''' </summary>
        ''' <value>String</value>
        ''' <returns>String</returns>
        ''' <remarks></remarks>
        Public Property ProviderName() As String
            Get
                Return strProviderName
            End Get
            Set(ByVal Value As String)
                strProviderName = String.Intern(Value)
            End Set
        End Property

    End Class
End Namespace