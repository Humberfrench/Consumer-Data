Imports System.Data
Imports System.Data.Common
Imports System.Text
Imports Consumer.Data.Basic.Data

Public Interface ICommand
    Inherits IDisposable

    Function CreateParameter() As Parameter
    Sub CreateParameter(ByVal Name As String, ByVal Type As DbType, ByVal Size As Integer, ByVal Value As Object)
    Function GetDataTable(Optional ByVal strTableName As String = "DataSet1") As DataTable
    Function GetDataSet(Optional ByVal strTableName As String = "DataSet1") As DataSet
    Function GetDataReader() As DbDataReader
    Function ExecuteNonQuery() As Integer
    Function Execute() As Integer
    Function ExecuteScalar() As Object
    Function GetSchema(ByVal strCollectionName As String) As DataTable
    Function GetSchema() As DataTable
    Sub ResetCommandTimeout()
    Sub ReNewCommand()
    Property Prepared() As Boolean
    ReadOnly Property Key() As String
    Property CommandText() As String
    ReadOnly Property CommandTextObject() As StringBuilder
    Property CommandTimeout() As Integer
    Property CommandType() As CommandType
    ReadOnly Property Parameters() As List(Of Parameter)

End Interface

Public Interface ITransactCommand
    Inherits ICommand

    Sub BeginTransaction()
    Sub CommitTransaction()
    Sub RollbackTransaction()

End Interface
