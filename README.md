# Consumer-Data
Provedor de Dados, para qualquer tipo

User Example, via VB

        Public Function Add(oModel As Model.Telefone) As Boolean
            Dim cmdDados As Command = Nothing
            Dim blnReturn As Boolean = False

            Try
                cmdDados = New Command("CharlieKey")
                cmdDados.CommandType = CommandType.StoredProcedure
                cmdDados.CommandText = "pr_add_telefone"
                cmdDados.Parameters.Add(New Parameter("@id_cliente", DbType.String, 2, oModel.Cliente))
                cmdDados.Parameters.Add(New Parameter("@ds_tipo_telefone", DbType.String, 10, oModel.TipoTelefone))
                cmdDados.Parameters.Add(New Parameter("@nr_ddd", DbType.String, 50, oModel.DDD))
                cmdDados.Parameters.Add(New Parameter("@nr_Telefone", DbType.String, 50, oModel.Numero))
                cmdDados.Parameters.Add(New Parameter("@nr_ramal", DbType.String, 50, oModel.Complemento))

                cmdDados.ExecuteNonQuery()
                blnReturn = True

            Catch ex As Exception
                blnReturn = False
                Throw ex
            Finally
                cmdDados.Dispose()
                cmdDados = Nothing

            End Try
            Return blnReturn

        End Function
 OU
 
         Public Function ObterTelefones(intCliente As Integer) As List(Of Model.Telefone)

            Dim cmdDados As Command = Nothing
            Dim dtDados As DataTable = Nothing
            Dim oReturn As Model.Telefone = Nothing
            Dim lstReturn As List(Of Model.Telefone) = Nothing

            Try
                cmdDados = New Command("CharlieKey")
                cmdDados.CommandType = CommandType.StoredProcedure
                cmdDados.CommandText = "pr_list_Telefone"
                cmdDados.Parameters.Add(New Parameter("@id_cliente", DbType.Int32, intCliente))

                dtDados = cmdDados.GetDataTable
                oReturn = New Model.Telefone
                lstReturn = New List(Of Model.Telefone)

                If dtDados.Rows.Count > 0 Then
                    For Each oRow In dtDados.Rows
                        oReturn.Codigo = oRow("id_Telefone").ToString()
                        oReturn.Cliente = oRow("id_cliente").ToString()
                        oReturn.TipoTelefone = oRow("ds_tipo_Telefone").ToString()
                        oReturn.DDD = oRow("nr_ddd").ToString()
                        oReturn.Numero = oRow("nr_Telefone").ToString()
                        oReturn.Complemento = oRow("nr_ramal").ToString()
                        lstReturn.Add(oReturn)
                    Next
                End If
            Catch ex As Exception
                lstReturn = Nothing
                Throw ex
            Finally
                oReturn = Nothing
                cmdDados.Dispose()
                cmdDados = Nothing
                dtDados = Nothing

            End Try
            Return lstReturn

        End Function
