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
