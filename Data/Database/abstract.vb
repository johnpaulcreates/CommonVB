Imports System.Data.Common

Namespace Data
    Namespace Database
        Public MustInherit Class Abstract

            Protected MustOverride Function GetConnection() As DbConnection
            Protected MustOverride Function GetDataAdapter(ByVal Query As String, ByVal Connection As DbConnection) As DbDataAdapter

            Protected iDatabaseType As eDatabaseType
            Public ReadOnly Property Type As eDatabaseType
                Get
                    Return iDatabaseType
                End Get
            End Property

            Public Function ExecuteQuery(ByVal sSQL As String) As DataSet

                Dim oCN As DbConnection = GetConnection()
                Dim DA As DbDataAdapter = GetDataAdapter(sSQL, oCN)

                Dim DS As New DataSet
                DA.Fill(DS)
                DS.Dispose()
                oCN.Close()

                Return DS

            End Function

        End Class
    End Namespace

End Namespace