Namespace Data
    Namespace Database
        ''' <summary>
        ''' Database class where data is stored in an XML file
        ''' </summary>
        ''' <remarks></remarks>
        Public Class XML
            Inherits Abstract

            Private oDS As DataSet

            Private sLocation As String

            ''' <summary>
            ''' Should not used for this database type
            ''' </summary>
            ''' <param name="Query"></param>
            ''' <param name="Connection"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Protected Overrides Function GetDataAdapter(ByVal Query As String, ByVal Connection As System.Data.Common.DbConnection) As System.Data.Common.DbDataAdapter
                Debug.Assert(False, "Method not applicable for database type")
                Return Nothing
            End Function

            ''' <summary>
            ''' Should not be used for this database type
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Protected Overrides Function GetConnection() As System.Data.Common.DbConnection
                Debug.Assert(False, "Method not applicable for database type")
                Return Nothing
            End Function

            Public Sub New(ByVal DatabaseLocation As String)

                Debug.Assert(False, "not completed yet")

                If System.IO.File.Exists(DatabaseLocation) = False Then
                    Dim aWobbler As New Exception("File Not Found: " & DatabaseLocation)
                    Throw aWobbler
                End If

                sLocation = DatabaseLocation

                oDS = New DataSet
                oDS.Load(sLocation)

                iDatabaseType = eDatabaseType.XML


            End Sub


            'i dont really want to expose the dataset, so wrap a facade around the useful stuff:
            Public Function CreateTable(ByVal TableName As String) As DataTable
                Return oDS.CreateTable(TableName)
            End Function
            Public ReadOnly Property Tables As DataTableCollection
                Get
                    Return oDS.Tables
                End Get
            End Property
            Public Sub Save(Optional ByVal Filename As String = "")
                oDS.Save(Filename)
            End Sub

        End Class

    End Namespace
End Namespace
