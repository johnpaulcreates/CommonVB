Namespace Data
    Public Class DataTable
        Inherits System.Data.DataTable

        Public Enum eDataType
            [Integer]
            [Boolean]
            [String]
        End Enum

        ''' <summary>
        ''' Add a new DataColumn to this DataTable
        ''' </summary>
        ''' <param name="ColumnName"></param>
        ''' <param name="DataType"></param>
        ''' <param name="AutoNumber"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function CreateColumn(ByVal ColumnName As String, Optional ByVal DataType As eDataType = eDataType.String, Optional ByVal AutoNumber As Boolean = False) As DataColumn

            Dim iType As System.Type = GetSystemDataType(DataType)
            Dim Col As New DataColumn(ColumnName, iType)

            Columns.Add(Col)
            If AutoNumber Then
                Col.AutoIncrement = True
            End If

            Return Col

        End Function

        Private Function GetSystemDataType(ByVal DataType As eDataType) As System.Type

            Select Case DataType
                Case eDataType.Integer
                    Return System.Type.GetType("System.Int32")
                Case eDataType.String
                    Return System.Type.GetType("System.String")
                Case eDataType.Boolean
                    Return System.Type.GetType("System.Boolean")
                Case Else
                    Dim aWobbler As New Exception("Unexpected Data Type: " & DataType.ToString)
                    Throw aWobbler
            End Select

        End Function

    End Class
End Namespace

