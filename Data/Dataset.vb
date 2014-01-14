Imports System.Data

Namespace Data
    Public Class DataSet
        Inherits System.Data.DataSet


        Private sFileName As String
        ''' <summary>
        ''' Filename of any serialised data
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property FileName() As String
            Get
                Return sFileName
            End Get
            Set(ByVal value As String)
                sFileName = value
            End Set
        End Property

        ''' <summary>
        ''' Load dataset from current Serialised XML File
        ''' </summary>
        ''' <remarks></remarks>
        Public Shadows Sub Load()
            Load(sFileName)
        End Sub

        ''' <summary>
        ''' Load Dataset from Serialised XML file
        ''' </summary>
        ''' <param name="Filename"></param>
        ''' <remarks></remarks>
        Public Shadows Sub Load(ByVal Filename As String)

            If System.IO.File.Exists(Filename) = False Then
                Dim aWobbler As New Exception("File Does Not Exist")
                Throw aWobbler
            End If

            Me.ReadXml(Filename)
            sFileName = Filename

        End Sub

        ''' <summary>
        ''' Serialise this Dataset to an XML file
        ''' </summary>
        ''' <param name="Filename"></param>
        ''' <remarks></remarks>
        Public Sub Save(Optional ByVal Filename As String = "")

            If Filename.Trim <> "" Then
                sFileName = Filename
            End If
            sFileName = sFileName.Trim

            If sFileName = "" Then
                Dim aWobbler As New Exception("Filename not set")
                Throw aWobbler
            End If

            Me.AcceptChanges()
            Me.WriteXml(Filename, System.Data.XmlWriteMode.WriteSchema)
            sFileName = Filename

        End Sub

        ''' <summary>
        ''' Add a new Table to this dataset
        ''' </summary>
        ''' <param name="TableName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function CreateTable(ByVal TableName As String) As DataTable

            'check to see if the table already exists
            If Me.Tables.Contains(TableName) Then
                Dim aWobbler As New Exception("Table '" & TableName & "' already exists")
                Throw aWobbler
            End If

            Dim DT As DataTable = New DataTable()
            DT.TableName = TableName
            Tables.Add(DT)

            Return DT

        End Function

    End Class
End Namespace

