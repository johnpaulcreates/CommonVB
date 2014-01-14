Imports System.Windows.Forms

Namespace WinForms
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks>
    ''' example: DirectCast(listView1, ListViewGroupSorter).SortGroups(True) 
    '''</remarks>
    Public Class ListViewGroupSorter

        Friend _listview As ListView

        Public Shared Operator =(ByVal listview As ListView, ByVal sorter As ListViewGroupSorter) As Boolean
            Return listview Is sorter._listview
        End Operator
        Public Shared Operator <>(ByVal listview As ListView, ByVal sorter As ListViewGroupSorter) As Boolean
            Return Not listview Is sorter._listview
        End Operator

        Public Shared Widening Operator CType(ByVal sorter As ListViewGroupSorter) As ListView
            Return sorter._listview
        End Operator
        Public Shared Widening Operator CType(ByVal listview As ListView) As ListViewGroupSorter
            Return New ListViewGroupSorter(listview)
        End Operator

        Friend Sub New(ByVal listview As ListView)
            _listview = listview
        End Sub

        Public Sub SortGroups(ByVal SortOrder As SortOrder)
            _listview.BeginUpdate()
            Dim lvgs As New List(Of ListViewGroup)()
            For Each lvg As ListViewGroup In _listview.Groups
                lvgs.Add(lvg)
            Next
            _listview.Groups.Clear()
            lvgs.Sort(New ListViewGroupSortComparer(SortOrder.Ascending))
            _listview.Groups.AddRange(lvgs.ToArray())
            _listview.EndUpdate()
        End Sub

#Region "overridden methods"

        Public Overrides Function Equals(ByVal obj As Object) As Boolean
            Return _listview.Equals(obj)
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return _listview.GetHashCode()
        End Function

        Public Overrides Function ToString() As String
            Return _listview.ToString()
        End Function

#End Region
    End Class

End Namespace
