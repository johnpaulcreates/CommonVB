Imports System.Windows.Forms

Namespace WinForms
    Public Class ListViewGroupSortComparer
        Implements IComparer(Of ListViewGroup)

        Private iSortOrder As SortOrder = SortOrder.Ascending
        Public Sub New(ByVal SortOrder As SortOrder)
            iSortOrder = SortOrder
        End Sub

        Public Function Compare(ByVal x As System.Windows.Forms.ListViewGroup, ByVal y As System.Windows.Forms.ListViewGroup) As Integer Implements System.Collections.Generic.IComparer(Of System.Windows.Forms.ListViewGroup).Compare

            If iSortOrder = SortOrder.Ascending Then
                Return String.Compare(DirectCast(x, ListViewGroup).Header, DirectCast(y, ListViewGroup).Header)
            Else
                Return String.Compare(DirectCast(y, ListViewGroup).Header, DirectCast(x, ListViewGroup).Header)
            End If
        End Function

    End Class
End Namespace

