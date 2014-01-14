Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Namespace WinForms

    ''' <summary>
    ''' Make using listviews easier and adds features
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ListViewHelper


#Region "Collapsable Groups"
        Private Const LVM_FIRST As Integer = &H1000
        Private Const LVM_SETGROUPINFO = LVM_FIRST + 147

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure LVGROUP
            Public cbSize As Integer
            Public mask As Integer
            <MarshalAs(UnmanagedType.LPTStr)> Public pszHeader As String
            Public cchHeader As Integer
            <MarshalAs(UnmanagedType.LPTStr)> Public pszFooter As String
            Public cchFooter As Integer
            Public iGroupId As Integer
            Public stateMask As Integer
            Public state As Integer
            Public uAlign As Integer
        End Structure

        Public Enum GroupState
            COLLAPSIBLE = 8
            COLLAPSED = 1
            EXPANDED = 0
        End Enum

        <DllImport("user32.dll")> _
        Private Shared Function SendMessage(ByVal window As IntPtr, ByVal message As Integer, ByVal wParam As Integer, ByVal lParam As IntPtr) As Integer
        End Function


        ''' <summary>
        ''' Set the collapse state for a single listviewgroup
        ''' </summary>
        ''' <param name="lvw"></param>
        ''' <param name="iGroup"></param>
        ''' <param name="State"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetGroupCollapse(ByVal lvw As ListView, ByVal iGroup As Integer, ByVal State As GroupState)

            Dim group As New LVGROUP()
            group.cbSize = Marshal.SizeOf(group)
            group.state = State ' LVGS_COLLAPSIBLE 
            group.mask = 4 ' LVGF_STATE 
            group.iGroupId = iGroup
            Dim ip As IntPtr = IntPtr.Zero

            Try
                ip = Marshal.AllocHGlobal(group.cbSize)
                Marshal.StructureToPtr(group, ip, True)
                SendMessage(lvw.Handle, LVM_SETGROUPINFO, iGroup, ip)

            Catch ex As Exception
                System.Diagnostics.Trace.WriteLine(ex.Message + Environment.NewLine + ex.StackTrace)
            Finally
                Marshal.FreeHGlobal(ip)
            End Try

        End Sub
        ''' <summary>
        ''' Set the collapse state for a single listviewgroup
        ''' </summary>
        ''' <param name="lvw"></param>
        ''' <param name="GroupName"></param>
        ''' <param name="State"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetGroupCollapse(ByVal lvw As ListView, ByVal GroupName As String, ByVal State As GroupState)

            For i As Integer = 0 To lvw.Groups.Count - 1
                If lvw.Groups(i).Name = GroupName Then
                    SetGroupCollapse(lvw, i, State)
                    Exit For
                End If
            Next

        End Sub
        ''' <summary>
        ''' Set the collapse state for all listviewgroup
        ''' </summary>
        ''' <param name="lvw"></param>
        ''' <param name="State"></param>
        ''' <remarks></remarks>
        Public Shared Sub SetGroupCollapseAll(ByVal lvw As ListView, ByVal State As GroupState)

            For i As Integer = 0 To lvw.Groups.Count - 1
                SetGroupCollapse(lvw, i, State)
            Next

        End Sub
#End Region
#Region "Groups"
        ''' <summary>
        ''' Finds the specififed group.  Optionally creates the group
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetGroup(ByVal lvw As ListView, ByVal GroupName As String, Optional ByVal CreateGroup As Boolean = True)

            If GroupName.Trim = "" Then
                Return Nothing
            End If

            Dim ThisGroup As ListViewGroup = Nothing
            Dim bFound As Boolean = False


            'see if the group already exisits
            Try
                ThisGroup = lvw.Groups(GroupName)
                bFound = True
            Catch ex As Exception

            End Try

            If ThisGroup Is Nothing Then
                bFound = False
            End If

            If bFound = False Then
                ThisGroup = lvw.Groups.Add(GroupName, GroupName)
                'sort the list
                Dim o As New ListViewGroupSorter(lvw)
                o.SortGroups(SortOrder.Ascending)
            End If

            Return ThisGroup

        End Function
#End Region
#Region "Export"
        ''' <summary>
        ''' Export the data in a listview as text
        ''' </summary>
        ''' <param name="lvw"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ExportListView(ByVal lvw As ListView, Optional ByVal ColumnSeperator As String = vbTab, Optional ByVal LineSeperator As String = vbCrLf) As String

            Dim S As String = ""
            'export the column headers:
            For i As Integer = 0 To lvw.Columns.Count - 1

                Dim aCol As ColumnHeader = lvw.Columns(i)
                S = S & aCol.Text
                If i = lvw.Columns.Count - 1 Then
                    S = S & LineSeperator
                Else
                    S = S & ColumnSeperator
                End If
            Next


            'export the items:
            For Each anItem As ListViewItem In lvw.Items

                For i As Integer = 0 To lvw.Columns.Count - 1
                    S = S & anItem.SubItems(i).Text
                    If i = lvw.Columns.Count - 1 Then
                        S = S & LineSeperator
                    Else
                        S = S & ColumnSeperator
                    End If
                Next

            Next
            Return S

        End Function

#End Region


        Public Sub ResizeHandler(ByVal sender As Object, ByVal e As System.EventArgs)

            'resize the listview by the normal algorythm
            Static InProgress As Boolean

            If InProgress Then
                Exit Sub
            End If

            Dim lvw As ListView = sender

            'assumes that all columns are visible
            'just resize them equally:
            'we can store a factor in the tag of the column?
            InProgress = True
            For i As Integer = 0 To lvw.Columns.Count - 1
                lvw.Columns(i).Width = lvw.ClientRectangle.Width / lvw.Columns.Count
            Next
            InProgress = False
        End Sub

    End Class

End Namespace