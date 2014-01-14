Namespace WinForms
    Public Class FormHelper

        Private Shared Sub SaveSetting(ByVal SettingName As String, ByVal Settings As Global.System.Configuration.ApplicationSettingsBase, ByVal value As Object)

            Dim bSettingAvailable As Boolean = False
            Try
                Dim oTemp As Object = Settings.Properties(SettingName)
                bSettingAvailable = True
            Catch ex As Exception

            End Try

            If bSettingAvailable Then
                Settings.Item(SettingName) = value
            Else
                Debug.Assert(False, "Setting is not defined for application: " & SettingName)
            End If

        End Sub

        Private Shared Function GetSetting(ByVal SettingName As String, ByVal Settings As Global.System.Configuration.ApplicationSettingsBase) As Object

            Dim bSettingAvailable As Boolean = False

            Try
                Dim oTemp As Object = My.Settings.Properties(SettingName)
                bSettingAvailable = True
            Catch ex As Exception
                Debug.Write("")
            End Try

            If bSettingAvailable Then
                Return Settings(SettingName)
            Else
                Debug.Assert(False, "Setting is not defined for application: " & SettingName)
                Return Nothing
            End If

        End Function

        ''' <summary>
        ''' Save the common settings for forms such as location, size etc
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub SaveSettings(ByVal aForm As Windows.Forms.Form, ByVal Settings As Global.System.Configuration.ApplicationSettingsBase)

            SaveSetting(aForm.Name & "_Size", Settings, aForm.Size)
            SaveSetting(aForm.Name & "_Location", Settings, aForm.Location)
            SaveSetting(aForm.Name & "_WindowState", Settings, CInt(aForm.WindowState))

            Settings.Save()

        End Sub

        ''' <summary>
        ''' Set the forms location and size from the project settings: Size, Location, WindowState
        ''' </summary>
        ''' <param name="aForm"></param>
        ''' <param name="Settings"></param>
        ''' <remarks></remarks>
        Public Shared Sub LoadSettings(ByVal aForm As Windows.Forms.Form, ByVal Settings As Global.System.Configuration.ApplicationSettingsBase)

            Dim oSize As System.Drawing.Size = GetSetting(aForm.Name & "_Size", Settings)
            If oSize.Height < aForm.MinimumSize.Height Then
                oSize.Height = aForm.MinimumSize.Height
            End If
            If oSize.Width < aForm.MinimumSize.Width Then
                oSize.Width = aForm.MinimumSize.Width
            End If
            aForm.Size = oSize

            Dim oPoint As System.Drawing.Point = GetSetting(aForm.Name & "_Location", Settings)
            If aForm.Location.X = 0 And aForm.Location.Y = 0 Then
            Else
                aForm.Location = oPoint
            End If

            aForm.WindowState = GetSetting(aForm.Name & "_WindowState", Settings)

        End Sub

        Public Shared Sub AddToContainer(ByVal aForm As System.Windows.Forms.Form, ByVal Container As System.Windows.Forms.Panel)

            If Container.Contains(aForm) = False Then
                aForm.TopLevel = False
                Container.Controls.Add(aForm)
                aForm.Dock = System.Windows.Forms.DockStyle.Fill
                aForm.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            End If

        End Sub

        Public Shared Sub ShowForm(ByVal Form As System.Windows.Forms.Form, Optional ByVal Owner As System.Windows.Forms.Form = Nothing)

            If My.Computer.Screen.WorkingArea.Contains(Form.Location) = False Then
                Form.Top = 0
                Form.Left = 0
            End If

            If Form.Visible = False Then


                If Owner Is Nothing Then
                    Form.Show()
                Else
                    Form.Show(Owner)
                End If
            End If

        End Sub

    End Class
End Namespace


