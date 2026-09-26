Imports System.Configuration
Imports System.IO

Namespace My

    Partial Friend Class MyApplication

        ' Runs before the main form is created. If the saved user settings file (user.config) is damaged, every My.Settings
        ' read throws inside Formtest's constructor and the program would die silently, so check it here and let the user decide.
        Private Sub MyApplication_Startup(sender As Object, e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) Handles Me.Startup

            Try
                Dim probe As Object = My.Settings.data29
                Exit Sub
            Catch ex As ConfigurationErrorsException
                HandleDamagedSettings(ex, e)
            Catch ex As Exception When TypeOf ex.InnerException Is ConfigurationErrorsException
                HandleDamagedSettings(DirectCast(ex.InnerException, ConfigurationErrorsException), e)
            End Try

        End Sub

        Private Sub HandleDamagedSettings(ex As ConfigurationErrorsException, e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs)

            Dim badFile As String = ex.Filename
            If String.IsNullOrEmpty(badFile) AndAlso TypeOf ex.InnerException Is ConfigurationErrorsException Then
                badFile = DirectCast(ex.InnerException, ConfigurationErrorsException).Filename
            End If

            Dim detail As String = ex.Message
            If ex.InnerException IsNot Nothing Then detail = ex.InnerException.Message

            Dim text As String =
                "WinGPIB's saved settings file is damaged and cannot be read." & vbCrLf & vbCrLf &
                "File: " & If(String.IsNullOrEmpty(badFile), "(location unknown)", badFile) & vbCrLf &
                "Problem: " & detail & vbCrLf & vbCrLf &
                "This can happen if the program or Windows stopped while settings were being saved." & vbCrLf & vbCrLf &
                "Yes  -  Rename the damaged file (it is kept with a .bad extension) and start WinGPIB with default settings. Your saved settings will be lost, but if you have a recent exported profile you can load it afterwards to restore them." & vbCrLf & vbCrLf &
                "No  -  Exit without changing anything."

            Dim answer As DialogResult = MessageBox.Show(text, "WinGPIB - Damaged settings file", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)

            If answer <> DialogResult.Yes OrElse String.IsNullOrEmpty(badFile) OrElse Not File.Exists(badFile) Then
                e.Cancel = True
                Exit Sub
            End If

            Try
                Dim keepName As String = badFile & ".bad"
                If File.Exists(keepName) Then File.Delete(keepName)
                File.Move(badFile, keepName)
                My.Settings.Reload()
            Catch ex2 As Exception
                MessageBox.Show("The damaged file could not be renamed (" & ex2.Message & ")." & vbCrLf & vbCrLf &
                                "Please delete it manually and start WinGPIB again:" & vbCrLf & badFile,
                                "WinGPIB - Damaged settings file", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
            End Try

        End Sub

    End Class

End Namespace
