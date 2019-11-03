Module ThemeManager
    Private currentLoadedTheme As String = String.Empty

    Public Function SystemTheme() As String
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1) Then
            Return Themes(Theme.Light)
        Else
            Return Themes(Theme.Dark)
        End If
    End Function

    Public Property AppAccentColor() As Color
        Get
            Return My.Settings.AppAccent
        End Get
        Set(ByVal value As Color)
            My.Settings.AppAccent = value
            My.Settings.Save()
        End Set
    End Property

    Public Function SystemAccentColor() As Color
        Dim accent As Integer = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\DWM", "AccentColor", -1)
        If accent > -1 Then
            Return Color.FromArgb(GetArgb(accent))
        Else
            Return AppAccentColor
        End If
    End Function

    Private Function GetArgb(abgr As Integer) As Integer
        Return ((abgr >> 24) << 24) _
            Or ((abgr >> 16) And &HFF) _
            Or ((abgr >> 8) And &HFF) << 8 _
            Or ((abgr) And &HFF) << 16
    End Function

    Public Enum Theme
        System
        Light
        Dark
    End Enum

    Public ReadOnly Themes() As String = {"System", "Light", "Dark"}

    Public Sub SetAppTheme(ByVal theme As String)
        My.Settings.AppTheme = theme
        My.Settings.Save()
    End Sub

    Public Function AppTheme() As String
        Return My.Settings.AppTheme
    End Function

    ''' <summary>
    ''' Theme specified form and its controls by app settings.
    ''' </summary>
    ''' <param name="form">Form to apply theme.</param>
    Public Sub ThemeForm(ByVal form As Form)
        Dim tAppTheme As String = AppTheme()
        If tAppTheme = Themes(Theme.System) Then
            tAppTheme = SystemTheme()
        End If

        If currentLoadedTheme IsNot tAppTheme Then
            LoadTheme($"{CurDir()}\themes\{tAppTheme}.xml")
        End If

        form.BackColor = themeColours("System.Windows.Forms.Form")("BackColor")
        form.ForeColor = themeColours("System.Windows.Forms.Form")("ForeColor")
        ThemeControls(form.Controls)
    End Sub

    Private themeColours As New Dictionary(Of String, Dictionary(Of String, Color))

    ' Load theme colours from a .xml file
    Private Sub LoadTheme(ByVal themeFile As String)
        themeColours.Clear()
        Dim themeDoc As XElement = XElement.Load(themeFile)

        For Each n In themeDoc.Elements
            Dim tDict As New Dictionary(Of String, Color)
            For Each nc In n.Elements
                tDict.Add(nc.Name.LocalName, ColorTranslator.FromHtml(nc.Value))
            Next
            themeColours.Add(n.Name.LocalName, tDict)
        Next

        currentLoadedTheme = Text.RegularExpressions.Regex.Match(themeFile, "[^\\]+?(?=\.xml)").Value
    End Sub

    Private Sub ThemeControls(ByVal controls As Control.ControlCollection)
        For Each control In controls
            Try
                control.BackColor = themeColours(control.GetType.FullName)("BackColor")
            Catch ex As Exception
                control.BackColor = control.DefaultBackColor
            End Try

            Try
                control.ForeColor = themeColours(control.GetType.FullName)("ForeColor")
            Catch ex As Exception
                control.ForeColor = control.DefaultForeColor
            End Try

            If control.GetType.FullName = "System.Windows.Forms.GroupBox" Then
                ThemeControls(control.Controls)
            End If
        Next
    End Sub
End Module
