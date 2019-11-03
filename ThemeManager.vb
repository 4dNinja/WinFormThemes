Module ThemeManager
    ReadOnly DarkBackColour As Color = Color.FromArgb(19, 19, 19)
    ReadOnly DarkForeColour As Color = Color.FromArgb(255, 255, 255)
    ReadOnly DarkBackColourLight As Color = Color.FromArgb(31, 31, 31)
    'ReadOnly DefaultAccentColour As Color = Color.Purple

    Private storedSystemUsingLightTheme As Boolean = 1

    Public Function IsSystemUsingLightTheme() As Boolean
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1) Then
            Return 1
        Else
            Return 0
        End If
    End Function

    Public Property AppAccent() As Color
        Get
            Return My.Settings.AppAccent
        End Get
        Set(ByVal value As Color)
            My.Settings.AppAccent = value
            My.Settings.Save()
        End Set
    End Property


    Public Function GetSystemAccent() As Color
        Dim accent As Integer = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\DWM", "AccentColor", -1)
        If accent > -1 Then
            Return Color.FromArgb(GetArgb(accent))
        Else
            Return AppAccent
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

    Private ReadOnly ThemeLookup() As String = {"System", "Light", "Dark"}

    Public Sub SetAppTheme(ByVal theme As Theme)
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
        Dim appTheme As String = ThemeManager.AppTheme()

        If appTheme = Theme.System Then
            If IsSystemUsingLightTheme() Then
                appTheme = Theme.Light
            Else
                appTheme = Theme.Dark
            End If
        End If

        If appTheme = Theme.Light Then
            form.BackColor = Form.DefaultBackColor
            form.ForeColor = Form.DefaultForeColor
            SetControlsLight(form.Controls)
        Else ' Theme.Dark
            form.BackColor = DarkBackColour
            form.ForeColor = DarkForeColour
            SetControlsDark(form.Controls)
        End If
    End Sub

    Private themeColours As New Dictionary(Of String, Color)
    Private controlColours As New Dictionary(Of String, ControlColour)

    Private Sub SetColours()
        themeColours.Add("System.Windows.Forms.Button.BackColor", Button.DefaultBackColor)
        themeColours.Add("System.Windows.Forms.Button.ForeColor", Button.DefaultForeColor)

        controlColours.Add("System.Windows.Forms.Button", New ControlColour("BackColor", Button.DefaultBackColor))
        controlColours.Add("System.Windows.Forms.Button", New ControlColour("ForeColor", Button.DefaultForeColor))



        '      controlColours.Add("System.Windows.Forms.Button", CType(Dictionary(Of String, Color),"BackColor", Button.DefaultBackColor)
    End Sub

    Private Structure ControlColour
        Dim control As String
        Dim colour As Color

        Public Sub New(controlNew As String, colourNew As Color)
            Me.control = controlNew
            Me.colour = colourNew
        End Sub
    End Structure

    Private Sub SetControlsLight(ByVal ctrlCol As Control.ControlCollection)
        For Each c In ctrlCol
            Select Case c.GetType.ToString
                Case "System.Windows.Forms.Button"
                    c.BackColor = Button.DefaultBackColor
                    c.ForeColor = Button.DefaultForeColor
                Case "System.Windows.Forms.Label"
                    c.BackColor = Label.DefaultBackColor
                    c.ForeColor = Label.DefaultForeColor
                Case "System.Windows.Forms.GroupBox"
                    c.BackColor = GroupBox.DefaultBackColor
                    c.ForeColor = GroupBox.DefaultForeColor
                    SetControlsLight(c.Controls)
                Case "System.Windows.Forms.TextBox"
                    c.BackColor = Color.FromKnownColor(KnownColor.ControlLight)
                    c.ForeColor = Color.FromKnownColor(KnownColor.WindowText)
                Case Else
                    c.BackColor = Form.DefaultBackColor
                    c.ForeColor = Form.DefaultForeColor
            End Select
        Next
    End Sub

    Private Sub SetControlsDark(ByVal controls As Control.ControlCollection)
        For Each control In controls
            Select Case control.GetType.ToString
                Case "System.Windows.Forms.Button"
                    control.BackColor = DarkBackColour
                    control.ForeColor = DarkForeColour
                Case "System.Windows.Forms.Label"
                    control.BackColor = DarkBackColour
                    control.ForeColor = DarkForeColour
                Case "System.Windows.Forms.GroupBox"
                    control.BackColor = DarkBackColour
                    control.ForeColor = DarkForeColour
                    SetControlsDark(control.Controls)
                Case "System.Windows.Forms.RadioButton"
                    control.BackColor = DarkBackColour
                    control.ForeColor = DarkForeColour
                Case Else
                    control.BackColor = DarkBackColourLight
                    control.ForeColor = DarkForeColour
            End Select
        Next
    End Sub
End Module
