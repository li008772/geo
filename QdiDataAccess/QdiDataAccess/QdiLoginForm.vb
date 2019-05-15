Option Explicit On
Option Strict On


Public Class QdiLoginForm
    'Inherits ESRI.ArcGIS.Desktop.AddIns.AddInEntryPoint
    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    ''' <summary>
    ''' handles click event for OK button
    ''' </summary>
    ''' <param name="sender">sender is an object of generic type System.Object, it is passing the control that is causing the event to fire.</param>
    ''' <param name="e"> e is an object of type System.EventArgs where EventArgs is the generic Class for event arguments or, the arguments the event is passed.</param>
    Protected Overridable Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    ''' <summary>
    ''' handles click event for Cancel button
    ''' </summary>
    ''' <param name="sender">sender is an object of generic type System.Object, it is passing the control that is causing the event to fire.</param>
    ''' <param name="e"> e is an object of type System.EventArgs where EventArgs is the generic Class for event arguments or, the arguments the event is passed.</param>
    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.UsernameTextBox.Text = ""
        'Me.PasswordTextBox.Text = ""
        Me.Close()
    End Sub

    ''' <summary>
    ''' gets the username from os and fills the text bar with it; then enables the OK button
    ''' </summary>
    ''' <param name="sender">sender is an object of generic type System.Object, it is passing the control that is causing the event to fire.</param>
    ''' <param name="e"> e is an object of type System.EventArgs where EventArgs is the generic Class for event arguments or, the arguments the event is passed.</param>
    Private Sub QdiLoginForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.OK.Enabled = False
        'If Debugger.IsAttached Then
        '    Me.UsernameTextBox.Text = My.Resources.r_DefaultLogin
        'Else
        '    Me.UsernameTextBox.Text = ""
        'End If
        Me.UsernameTextBox.Text = System.Environment.UserName
        'TextBox_TextChanged(sender, e)
        Me.OK.Enabled = True
    End Sub

    'Private Sub TextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsernameTextBox.TextChanged
    '    Me.OK.Enabled = False

    '    'If (Len(Me.PasswordTextBox.Text.Trim()) > 0) Then
    '    If (Len(Me.UsernameTextBox.Text.Trim()) > 0) Then
    '        Me.OK.Enabled = True
    '    End If
    '    'End If
    'End Sub


End Class


