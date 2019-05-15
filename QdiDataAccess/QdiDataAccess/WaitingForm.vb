Public Class frmWaitingForm

    ''' <summary>
    ''' handles the "Try Again" button
    ''' </summary>
    ''' <param name="sender">sender is an object of generic type System.Object, it is passing the control that is causing the event to fire.</param>
    ''' <param name="e"> e is an object of type System.EventArgs where EventArgs is the generic Class for event arguments or, the arguments the event is passed.</param>
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.DialogResult = Windows.Forms.DialogResult.Retry
        Me.Close()
    End Sub

    ''' <summary>
    ''' handles the "Force Save" button
    ''' </summary>
    ''' <param name="sender">sender is an object of generic type System.Object, it is passing the control that is causing the event to fire.</param>
    ''' <param name="e"> e is an object of type System.EventArgs where EventArgs is the generic Class for event arguments or, the arguments the event is passed.</param>
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim pDisconnectAnswer As Windows.Forms.DialogResult
        pDisconnectAnswer = Windows.Forms.MessageBox.Show("Are sure you want to force your edits to be saved? You will probably be interupting the saving of someone else's edits and their work will be lost. This should only be used if the other person's session crashed.", "Force Close?", Windows.Forms.MessageBoxButtons.OKCancel, Windows.Forms.MessageBoxIcon.Question)
        If (pDisconnectAnswer = MsgBoxResult.Ok) Then
            Me.DialogResult = Windows.Forms.DialogResult.Abort
            Me.Close()
        End If
    End Sub

    ''' <summary>
    ''' handles the "Cancel" button
    ''' </summary>
    ''' <param name="sender">sender is an object of generic type System.Object, it is passing the control that is causing the event to fire.</param>
    ''' <param name="e"> e is an object of type System.EventArgs where EventArgs is the generic Class for event arguments or, the arguments the event is passed.</param>
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
End Class