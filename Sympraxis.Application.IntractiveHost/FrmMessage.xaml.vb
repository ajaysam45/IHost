Public Class FrmMessage
    Public Property frmResult As String = ""
    Public Property lblMessage As String = ""
    Public Property _Result As String = ""
    Private Sub BtnReadonly_Click(sender As Object, e As RoutedEventArgs)
        _Result = "btnReadonly"
        Me.Close()
    End Sub

    Private Sub BtnNotify_Click(sender As Object, e As RoutedEventArgs)
        _Result = "btnNotify"
        Me.Close()
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        _Result = "btnCancel"
        Me.Close()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        txtMessage.Text = lblMessage
        BtnReadonly.Focus()
    End Sub
End Class
