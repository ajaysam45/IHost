Public Class Help
    Public HelpUrl As String

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        WebNavigate.Navigate(HelpUrl)  
    End Sub
End Class
