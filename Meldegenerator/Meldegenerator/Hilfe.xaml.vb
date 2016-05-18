Public Class Hilfe
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        WebView.Navigate(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString & "\Meldegenerator_HMI_Alarms\Doku.pdf")
    End Sub

    Private Sub Window_Closing(sender As Object, e As ComponentModel.CancelEventArgs)
        WebView.Dispose()
    End Sub
End Class
