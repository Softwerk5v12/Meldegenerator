Public Class CPUAuswahl

    Property Namen As New List(Of String)

    Property Rückgabe As String




    Private Sub CB_CPUAuswahl_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Rückgabe = CB_CPUAuswahl.SelectedItem.ToString()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        CB_CPUAuswahl.Items.Add(Namen)
    End Sub
End Class
