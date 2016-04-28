Class MainWindow
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

        Dim BT As New Button

        BT.Width = 400
        BT.Height = 200


        WPMain.Children.Add(BT)



        Dim Fred As String = "fefeff"
    End Sub

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        Dim Auswahl As New TextBox
        Auswahl.Width = 150
        Auswahl.Height = 25

        WPMain.Children.Add(Auswahl)

    End Sub
End Class
