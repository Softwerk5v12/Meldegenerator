Imports System.Windows.Threading
Imports System.IO
Imports System.Windows.Forms

Class MainWindow

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)






        '  Dispatcher.BeginInvoke(Sub() PGbar.DoWork(), DispatcherPriority.Background)                                 )

    End Sub




    Private Sub TiaProjektauswählen()
        ' Dim OR As open


        'PGbar.OpenProjekt()


    End Sub

    Private Sub ProjektÖffnen_Click(sender As Object, e As RoutedEventArgs)
        Dim OFD As New OpenFileDialog With {.Multiselect = False, .Filter = "TIA files (*.ap13)|*.ap13"}

        OFD.ShowDialog()


        MsgBox(OFD.FileName)

    End Sub

    Dim _XML As New XML
    Private Sub BT_XML_Click(sender As Object, e As RoutedEventArgs)
        _XML.LoadXML()
    End Sub

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        _XML.Write_Excel()
    End Sub
End Class

