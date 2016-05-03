Imports System.Windows.Threading
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports System.Reflection

Class MainWindow



    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)


        AddHandler _xml.StatusChaged, AddressOf Changed

    End Sub




    Private Sub TiaProjektauswählen()
        ' Dim OR As open


        'PGbar.OpenProjekt()


    End Sub

    Private Sub Changed()
        PBar.LBAnzahlTITEL.Content = _xml.Status

        Me.Dispatcher.Invoke(DispatcherPriority.Background, Function() 0)
    End Sub

    Private Sub ProjektÖffnen_Click(sender As Object, e As RoutedEventArgs)
        Dim OFD As New System.Windows.Forms.OpenFileDialog With {.Multiselect = False, .Filter = "TIA files (*.ap13)|*.ap13"}






        OFD.ShowDialog()


        If Not OFD.FileName = "" Then
            Dispatcher.BeginInvoke(Sub() PBar.ExportvonTIA(OFD.FileName))
        Else
            MsgBox("Projekt schließen danach Meldegenerierung erneut ausführen", MsgBoxStyle.Critical)
        End If


    End Sub

    Dim _xml As New XML

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        _xml.RunXML()
    End Sub
End Class

