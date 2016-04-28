Imports System.Windows.Threading
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports System.Reflection

Class MainWindow



    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)




    End Sub




    Private Sub TiaProjektauswählen()
        ' Dim OR As open


        'PGbar.OpenProjekt()


    End Sub

    Private Sub ProjektÖffnen_Click(sender As Object, e As RoutedEventArgs)
        Dim OFD As New System.Windows.Forms.OpenFileDialog With {.Multiselect = False, .Filter = "TIA files (*.ap13)|*.ap13"}


        OFD.ShowDialog()

        PBar.ExportvonTIA(OFD.FileName)


    End Sub


End Class

