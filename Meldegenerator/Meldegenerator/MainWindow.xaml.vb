Imports System.Windows.Threading
Imports System.IO
Imports System.Windows.Forms

Class MainWindow

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

        AddHandler _XML.StatusChaged, AddressOf StatusChanged





        '  Dispatcher.BeginInvoke(Sub() PGbar.DoWork(), DispatcherPriority.Background)                                 )

    End Sub

    Private Sub StatusChanged()
        PBar.LBAnzahlTITEL.Content = _XML.Status
        Me.Dispatcher.Invoke(Windows.Threading.DispatcherPriority.ContextIdle, Function() 0)
    End Sub
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

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
        _XML.RunXML()
    End Sub

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)

    End Sub
End Class

