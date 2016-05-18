Imports System.Windows.Threading
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports System.Reflection
Imports System.ComponentModel
Imports Siemens.Engineering
Imports Siemens.Engineering.Hmi
Imports Siemens.Engineering.HW
Imports Siemens.Engineering.SW
Imports Siemens.Engineering.Compiler
Imports System



Class MainWindow
    Property Version As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
    Dim WithEvents bgw As New BackgroundWorker
    Dim generiere_excel As New XML
    Public Übergabeparameter As New List(Of Object)
    Dim Abbruch As Boolean





    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

        Ordner_öffnen.IsEnabled = False

        'erlaubt zugriff auf die windows form
        bgw.WorkerReportsProgress = True
        'erlaubt unterbrechung des bgw z.b. wenn das programm beendet wird
        bgw.WorkerSupportsCancellation = True


        AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf MyResolveEventHandler


    End Sub



    Function MyResolveEventHandler(ByVal sender As Object,
                               ByVal args As ResolveEventArgs) As [Assembly]
        'This handler is called only when the common language runtime tries to bind to the assembly and fails.        

        'Retrieve the list of referenced assemblies in an array of AssemblyName.
        Dim objExecutingAssemblies As [Assembly]
        objExecutingAssemblies = [Assembly].GetExecutingAssembly()
        Dim arrReferencedAssmbNames() As AssemblyName
        arrReferencedAssmbNames = objExecutingAssemblies.GetReferencedAssemblies()

        'Loop through the array of referenced assembly names.
        Dim strAssmbName As AssemblyName
        For Each strAssmbName In arrReferencedAssmbNames

            'Look for the assembly names that have raised the "AssemblyResolve" event.
            If (strAssmbName.FullName.Substring(0, strAssmbName.FullName.IndexOf(",")) = args.Name.Substring(0, args.Name.IndexOf(","))) Then

                Dim filePathReg As RegistryKey

                filePathReg = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Wow6432Node\\Siemens\\Automation\\_InstalledSW\\TIAP13\\TIA_Opns")

                If filePathReg Is Nothing Then
                    filePathReg = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Siemens\\Automation\\_InstalledSW\\TIAP13\\TIA_Opns")
                End If


                'Build the path of the assembly from where it has to be loaded.
                Dim strTempAssmbPath As String
                strTempAssmbPath = filePathReg.GetValue("Path").ToString() & "\PublicAPI\V13 SP1\" & args.Name.Substring(0, args.Name.IndexOf(",")) & ".dll"


                If File.Exists(strTempAssmbPath) Then

                    Dim MyAssembly As [Assembly]

                    'Load the assembly from the specified path. 
                    MyAssembly = [Assembly].LoadFrom(strTempAssmbPath)

                    'Return the loaded assembly.
                    Return MyAssembly
                Else
                    MsgBox("TIA-Openness nicht installiert!, Bitte installieren", MsgBoxStyle.Critical)
                End If

            ElseIf strAssmbName.FullName.Contains("Sienems") Then
                MsgBox("TIA-Openness nicht installiert!, Bitte installieren", MsgBoxStyle.Critical)
            End If
        Next

    End Function



    '  Dim _xml As New XML

    Private Sub ProjektÖffnen_Click(sender As Object, e As RoutedEventArgs)
        Dim OFD As New System.Windows.Forms.OpenFileDialog With {.Multiselect = False, .Filter = "TIA files (*.ap13)|*.ap13"}

        OFD.ShowDialog()


        If Not OFD.FileName = "" Then


            Try
                bgw.RunWorkerAsync(OFD.FileName)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try




        Else
            MsgBox("Generierung Abgebrochen")
        End If

        ' _xml.RunXML()

    End Sub


    Dim MyProjekt As Project
    Dim MyTiaPortal As TiaPortal
    Dim TIAoffen As Boolean = False
    Dim TIAProjektoffen As Boolean = False



    Public Sub bgw_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw.DoWork



        Dim Err_Meldebaustein As Boolean = True

        Dim pfad As String = Convert.ToString(e.Argument)

        'Report ausgeben, prüfen ob abgebrochen wurde
        bgw.ReportProgress(10)
        System.Threading.Thread.Sleep(100)
        If Abbruch = True Then
            GoTo Abgebrochen
        End If
        '++++++++++++++++++++++++++++++++++++++++++++


        MyTiaPortal = New TiaPortal(TiaPortalMode.WithoutUserInterface)
        TIAoffen = True

        'Report ausgeben, prüfen ob abgebrochen wurde
        bgw.ReportProgress(20)
        System.Threading.Thread.Sleep(100)
        If Abbruch = True Then
            GoTo Abgebrochen
        End If
        '++++++++++++++++++++++++++++++++++++++++++++

        MyProjekt = MyTiaPortal.Projects.Open(pfad)
        TIAProjektoffen = True


        Dim CPU_Namen As New List(Of String)



        Dim Ausgewaehlte_CPU_Objekt As ControllerTarget

        Dim Ausgewaehlte_CPU_Liste As New List(Of ControllerTarget)

        'Report ausgeben, prüfen ob abgebrochen wurde
        bgw.ReportProgress(30)
        System.Threading.Thread.Sleep(100)
        If Abbruch = True Then
            GoTo Abgebrochen
        End If
        '++++++++++++++++++++++++++++++++++++++++++++



        For Each Device In MyProjekt.Devices


            'If Device.TypeName.Contains("1500") Or Device.TypeName.Contains("1200") Or Device.TypeName.Contains("300") Or Device.TypeName.Contains("400") Then

            Dim devitemAggregation As IDeviceItemAggregation
                Dim devitemassosiation As IDeviceItemAssociation
                Dim devitem As IDeviceItem

                devitemAggregation = Device.DeviceItems
                devitemassosiation = Device.Elements

                'Dim index As Integer = 0

                'CPUs im Projekt suchen
                For Each devitem In devitemAggregation
                    If devitem.TypeName.Contains("CPU") And devitem.Name IsNot vbNullString Then

                        CPU_Namen.Add(devitem.Name)
                        Ausgewaehlte_CPU_Liste.Add(devitem)

                        'index = index + 1

                    End If
                Next

            'End If

        Next


        'Report ausgeben, prüfen ob abgebrochen wurde
        bgw.ReportProgress(40)
        System.Threading.Thread.Sleep(100)
        If Abbruch = True Then
            GoTo Abgebrochen
        End If
        '++++++++++++++++++++++++++++++++++++++++++++

        'Aufruf Backgroundworker Progrss changed für bearbeitung der Oberfläche


        Übergabeparameter.Clear()



        Übergabeparameter.Add("CPU_ausw")
        Übergabeparameter.Add(CPU_Namen)




        bgw.ReportProgress(45)
        Do Until Übergabeparameter.Count = 3
            System.Threading.Thread.Sleep(5000)
        Loop




        Dim CPU_Nr As Integer


        CPU_Nr = Übergabeparameter.Item(2)

        Übergabeparameter.Clear()

        Ausgewaehlte_CPU_Objekt = Ausgewaehlte_CPU_Liste.ElementAt(CPU_Nr)




        'Exportordner Erstellen und Pfad vorgeben

        'Report ausgeben, prüfen ob abgebrochen wurde
        bgw.ReportProgress(50)
        System.Threading.Thread.Sleep(100)
        If Abbruch = True Then
            GoTo Abgebrochen
        End If
        '++++++++++++++++++++++++++++++++++++++++++++

        Dim XML_pfad As String


        XML_pfad = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Meldegenerator_XML"


        System.IO.Directory.CreateDirectory(XML_pfad)
        System.IO.Directory.CreateDirectory(XML_pfad & "\Datentypen")


        'alle Datentypen löschen
        For Each file In System.IO.Directory.GetFiles(XML_pfad & "\Datentypen")

            System.IO.File.Delete(file)

        Next

        'leeren Ordner Datentypen löschen

        System.IO.File.Delete(XML_pfad & "\Meldungen.xml")




        'Report ausgeben, prüfen ob abgebrochen wurde
        bgw.ReportProgress(60)
        System.Threading.Thread.Sleep(100)
        If Abbruch = True Then
            GoTo Abgebrochen
        End If
        '++++++++++++++++++++++++++++++++++++++++++++

        'Baustein Suchen und Exportieren

        Dim Bausteinordner As ProgramblockSystemFolder



        Bausteinordner = Ausgewaehlte_CPU_Objekt.ProgramblockFolder

        For Each Baustein In Bausteinordner.Blocks

            If Baustein.Name = "Meldungen" Then
                If Baustein.IsConsistent Then
                    Try
                        Baustein.Export(XML_pfad & "\" & Baustein.Name & ".xml", ExportOptions.None)

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                Else
                    MsgBox("Meldebaustein nicht übersetzt")
                    Abbruch = True
                End If
                Err_Meldebaustein = False
                generiere_excel.DBNummer = Baustein.Number
            End If

        Next

        If Err_Meldebaustein Then
            MsgBox("Kein Meldebaustein im Programm-Ordner gefunden. Bitte Baustein 'Meldungen' im Bausteinordner (ohne Unterordner) anlegen ")
            Abbruch = True
        End If


        'Report ausgeben, prüfen ob abgebrochen wurde
        bgw.ReportProgress(65)
        System.Threading.Thread.Sleep(100)
        If Abbruch = True Then
            GoTo Abgebrochen
        End If
        '++++++++++++++++++++++++++++++++++++++++++++


        'Datentypen exportieren
        Dim Datentyp_Systemordner As ControllerDatatypeSystemFolder
        Dim i As Int32

        Datentyp_Systemordner = Ausgewaehlte_CPU_Objekt.ControllerDatatypeFolder

        i = 1

        For Each Folder In Datentyp_Systemordner.Folders
            If Folder.Name = "Meldungen" Then
                For Each Datatype In Folder.Datatypes
                    If Datatype.IsConsistent Then
                        Try
                            Datatype.Export(XML_pfad & "\Datentypen\Datentyp_" & i & ".xml", ExportOptions.None)
                            i = i + 1
                        Catch ex As Exception
                            MsgBox(ex.ToString)
                        End Try
                    Else
                        MsgBox("Datentyp: " & Datatype.Name & " nicht übersetzt")
                        Abbruch = True
                    End If
                Next
            End If
        Next


        'Report ausgeben, prüfen ob abgebrochen wurde
        bgw.ReportProgress(70)
        System.Threading.Thread.Sleep(100)
        If Abbruch = True Then
            GoTo Abgebrochen
        End If
        '++++++++++++++++++++++++++++++++++++++++++++

        MyProjekt.Close()
        TIAProjektoffen = False

        'Report ausgeben, prüfen ob abgebrochen wurde
        bgw.ReportProgress(75)
        System.Threading.Thread.Sleep(100)
        If Abbruch = True Then
            GoTo Abgebrochen
        End If
        '++++++++++++++++++++++++++++++++++++++++++++

        MyTiaPortal.Dispose()
        TIAoffen = False

        'Ausführung der Generierung der Excel-Tabelle

        generiere_excel.CPUnummer = (CPU_Nr + 1)
        generiere_excel.CPUName = CPU_Namen.Item(CPU_Nr)

        'Report ausgeben, prüfen ob abgebrochen wurde
        bgw.ReportProgress(90)
        System.Threading.Thread.Sleep(100)
        If Abbruch = True Then
            GoTo Abgebrochen
        End If
        '++++++++++++++++++++++++++++++++++++++++++++


        Try
            generiere_excel.RunXML()
        Catch ex As Exception
            MsgBox(ex.ToString)
            bgw.CancelAsync()
        End Try



        bgw.ReportProgress(95)
        System.Threading.Thread.Sleep(100)




Abgebrochen:
        If Abbruch = True Then

            If TIAProjektoffen Then
                MyProjekt.Close()

            End If

            If TIAoffen Then
                MyTiaPortal.Dispose()

            End If

        End If




    End Sub






    Public Sub bgw_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles bgw.ProgressChanged
        Ordner_öffnen.IsEnabled = False

        Dim CPUA As New CPUAuswahl

        If Übergabeparameter.Count > 0 Then
            If Übergabeparameter.Item(0) = "CPU_ausw" Then
                Dim Namensliste As New List(Of String)
                Namensliste = Übergabeparameter.Item(1)

                CPUA.Namen = Namensliste
                'Wurden Mehrere CPUs im Projekt gefunden, muss eine ausgewählt werden. Ansonsten wird automatisch die eine CPU genommen
                If Namensliste.Count > 1 Then
                    CPUA.ShowDialog()
                    Übergabeparameter.Add(CPUA.Rückgabe)
                Else
                    Übergabeparameter.Add(0)
                End If

            End If
        End If

        If Abbruch = False Then
            If e.ProgressPercentage = 10 Then
                PBar.LBAnzahlTITEL.Content = "Öffne TIA"
                ProjektÖffnen.IsEnabled = False
            End If

            If e.ProgressPercentage = 20 Then
                PBar.LBAnzahlTITEL.Content = "Öffne Projekt in TIA"
            End If

            If e.ProgressPercentage = 30 Then
                PBar.LBAnzahlTITEL.Content = "Suche Steuerungen im Projekt"
            End If

            If e.ProgressPercentage = 40 Then
                PBar.LBAnzahlTITEL.Content = "Steuerungen listen / auswählen"
            End If

            If e.ProgressPercentage = 45 Then
                PBar.LBAnzahlTITEL.Content = "Gewählte Stuerung wird ausgelesen"
            End If

            If e.ProgressPercentage = 50 Then
                PBar.LBAnzahlTITEL.Content = "Erstelle XML Export-Pfad"
            End If

            If e.ProgressPercentage = 60 Then
                PBar.LBAnzahlTITEL.Content = "Exportiere Baustein"
            End If

            If e.ProgressPercentage = 65 Then
                PBar.LBAnzahlTITEL.Content = "Exportiere Datentypen"
            End If

            If e.ProgressPercentage = 70 Then
                PBar.LBAnzahlTITEL.Content = "Schließe Projekt in TIA"
            End If

            If e.ProgressPercentage = 75 Then
                PBar.LBAnzahlTITEL.Content = "Schließe TIA Portal"
            End If

            If e.ProgressPercentage = 90 Then
                PBar.LBAnzahlTITEL.Content = "XML´s auslesen und Excel-Tabelle erstellen"
            End If
        End If

        PBar.PGBarDaten.Value = e.ProgressPercentage





    End Sub



    Private Sub bgw_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw.RunWorkerCompleted
        If PBar.PGBarDaten.Value = 0 Or PBar.PGBarDaten.Value = 10 Or PBar.PGBarDaten.Value = 20 And Not Abbruch Then
            MsgBox("Fehler:" & vbNewLine & "- TIA Openness nicht installiert" & vbNewLine & "- TIA Openness Benutzereinstellungen fehlen" & vbNewLine & "- .Net Framework 4.0 oder höher installieren")
            Abbruch = True
        ElseIf PBar.PGBarDaten.Value = 95 Then

            PBar.LBAnzahlTITEL.Content = "Fertig"
            PBar.PGBarDaten.Value = 100
            Ordner_öffnen.IsEnabled = True
            TB_HMIVariableName.Text = generiere_excel.HMIVariablenName
            TB_HMIVariableDatentyp.Text = generiere_excel.HMIVariableDatentyp
            SP_HMIVariableName.Visibility = Visibility.Visible
            SP_HMIVariableDatentyp.Visibility = Visibility.Visible
        End If

        If Abbruch = True Then

            PBar.LBAnzahlTITEL.Content = "Generierung Abgebrochen"
            PBar.PGBarDaten.Value = 100
            Abbruch = False

        End If

        TIAProjektoffen = False
        TIAoffen = False
        ProjektÖffnen.IsEnabled = True

    End Sub


    Private Sub Abbrechen_Click(sender As Object, e As RoutedEventArgs)

        Abbruch = True

        PBar.LBAnzahlTITEL.Content = "wird Abgebrochen"
        PBar.PGBarDaten.Value = 100

    End Sub



    Private Sub Ordner_öffnen_Click(sender As Object, e As RoutedEventArgs)

        System.Diagnostics.Process.Start("explorer", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Meldegenerator_HMI_Alarms")

    End Sub
End Class

