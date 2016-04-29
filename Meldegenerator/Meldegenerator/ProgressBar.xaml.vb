Imports Siemens.Engineering
Imports Siemens.Engineering.Hmi
Imports Siemens.Engineering.HW
Imports Siemens.Engineering.SW
Imports Siemens.Engineering.Compiler
Imports System.IO
Imports Microsoft.Win32
Imports System.Reflection
Imports System





Public Class ProgressBar



    Private Sub UserControl_Initialized(sender As Object, e As EventArgs)




    End Sub

    Private Sub UserControl_Loaded(sender As Object, e As RoutedEventArgs)
        AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf MyResolveEventHandler
    End Sub


    Private Sub Rectangle_MouseLeftButtonDown(sender As Object, e As RoutedEventArgs)



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
                End If

            End If
        Next

    End Function



    Public Sub ExportvonTIA(ByVal Pfad As String)
        'Dim MyTiaPortal As New TiaPortal
        'Dim MyProjekt As Project



        PGBarDaten.Value = 0


        Dim MyTiaPortal = New TiaPortal(TiaPortalMode.WithoutUserInterface)


        Dim MyProjekt = MyTiaPortal.Projects.Open(Pfad)


        Dim CPUA As New CPUAuswahl




        Dim Ausgewaehlte_CPU_Objekt As ControllerTarget

        Dim Ausgewaehlte_CPU_Liste As New List(Of ControllerTarget)



        For Each Device In MyProjekt.Devices


            If Device.Name.Contains("S71500") Or Device.Name.Contains("S71200") Then

                Dim devitemAggregation As IDeviceItemAggregation
                Dim devitemassosiation As IDeviceItemAssociation
                Dim devitem As IDeviceItem

                devitemAggregation = Device.DeviceItems
                devitemassosiation = Device.Elements

                'CPUs im Projekt suchen
                For Each devitem In devitemAggregation
                    If devitem.TypeName.Contains("CPU") And devitem.Name IsNot vbNullString Then

                        CPUA.Namen.Add(devitem.Name)
                        Ausgewaehlte_CPU_Liste.Add(devitem)

                    End If
                Next

            End If

        Next


        'wirden mehere CPUs im Projekt gefunden, Dropdown-Auswahl öffnen

        If CPUA.Namen.Count > 1 Then
            CPUA.ShowDialog()

            Ausgewaehlte_CPU_Objekt = Ausgewaehlte_CPU_Liste.ElementAt(CPUA.Rückgabe)
        ElseIf CPUA.Namen.First IsNot vbNullString Then

            Ausgewaehlte_CPU_Objekt = Ausgewaehlte_CPU_Liste.ElementAt(0)
        End If






        'Exportordner Erstellen und Pfad vorgeben


        Dim XML_pfad As String


        XML_pfad = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) & "\Meldegenerator_XML"


        System.IO.Directory.CreateDirectory(XML_pfad)
        System.IO.Directory.CreateDirectory(XML_pfad & "\Datentypen")


        'alle Datentypen löschen
        For Each file In System.IO.Directory.GetFiles(XML_pfad & "\Datentypen")

            System.IO.File.Delete(file)

        Next

        'leeren Ordner Datentypen löschen

        System.IO.File.Delete(XML_pfad & "\Meldungen.xml")














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
                End If
            End If

        Next





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
                    End If
                Next
            End If
        Next





        MyProjekt.Close()
        MyTiaPortal.Dispose()


    End Sub


End Class
