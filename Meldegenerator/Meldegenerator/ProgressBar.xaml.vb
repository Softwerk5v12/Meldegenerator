Imports Siemens.Engineering
Imports Siemens.Engineering.Hmi
Imports Siemens.Engineering.HW
Imports Siemens.Engineering.Compiler
Imports System.IO
Imports Microsoft.Win32
Imports System.Reflection
Imports System





Public Class ProgressBar


    Private Sub Rectangle_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)



    End Sub

    Private Function MyResolver(sender As Object, e As ResolveEventArgs) As Assembly
        Dim index As Int32 = e.Name.IndexOf(",")
        Dim name As String = e.Name.Substring(0, index) + ".dll"


        Dim filePathReg As RegistryKey
        filePathReg = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Wow6432Node\\Siemens\\Automation\\_InstalledSW\\TIAP13\\TIA_Opns")

        If filePathReg Is Nothing Then
            filePathReg = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Siemens\\Automation\\_InstalledSW\\TIAP13\\TIA_Opns")
        End If

        Dim filePath = filePathReg.GetValue("Path").ToString() + "PublicAPI\\V13 SP1"

        Dim PathClass As Path
        Dim path As String

        path = PathClass.Combine(path1:=filePath, path2:=name)

        Dim fullPath = PathClass.GetFullPath(path)

        If File.Exists(fullPath) Then
            Return Assembly.LoadFrom(fullPath)
        End If

    End Function






    Public Sub ExportvonTIA(ByVal Pfad As String)
        'Dim MyTiaPortal As New TiaPortal
        'Dim MyProjekt As Project


        'Prüfe auf 64Bit version









        PGBarDaten.Value = 0

        Dim MyTiaPortal = New TiaPortal(TiaPortalMode.WithoutUserInterface)


        Dim MyProjekt = MyTiaPortal.Projects.Open(Pfad)




        For Each Device In MyProjekt.Devices

            MsgBox(Device.Name)


        Next


        MyProjekt.Close()
        MyTiaPortal.Dispose()


    End Sub








End Class
