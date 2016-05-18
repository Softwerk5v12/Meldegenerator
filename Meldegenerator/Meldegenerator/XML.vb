Imports System.IO
Imports System.Environment

Imports Excel = Microsoft.Office.Interop.Excel






Public Class XML


    Const Column0_A As String = "ID"
    Const Column1_B As String = "Name"
    Const Column2_C As String = "Event text [de-DE], Alarm text"
    Const Column3_D As String = "FieldInfo [Alarm text]"
    Const Column4_E As String = "Class"
    Const Column5_F As String = "Trigger tag"
    Const Column6_G As String = "Trigger bit"
    Const Column7_H As String = "Acknowledgement tag"
    Const Column8_I As String = "Acknowledgement bit"
    Const Column9_J As String = "PLC acknowledgement tag"
    Const Column10_K As String = "PLC acknowledgement bit"
    Const Column11_L As String = "Group"
    Const Column12_M As String = "Report"
    Const Column13_N As String = "Info text [de-DE], Info text"



    Public Event StatusChaged()
    Private Sub _StatusChanged(ByVal _Status As String)
        Status = _Status
        RaiseEvent StatusChaged()
    End Sub

    Property CPUnummer As Integer = 1
    Property DBNummer As Integer = 260
    Property CPUName As String = ""
    Property Status As String
    Property HMIVariableDatentyp As String = ""
    Property HMIVariablenName As String = ""
    Dim Meldungen As New List(Of HMIAlarms)


    Dim Datentypen As New List(Of HMIAlarms)

    Dim TagName As String = "Trigger_AT_" & CPUnummer.ToString & "_DB"


    Public Sub RunXML()
        Dim XMLFile As XDocument
        _StatusChanged("Load XML")
        HMIVariableDatentyp = "super"
        For Each file As String In Directory.GetFiles(GetFolderPath(SpecialFolder.MyDocuments) & "\Meldegenerator_XML\Datentypen")
            GetDatatyp(file)
        Next

        Try
            XMLFile = XDocument.Load(GetFolderPath(SpecialFolder.MyDocuments) & "\Meldegenerator_XML\Meldungen.xml")


            Dim Interface_Sections As XElement =
            (From el In XMLFile.<Document>.<SW.DataBlock>.<AttributeList>.<Interface>
             Select el).First
            '   MsgBox(XMLFile.Elements.Count)
            GetHMIMeldungen(Interface_Sections)

        Catch ex As Exception
            _StatusChanged("Fehler beim XML öffnen")
        End Try

        HMIVariableDatentyp = "Array [0.." & AddressWord & "] Of Word"
        HMIVariablenName = AddressTag

        CreateWorkbook()
        'XMLFile.Root.Remove()
        '      XMLFile.Save(GetFolderPath(SpecialFolder.MyDocuments) & "\Meldegenerator_XML\Meldungen.xml")
        Write_Excel()



    End Sub


    Private AddressWord As Integer = -1
    ' Private AddressName As Integer = -1
    Private AddressBit As Integer = 7
    Private AddressBitforArray As Integer
    Private AddressTag As String

    Private Sub CountDBAdresse()

        _StatusChanged("Calculate DB Address")



        If AddressBit >= 15 Then
            AddressBit = 0

        Else
            AddressBit = AddressBit + 1
        End If

        If AddressBit = 8 Then
            AddressWord = AddressWord + 1
            '  AddressName = AddressName + 1
        End If
        AddressBitforArray = AddressBit * AddressWord
        TagName = "Trigger_AT_" & CPUnummer.ToString & "_DB"

        AddressTag = """" & TagName & DBNummer & """"
    End Sub


    Dim ID As Integer = CPUnummer * 10000
    Dim SiemensNamespace As XNamespace = "http://www.siemens.com/automation/Openness/SW/Interface/v1" ' must match declaration In document
    Private Sub GetHMIMeldungen(ByVal Interface_Sections As XElement)
        _StatusChanged("XML initialisieren")


        '  CountDBAdresse()

        Dim SelectionsElemente As IEnumerable(Of XElement) =
        From element In Interface_Sections.Elements(SiemensNamespace + "Sections")
        Select element

        Dim SectionElement As XElement = SelectionsElemente.ElementAt(0)


        Dim Selection = From element In SelectionsElemente.Elements(SiemensNamespace + "Section") Select element

        ID = CPUnummer * 10000

        Dim Meldeklassen = (From element In Selection.Elements(SiemensNamespace + "Member") Select element)
        For i As Integer = 0 To Meldeklassen.Count - 1


            Dim AktuelleMeldeklasse = Meldeklassen.ElementAt(i).Elements

            If AktuelleMeldeklasse.First.Parent.FirstAttribute.Value Like "M_*" Then
                MeldungenGenerieren(AktuelleMeldeklasse)
                AddressBit = 6
                CountDBAdresse()
            ElseIf AktuelleMeldeklasse.First.Parent.FirstAttribute.Value Like "S_*" Then
                MeldungenGenerieren(AktuelleMeldeklasse)
                AddressBit = 6
                CountDBAdresse()
            Else

                MsgBox("DIe Meldeklasse ist falsch benannt, der Klassenname muss mit ""M_"" oder ""S_"" beginnen")
            End If

        Next

    End Sub



    Private Sub MeldungenGenerieren(ByVal Meldeklasse As IEnumerable(Of XElement))
        Dim MeldungsBoolafter_others As Boolean = False
        Dim MeldungAlarmtext As String = Nothing
        Dim MeldungStructName As String = Nothing
        Dim Meldungcounter As Integer = 0
        _StatusChanged("Meldungen aus XML lesen")

        Dim Meldeklassenname As String = Meldeklasse.First.Value
        'Dim Meldeklassenname As String = ""
        For Each Meldung As XElement In Meldeklasse

            '  MsgBox(Meldung.Name.ToString)
            If Meldung.Name = "{" & SiemensNamespace.ToString & "}Member" Then



                If Meldung.@Datatype.ToString = "Bool" Then

                    If MeldungsBoolafter_others = True Then
                        AddressBit = 6
                        CountDBAdresse()
                        MeldungsBoolafter_others = False
                    End If
                    CountDBAdresse()
                    Meldungen.Add(New HMIAlarms With {.AlarmText = Meldung.Descendants(SiemensNamespace + "MultiLanguageText").Value,
                                      .Meldeklasse = Meldeklassenname, .Name = Meldung.FirstAttribute.Value & ID, .Datentyp = Meldung.@Datatype.ToString,
                                      .ID = ID, .TriggerTag = AddressTag, .TrigerBit = AddressBitforArray})

                    ID = ID + 1

                ElseIf Meldung.@Datatype.ToString = "Struct" Then
                    AddressBit = 6
                    CountDBAdresse()

                    MeldungsBoolafter_others = True
                    MeldungStructName = Meldung.FirstAttribute.Value

                    Dim StructElement = (From element In Meldung.Nodes Select element)

                    For i As Integer = 1 To StructElement.Count - 1
                        CountDBAdresse()
                        Dim StructMeldung As XElement = StructElement.ElementAt(i)

                        Meldungen.Add(New HMIAlarms With {.AlarmText = MeldungStructName & " " & StructMeldung.Descendants(SiemensNamespace + "MultiLanguageText").Value,
                         .Meldeklasse = Meldeklassenname, .Name = Meldung.FirstAttribute.Value & ID, .Datentyp = StructMeldung.@Datatype.ToString,
                        .ID = ID, .TriggerTag = AddressTag, .TrigerBit = AddressBitforArray})
                        ID = ID + 1

                        Meldungcounter = Meldungcounter + 1

                    Next

                ElseIf Meldung.@Datatype.ToString Like "Array*" Then
                    AddressBit = 6
                    CountDBAdresse()
                    MeldungsBoolafter_others = True
                    ' MsgBox("Array noch nicht ausprogrammiert")
                    Dim GETDatatyp As String = Meldung.@Datatype

                    Dim ArrayBeginn As String = Meldung.@Datatype
                    ArrayBeginn = ArrayBeginn.Substring(6)
                    ArrayBeginn = ArrayBeginn.Remove(1)
                    'MsgBox(ArrayBeginn)

                    Dim ArrayEnde As String = Meldung.@Datatype

                    ArrayEnde = ArrayEnde.Substring(ArrayEnde.LastIndexOf("."))
                    ArrayEnde = ArrayEnde.Remove(ArrayEnde.IndexOf("]"))
                    ArrayEnde = ArrayEnde.Replace(".", "")
                    'MsgBox(ArrayEnde)



                    GETDatatyp = StrReverse(GETDatatyp)
                    GETDatatyp = GETDatatyp.Remove(GETDatatyp.LastIndexOf("fo"))
                    GETDatatyp = StrReverse(GETDatatyp)
                    GETDatatyp = GETDatatyp.Replace(" ", "")
                    'MsgBox(GETDatatyp)

                    Select Case GETDatatyp
                        Case "Byte"
                            For Arraynummer As Integer = CInt(ArrayBeginn) To CInt(ArrayEnde)
                                '  MsgBox(Meldung.Descendants(SiemensNamespace + "MultiLanguageText").Value)
                                For j = 0 To 7
                                    CountDBAdresse()
                                    Meldungen.Add(New HMIAlarms With {.AlarmText = Meldung.FirstAttribute.Value & ID,
                                                                            .Meldeklasse = Meldeklassenname, .Name = Meldung.FirstAttribute.Value & ID, .Datentyp = Meldung.@Datatype.ToString,
                                                                          .ID = ID, .TriggerTag = AddressTag, .TrigerBit = AddressBitforArray})
                                    ID = ID + 1

                                Next
                            Next
                        Case "Struct"
                            For Arraynummer As Integer = CInt(ArrayBeginn) To CInt(ArrayEnde)
                                AddressBit = 6
                                CountDBAdresse()
                                MeldungsBoolafter_others = True
                                MeldungStructName = Meldung.FirstAttribute.Value

                                Dim StructElement = (From element In Meldung.Nodes Select element)

                                For i As Integer = 0 To StructElement.Count - 1
                                    CountDBAdresse()
                                    Dim StructMeldung As XElement = StructElement.ElementAt(i)

                                    Meldungen.Add(New HMIAlarms With {.AlarmText = MeldungStructName & " " & Arraynummer & " " & StructMeldung.Descendants(SiemensNamespace + "MultiLanguageText").Value,
                                    .Meldeklasse = Meldeklassenname, .Name = Meldung.FirstAttribute.Value & ID, .Datentyp = StructMeldung.@Datatype.ToString,
                                    .ID = ID, .TriggerTag = AddressTag, .TrigerBit = AddressBitforArray})
                                    ID = ID + 1

                                    Meldungcounter = Meldungcounter + 1

                                Next
                            Next

                        Case Else ' Datentypen 
                            For Arraynummer As Integer = CInt(ArrayBeginn) To CInt(ArrayEnde)
                                AddressBit = 6
                                CountDBAdresse()
                                MeldungsBoolafter_others = True

                                Dim TypName As String

                                TypName = Meldung.FirstAttribute.Value



                                Const quote As String = """"
                                Dim TyponeHochkomma As String = GETDatatyp.Replace(quote, "")

                                Dim LO_Type = (From Element In Datentypen Where Element.Typname = TyponeHochkomma)


                                ' Console.WriteLine("Meldung" & TyponeHochkomma)
                                'Catch ex As Exception
                                '    MsgBox("Datenty nicht vorhanden")
                                'End Try
                                If LO_Type.Count = 0 Then
                                    MsgBox("Datentyp: " & Meldung.LastAttribute.Value & " nicht gefunden")
                                End If


                                For Each i As HMIAlarms In LO_Type
                                    CountDBAdresse()
                                    Meldungen.Add(New HMIAlarms With {.AlarmText = TypName & " " & Arraynummer & " " & i.AlarmText,
                                   .Meldeklasse = Meldeklassenname, .Name = i.Name & ID, .Datentyp = i.Datentyp,
                                   .ID = ID, .TriggerTag = AddressTag, .TrigerBit = AddressBitforArray})
                                    ID = ID + 1



                                Next
                            Next
                    End Select

                    '   Console.WriteLine(StructElement.Count)
                Else
                    MeldungsBoolafter_others = True
                    AddressBit = 6
                    CountDBAdresse()
                    Dim TypName As String

                    TypName = Meldung.FirstAttribute.Value



                    Const quote As String = """"
                    Dim TyponeHochkomma As String = Meldung.LastAttribute.Value.Replace(quote, "")

                    Dim LO_Type = (From Element In Datentypen Where Element.Typname = TyponeHochkomma)


                    ' Console.WriteLine("Meldung" & TyponeHochkomma)
                    'Catch ex As Exception
                    '    MsgBox("Datenty nicht vorhanden")
                    'End Try
                    If LO_Type.Count = 0 Then
                        MsgBox("Datentyp: " & Meldung.LastAttribute.Value & " nicht gefunden")
                    End If


                    For Each i As HMIAlarms In LO_Type
                        CountDBAdresse()
                        Meldungen.Add(New HMIAlarms With {.AlarmText = TypName & " " & i.AlarmText,
                       .Meldeklasse = Meldeklassenname, .Name = i.Name & ID, .Datentyp = i.Datentyp,
                       .ID = ID, .TriggerTag = AddressTag, .TrigerBit = AddressBitforArray})
                        ID = ID + 1



                    Next

                End If
            End If
        Next


    End Sub



    Private Sub GetDatatyp(ByVal Pfad As String)
        _StatusChanged("Datentypen auslesen")

        Dim XMLFile As XDocument
        XMLFile = XDocument.Load(Pfad)



        Dim SiemensNamespace As XNamespace = "http://www.siemens.com/automation/Openness/SW/Interface/v1" ' must match declaration In document



        Dim Name As XElement = (From el In XMLFile.<Document>.<SW.ControllerDatatype>.<AttributeList>
                                Select el).Last.LastNode




        'Console.WriteLine(Name.Value)

        Dim Interface_Sections As XElement =
            (From el In XMLFile.<Document>.<SW.ControllerDatatype>.<AttributeList>.<Interface>
             Select el).First



        'Dim zähler As Integer = 0
        'For Each el As XElement In Interface_Sections
        Dim SelectionsElemente As IEnumerable(Of XElement) =
        From element In Interface_Sections.Elements(SiemensNamespace + "Sections")
        Select element
        '    zähler = zähler + 1

        '  Dim idf As IEnumerable(Of XElement) = SelectionsElemente.DescendantNodes

        ' names.ElementAt(Random.Next(0, names.Length))
        Dim SectionElement As XElement = SelectionsElemente.ElementAt(0)



        Dim Selection = From element In SelectionsElemente.Elements(SiemensNamespace + "Section") Select element

        Dim Typklassen = (From element In Selection.Elements(SiemensNamespace + "Member") Select element)
        '   Console.WriteLine(SelectionElement.Descendants(SiemensNamespace + "MultiLanguageText").Skip(1).Take(20).Value)

        '  Dim LO_TypStörungen As New List(Of HMIAlarms)

        ' Dim TypName As String = Nothing
        For i As Integer = 0 To Typklassen.Count - 1

            Dim Typmeldungen As XElement = Typklassen.ElementAt(i)


            ' TypName = Typmeldungen.FirstAttribute.Value

            ' Console.WriteLine(Name.Value)
            Datentypen.Add(New HMIAlarms With {.AlarmText = Typmeldungen.Descendants(SiemensNamespace + "MultiLanguageText").Value,
                                 .Name = Typmeldungen.FirstAttribute.Value, .Datentyp = Typmeldungen.LastAttribute.Value, .Typname = Name.Value})



        Next




    End Sub


    Public Sub Write_Excel()

        System.IO.Directory.CreateDirectory(GetFolderPath(SpecialFolder.MyDocuments) & "\Meldegenerator_HMI_Alarms")


        '   excelApp.Run()
        ExcelDatenEinfügen()

        ExcelSpeichern(GetFolderPath(SpecialFolder.MyDocuments) & "\Meldegenerator_HMI_Alarms\HMIAlarms_" & CPUName & ".xlsx")
        ' Excel._Worksheet = (Excel.Worksheet)
        'Property ExcelFile As String
        '   Property ExcelBlatt As Byte


    End Sub

    Dim excelApp As Excel.Application = Nothing
    Dim wkbk As Excel.Workbook
    Dim sheet As Excel.Worksheet
    Sub CreateWorkbook()

        _StatusChanged("Create Excel File")

        ' Start Excel and create a workbook and worksheet.
        excelApp = New Excel.Application
        wkbk = excelApp.Workbooks.Add()
        sheet = CType(wkbk.Sheets.Add(), Excel.Worksheet)
        sheet.Name = "DiscreteAlarms"






    End Sub


    Sub ExcelDatenEinfügen()
        _StatusChanged("Daten In Excel File schreiben")


        'For i = 1 To values.Length - 1
        '    sheet.Cells(i, 1) = values(i)
        'Next
        Dim i As Integer = 1
        sheet.Cells(1, 1) = Column0_A
        sheet.Cells(1, 2) = Column1_B

        sheet.Cells(1, 3) = Column2_C
        sheet.Cells(1, 4) = Column3_D
        sheet.Cells(1, 5) = Column4_E
        sheet.Cells(1, 6) = Column5_F
        sheet.Cells(1, 7) = Column6_G
        sheet.Cells(1, 8) = Column7_H
        sheet.Cells(1, 9) = Column8_I
        sheet.Cells(1, 10) = Column9_J
        sheet.Cells(1, 11) = Column10_K
        sheet.Cells(1, 12) = Column11_L
        sheet.Cells(1, 13) = Column12_M
        sheet.Cells(1, 14) = Column13_N

        i = i + 1

        For Each Alarm As HMIAlarms In Meldungen
            sheet.Cells(i, 1) = Alarm.ID
            sheet.Cells(i, 2) = Alarm.Name
            sheet.Cells(i, 3) = Alarm.AlarmText
            sheet.Cells(i, 4) = Alarm.FieldInfo
            sheet.Cells(i, 5) = Alarm.Meldeklasse
            sheet.Cells(i, 6) = Alarm.TriggerTag
            sheet.Cells(i, 7) = Alarm.TrigerBit
            sheet.Cells(i, 8) = Alarm.Acknowledgementtag
            sheet.Cells(i, 9) = Alarm.Acknoledgementbit
            sheet.Cells(i, 10) = Alarm.PLCAcknowledgementTag
            sheet.Cells(i, 11) = Alarm.PLCAcknowledgementBit
            sheet.Cells(i, 12) = Alarm.Group
            sheet.Cells(i, 13) = Alarm.Report
            sheet.Cells(i, 14) = Alarm.InfoText

            i = i + 1
        Next
        Meldungen.Clear()


    End Sub


    Sub ExcelSpeichern(ByVal filePath As String)
        _StatusChanged("Excel Speicher, abschliessen")
        Try


            excelApp.DisplayAlerts = False
            Dim folderPath = My.Computer.FileSystem.GetParentPath(filePath)
            If Not My.Computer.FileSystem.DirectoryExists(folderPath) Then
                My.Computer.FileSystem.CreateDirectory(folderPath)
            End If
            wkbk.SaveAs(filePath)

            excelApp.DisplayAlerts = False
            'Dim folderPath = My.Computer.FileSystem.GetParentPath(filePath)
            'If Not My.Computer.FileSystem.DirectoryExists(folderPath) Then
            '    My.Computer.FileSystem.CreateDirectory(folderPath)
            'End If
            wkbk.SaveAs(filePath)


            wkbk.Close()
            sheet = Nothing
            wkbk = Nothing

            ' Close Excel.


            excelApp.Quit()
            excelApp = Nothing

            _StatusChanged("Fertig")
        Catch ex As System.Runtime.InteropServices.COMException
            _StatusChanged("Fehler beim EXCEL speichern")
            MsgBox("Das Excel File konnte nicht gespeichert werden, es ist vielleicht geöffnet.")
        End Try

        '  MsgBox("Fertig")
    End Sub

End Class



Public Class HMIAlarms
    Public ID As Integer
    Public Name As String
    Public AlarmText As String
    Public FieldInfo As String
    Public Meldeklasse As String
    Public TriggerTag As String
    Public TrigerBit As Integer
    Public Acknowledgementtag As String = "<No value>"
    Public Acknoledgementbit As Integer = 0
    Public PLCAcknowledgementTag As String = "<No value>"
    Public PLCAcknowledgementBit As Integer = 0
    Public Group As String = "<No value>"
    Public Report As String = "'False"
    Public InfoText As String = "<No value>"
    Public Datentyp As String
    Public Typname As String




End Class