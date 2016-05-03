Imports System.IO
Imports System.Environment

Imports Excel = Microsoft.Office.Interop.Excel






Public Class XML
    Public Event StatusChaged()
    Private Sub _StatusChanged(ByVal _Status As String)
        Status = _Status
        RaiseEvent StatusChaged()
    End Sub

    Property CPUnummer As Integer = 1
    Property DBNummer As Integer = 260

    Property Status As String


    Dim Meldungen As New List(Of HMIAlarms)
    Dim Störungen As New List(Of HMIAlarms)

    Dim Datentypen As New List(Of HMIAlarms)

    Dim TagName As String = "Trigger_AT_" & CPUnummer.ToString & "_DB"

    Friend XMLFile As XDocument
    Public Sub RunXML()

        _StatusChanged("Load XML")

        For Each file As String In Directory.GetFiles(GetFolderPath(SpecialFolder.MyDocuments) & "\Meldegenerator_XML\Datentypen")
            GetDatatyp(file)
        Next
        Try
            XMLFile = XDocument.Load(GetFolderPath(SpecialFolder.MyDocuments) & "\Meldegenerator_XML\Meldungen.xml")
        Catch ex As Exception
            _StatusChanged("Fehler beim XML öffnen")
        End Try
        GetHMIMeldungen()



        Write_Excel()



    End Sub


    Private AddresssWord As Integer = -2
    Private AddressBit As Integer = 7
    Private AddressTag As String

    Private Sub CountDBAdresse()

        _StatusChanged("Calculate DB Address")

        If AddressBit >= 15 Then
            AddressBit = 0

        Else
            AddressBit = AddressBit + 1
        End If

        If AddressBit = 8 Then
            AddresssWord = AddresssWord + 2
        End If

        TagName = "Trigger_AT_" & CPUnummer.ToString & "_DB"

        AddressTag = """" & TagName & DBNummer & ".DBW" & AddresssWord & """"
    End Sub


    Dim ID As Integer = CPUnummer * 10000
    Dim SiemensNamespace As XNamespace = "http://www.siemens.com/automation/Openness/SW/Interface/v1" ' must match declaration In document
    Private Sub GetHMIMeldungen()
        _StatusChanged("XML initialisieren")

        Dim Interface_Sections As XElement =
            (From el In XMLFile.<Document>.<SW.DataBlock>.<AttributeList>.<Interface>
             Select el).First
        CountDBAdresse()

        Dim SelectionsElemente As IEnumerable(Of XElement) =
        From element In Interface_Sections.Elements(SiemensNamespace + "Sections")
        Select element

        Dim SectionElement As XElement = SelectionsElemente.ElementAt(0)


        Dim Selection = From element In SelectionsElemente.Elements(SiemensNamespace + "Section") Select element



        Dim Meldeklassen = (From element In Selection.Elements(SiemensNamespace + "Member") Select element)

        ID = CPUnummer * 10000

        Dim Meldeklasse = (From element In Meldeklassen Where element.FirstAttribute.Value Like "M_*")

        MeldungenGenerieren(Meldeklasse)


        Dim Störklasse = (From element In Meldeklassen Where element.FirstAttribute.Value Like "S_*")
        StörungenGenerieren(Störklasse)



    End Sub



    Private Sub MeldungenGenerieren(ByVal Meldeklasse As IEnumerable(Of XElement))
        Dim MeldungsBoolafter_others As Boolean = False
        Dim MeldungAlarmtext As String = Nothing
        Dim MeldungStructName As String = Nothing
        Dim Meldungcounter As Integer = 0
        _StatusChanged("Meldungen aus XML lesen")
        For Each Meldung As XElement In Meldeklasse.Elements



            If Meldung.Name = "{" & SiemensNamespace.ToString & "}Member" Then

                If Meldung.@Datatype.ToString = "Bool" Then
                    If MeldungsBoolafter_others = True Then
                        AddressBit = 7
                        CountDBAdresse()
                        MeldungsBoolafter_others = False
                    End If
                    Meldungen.Add(New HMIAlarms With {.AlarmText = Meldung.Descendants(SiemensNamespace + "MultiLanguageText").Value,
                                      .Meldeklasse = "Meldungen", .Name = Meldung.FirstAttribute.Value, .Datentyp = Meldung.@Datatype.ToString,
                                      .ID = ID, .TriggerTag = AddressTag, .TrigerBit = AddressBit})
                    ID = ID + 1
                    CountDBAdresse()
                ElseIf Meldung.@Datatype.ToString = "Struct" Then
                    AddressBit = 7
                    CountDBAdresse()

                    MeldungsBoolafter_others = True
                    MeldungStructName = Meldung.FirstAttribute.Value

                    Dim StructElement = (From element In Meldung.Nodes Select element)

                    For i As Integer = 1 To StructElement.Count - 1
                        Dim StructMeldung As XElement = StructElement.ElementAt(i)

                        Meldung.Add(New HMIAlarms With {.AlarmText = MeldungStructName & " " & StructMeldung.Descendants(SiemensNamespace + "MultiLanguageText").Value,
                        .Meldeklasse = "Meldung", .Name = StructMeldung.FirstAttribute.Value, .Datentyp = StructMeldung.FirstAttribute.Value,
                        .ID = ID, .TriggerTag = AddressTag, .TrigerBit = AddressBit})
                        ID = ID + 1
                        CountDBAdresse()
                        Meldungcounter = Meldungcounter + 1

                    Next

                ElseIf Meldung.@Datatype.ToString Like "Array*" Then
                    AddressBit = 7
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
                            For i As Integer = CInt(ArrayBeginn) To CInt(ArrayEnde)
                                For j = 0 To 7
                                    Meldungen.Add(New HMIAlarms With {.AlarmText = Meldung.FirstAttribute.Value,
                                                                          .Meldeklasse = "Meldungen", .Name = Meldung.FirstAttribute.Value & ID, .Datentyp = Meldung.@Datatype.ToString,
                                                                          .ID = ID, .TriggerTag = AddressTag, .TrigerBit = AddressBit})
                                    ID = ID + 1
                                    CountDBAdresse()
                                Next
                            Next
                    End Select
                    AddressBit = 7
                    CountDBAdresse()
                    '   Console.WriteLine(StructElement.Count)
                Else
                    MeldungsBoolafter_others = True
                    AddressBit = 7
                    CountDBAdresse()
                    Dim TypName As String

                    TypName = Meldung.FirstAttribute.Value


                    Dim Test = (From Element In Datentypen Select Element)
                    Const quote As String = """"
                    Dim TyponeHochkomma As String = Meldung.LastAttribute.Value.Replace(quote, "")

                    Dim LO_Type = (From Element In Datentypen Where Element.Typname = TyponeHochkomma Select Element)
                    For Each i As HMIAlarms In Test
                        Console.WriteLine(i.Typname)
                        'Console.WriteLine(i.Typname.ToString)
                    Next

                    Console.WriteLine("Meldung" & TyponeHochkomma)
                    'Catch ex As Exception
                    '    MsgBox("Datenty nicht vorhanden")
                    'End Try
                    If LO_Type.Count = 0 Then
                        MsgBox("Datentyp: " & Meldung.LastAttribute.Value & " nicht gefunden")
                    End If


                    For Each i As HMIAlarms In LO_Type

                        Meldung.Add(New HMIAlarms With {.AlarmText = TypName & " " & i.AlarmText,
                       .Meldeklasse = "Meldung", .Name = i.Name, .Datentyp = i.Datentyp,
                       .ID = ID, .TriggerTag = AddressTag, .TrigerBit = AddressBit})
                        ID = ID + 1
                        CountDBAdresse()

                    Next


                End If

            End If

        Next

    End Sub


    Private Sub StörungenGenerieren(ByVal Störklasse As IEnumerable(Of XElement))
        Dim StörungsBoolafter_others As Boolean = False
        AddressBit = 7
        CountDBAdresse()
        _StatusChanged("Störungen aus XML lesen")
        For Each Störung As XElement In Störklasse.Elements


            If Störung.Name = "{" & SiemensNamespace.ToString & "}Member" Then


                Dim StörungAlarmtext As String = Nothing
                Dim StörungStructName As String = Nothing
                Dim Störungcounter As Integer = 0

                If Störung.@Datatype.ToString = "Bool" Then
                    If StörungsBoolafter_others = True Then
                        AddressBit = 7
                        CountDBAdresse()
                        StörungsBoolafter_others = False
                    End If
                    StörungStructName = ""


                    Störungen.Add(New HMIAlarms With {.AlarmText = StörungStructName & Störung.Descendants(SiemensNamespace + "MultiLanguageText").Value,
                             .Meldeklasse = "Störung", .Name = Störung.FirstAttribute.Value, .Datentyp = Störung.@Datatype.ToString,
                             .ID = ID, .TriggerTag = AddressTag, .TrigerBit = AddressBit})
                    ID = ID + 1
                    CountDBAdresse()


                ElseIf Störung.@Datatype.ToString = "Struct" Then
                    StörungsBoolafter_others = True
                    AddressBit = 7
                    CountDBAdresse()
                    StörungStructName = Störung.FirstAttribute.Value

                    Dim StructElement = (From element In Störung.Nodes Select element)

                    For i As Integer = 1 To StructElement.Count - 1
                        Dim StructStörung As XElement = StructElement.ElementAt(i)

                        Störungen.Add(New HMIAlarms With {.AlarmText = StörungStructName & " " & StructStörung.Descendants(SiemensNamespace + "MultiLanguageText").Value,
                        .Meldeklasse = "Störung", .Name = StructStörung.FirstAttribute.Value, .Datentyp = StructStörung.FirstAttribute.Value,
                        .ID = ID, .TriggerTag = AddressTag, .TrigerBit = AddressBit})
                        ID = ID + 1
                        CountDBAdresse()
                        Störungcounter = Störungcounter + 1
                    Next


                ElseIf Störung.@Datatype.ToString Like "Array*" Then
                    StörungsBoolafter_others = True
                    AddressBit = 7
                    CountDBAdresse()
                    '   MsgBox("Array noch nicht ausprogrammiert")
                    Dim GETDatatyp As String = Störung.@Datatype

                    Dim ArrayBeginn As String = Störung.@Datatype
                    ArrayBeginn = ArrayBeginn.Substring(6)
                    ArrayBeginn = ArrayBeginn.Remove(1)
                    '  MsgBox(ArrayBeginn)

                    Dim ArrayEnde As String = Störung.@Datatype

                    ArrayEnde = ArrayEnde.Substring(ArrayEnde.LastIndexOf("."))
                    ArrayEnde = ArrayEnde.Remove(ArrayEnde.IndexOf("]"))
                    ArrayEnde = ArrayEnde.Replace(".", "")
                    ' MsgBox(ArrayEnde)



                    GETDatatyp = StrReverse(GETDatatyp)
                    GETDatatyp = GETDatatyp.Remove(GETDatatyp.LastIndexOf("fo"))
                    GETDatatyp = StrReverse(GETDatatyp)
                    GETDatatyp = GETDatatyp.Replace(" ", "")
                    '  MsgBox(GETDatatyp)

                    Select Case GETDatatyp
                        Case "Byte"
                            For i As Integer = CInt(ArrayBeginn) To CInt(ArrayEnde)
                                For j = 0 To 7
                                    Meldungen.Add(New HMIAlarms With {.AlarmText = Störung.FirstAttribute.Value,
                                                                          .Meldeklasse = "Störung", .Name = Störung.FirstAttribute.Value & ID, .Datentyp = Störung.@Datatype.ToString,
                                                                          .ID = ID, .TriggerTag = AddressTag, .TrigerBit = AddressBit})
                                    ID = ID + 1
                                    CountDBAdresse()
                                Next
                            Next

                    End Select
                    '   Console.WriteLine(StructElement.Count)
                Else
                    StörungsBoolafter_others = True
                    AddressBit = 7
                    CountDBAdresse()
                    Dim TypName As String

                    TypName = Störung.FirstAttribute.Value


                    Dim Test = (From Element In Datentypen Select Element)
                    Const quote As String = """"
                    Dim TyponeHochkomma As String = Störung.LastAttribute.Value.Replace(quote, "")

                    Dim LO_Type = (From Element In Datentypen Where Element.Typname = TyponeHochkomma Select Element)
                    For Each i As HMIAlarms In Test
                        '  Console.WriteLine(i.Typname)
                        'Console.WriteLine(i.Typname.ToString)
                    Next

                    ' Console.WriteLine("Störung" & TyponeHochkomma)
                    'Catch ex As Exception
                    '    MsgBox("Datenty nicht vorhanden")
                    'End Try
                    If LO_Type.Count = 0 Then
                        MsgBox("Datentyp: " & Störung.LastAttribute.Value & " nicht gefunden")
                    End If


                    For Each i As HMIAlarms In LO_Type

                        Störungen.Add(New HMIAlarms With {.AlarmText = TypName & " " & i.AlarmText,
                       .Meldeklasse = "Störung", .Name = i.Name, .Datentyp = i.Datentyp,
                       .ID = ID, .TriggerTag = AddressTag, .TrigerBit = AddressBit})
                        ID = ID + 1
                        CountDBAdresse()

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

        Dim TypName As String = Nothing
        For i As Integer = 0 To Typklassen.Count - 1

            Dim Typmeldungen As XElement = Typklassen.ElementAt(i)

            If i = 0 Then
                TypName = Typmeldungen.FirstAttribute.Value
            Else
                ' Console.WriteLine(Name.Value)
                Datentypen.Add(New HMIAlarms With {.AlarmText = TypName & Typmeldungen.Descendants(SiemensNamespace + "MultiLanguageText").Value,
                                .Meldeklasse = "Störung", .Name = Typmeldungen.FirstAttribute.Value, .Datentyp = Typmeldungen.LastAttribute.Value, .Typname = Name.Value})
            End If

            ' Console.WriteLine(Meldeklassen.Elements)

            '   If TypMeldung.@Datatype.ToString = "Struct" Then

            '  TypName = TypMeldung.FirstAttribute.ToString


        Next




    End Sub


    Public Sub Write_Excel()

        System.IO.Directory.CreateDirectory(GetFolderPath(SpecialFolder.MyDocuments) & "\Meldegenerator_HMI_Alarms")

        CreateWorkbook()
        '   excelApp.Run()
        ExcelDatenEinfügen()

        ExcelSpeichern(GetFolderPath(SpecialFolder.MyDocuments) & "\Meldegenerator_HMI_Alarms\HMIAlarms.xlsx")
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
        sheet.Cells(1, 1) = Meldungen(0).Column0_A
        sheet.Cells(1, 2) = Meldungen(0).Column1_B

        sheet.Cells(1, 3) = Meldungen(0).Column2_C
        sheet.Cells(1, 4) = Meldungen(0).Column3_D
        sheet.Cells(1, 5) = Meldungen(0).Column4_E
        sheet.Cells(1, 6) = Meldungen(0).Column5_F
        sheet.Cells(1, 7) = Meldungen(0).Column6_G
        sheet.Cells(1, 8) = Meldungen(0).Column7_H
        sheet.Cells(1, 9) = Meldungen(0).Column8_I
        sheet.Cells(1, 10) = Meldungen(0).Column9_J
        sheet.Cells(1, 11) = Meldungen(0).Column10_K
        sheet.Cells(1, 12) = Meldungen(0).Column11_L
        sheet.Cells(1, 13) = Meldungen(0).Column12_M
        sheet.Cells(1, 14) = Meldungen(0).Column13_N

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
        For Each Alarm As HMIAlarms In Störungen
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
    Public Report As String = "False"
    Public InfoText As String = "<No value>"
    Public Datentyp As String
    Public Typname As String

    Public ReadOnly Column0_A As String = "ID"
    Public ReadOnly Column1_B As String = "Name"
    Public ReadOnly Column2_C As String = "Event text [de-DE], Alarm text"
    Public ReadOnly Column3_D As String = "FieldInfo [Alarm text]"
    Public ReadOnly Column4_E As String = "Class"
    Public ReadOnly Column5_F As String = "Trigger tag"
    Public ReadOnly Column6_G As String = "Trigger bit"
    Public ReadOnly Column7_H As String = "Acknowledgement tag"
    Public ReadOnly Column8_I As String = "Acknowledgement bit"
    Public ReadOnly Column9_J As String = "PLC acknowledgement tag"
    Public ReadOnly Column10_K As String = "PLC acknowledgement bit"
    Public ReadOnly Column11_L As String = "Group"
    Public ReadOnly Column12_M As String = "Report"
    Public ReadOnly Column13_N As String = "Info text [de-DE], Info text"



End Class