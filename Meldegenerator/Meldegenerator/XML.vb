Imports System.IO
Imports System.Environment

Imports Excel = Microsoft.Office.Interop.Excel






Public Class XML

    Property CPUnummer As Integer = 1
    Property DBNummer As Integer = 260



    Dim Meldungen As New List(Of HMIAlarms)
    Dim Störungen As New List(Of HMIAlarms)

    Dim Datentypen As New List(Of HMIAlarms)

    Dim TagName As String = "Trigger_AT_" & CPUnummer.ToString & "_DB"

    Friend XMLFile As XDocument
    Public Sub LoadXML()


        For Each file As String In Directory.GetFiles(GetFolderPath(SpecialFolder.MyDocuments) & "\Meldegenerator_XML\Datentypen")
            GetDatatyp(file)
        Next

        XMLFile = XDocument.Load("C:\Users\m.baminger\Documents\Meldegenerator_XML\Meldungen.xml")

        GetHMIMeldungen()

        For Each f As HMIAlarms In Datentypen
            Console.WriteLine(f.AlarmText)
        Next

        '   Next




    End Sub


    Private Sub CountDBAdresse()

    End Sub


    Private Sub GetHMIMeldungen()
        Dim SiemensNamespace As XNamespace = "http://www.siemens.com/automation/Openness/SW/Interface/v1" ' must match declaration In document



        Dim Interface_Sections As XElement =
            (From el In XMLFile.<Document>.<SW.DataBlock>.<AttributeList>.<Interface>
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


        'If (SectionElement.Name.Namespace Is XNamespace.None) Then
        '    Console.WriteLine("The element el2 is in no namespace.")
        'Else
        '    Console.WriteLine("The element el2 is in a namespace.")
        '    'a = el2.Name.NamespaceName
        '    '  Dim Node1 As XNode = el2.FirstNode

        'End If


        ' Console.WriteLine(SelectionsElemente.)

        '   Dim AlleMeldungen As IEnumerable(Of XElement) = (From element In SelectionElement.Descendants(SiemensNamespace + "Member").Attributes Where element.Value = "M_").First

        '   Dim LO_Meldungen As List(Of XElement) = SelectionElement.Descendants(SiemensNamespace + "MultiLanguageText").ToList

        '  Dim LO_Meldungen = (From element In SelectionsElemente(SiemensNamespace + "Member") Select element.@Name = "_M")
        Dim Selection = From element In SelectionsElemente.Elements(SiemensNamespace + "Section") Select element

        Dim Meldeklassen = (From element In Selection.Elements(SiemensNamespace + "Member") Select element)
        '   Console.WriteLine(SelectionElement.Descendants(SiemensNamespace + "MultiLanguageText").Skip(1).Take(20).Value)




        Dim ID As Integer = 10000
        Dim Meldeklasse = (From element In Meldeklassen Where element.FirstAttribute.Value = "M_")



        For Each Meldung As XElement In Meldeklasse.Elements
            ' Console.WriteLine(Meldeklassen.Elements)

            '  Console.WriteLine(Meldung.Name)
            '  If Meldung.Name Then
            '   If Meldung.Parent.FirstAttribute = "M_" Then

            If Meldung.Name = "{" & SiemensNamespace.ToString & "}Member" Then


                Meldungen.Add(New HMIAlarms With {.AlarmText = Meldung.Descendants(SiemensNamespace + "MultiLanguageText").Value,
                                      .Meldeklasse = "Meldungen", .Name = Meldung.FirstAttribute.Value, .Datentyp = Meldung.@Datatype.ToString,
                                      .ID = ID, .TriggerTag = })
                ID = ID + 1

            End If

        Next

        Dim Störklasse = (From element In Meldeklassen Where element.FirstAttribute.Value = "S_")

        ' Console.WriteLine(Störklasse.Count)


        '  Dim StructElement As IEnumerable(Of XElement) = (From element In Störklasse.Elements() Select element Where element.FirstAttribute = StructName)


        For Each Störung As XElement In Störklasse.Elements
            ' Console.WriteLine(Meldeklassen.Elements)

            '  Console.WriteLine(Störung.Name)
            '  If Meldung.Name Then
            '   If Meldung.Parent.FirstAttribute = "M_" Then

            If Störung.Name = "{" & SiemensNamespace.ToString & "}Member" Then


                Dim Alarmtext As String = Nothing
                Dim StructName As String = Nothing
                Dim counter As Integer = 0

                If Störung.@Datatype.ToString = "Bool" Then
                    StructName = ""


                    Störungen.Add(New HMIAlarms With {.AlarmText = StructName & Störung.Descendants(SiemensNamespace + "MultiLanguageText").Value,
                             .Meldeklasse = "Störung", .Name = Störung.FirstAttribute.ToString, .Datentyp = Störung.@Datatype.ToString,
                             .ID = ID})
                    ID = ID + 1


                ElseIf Störung.@Datatype.ToString = "Struct" Then

                    StructName = Störung.FirstAttribute.ToString

                    Dim StructElement = (From element In Störung.Nodes Select element)

                    For i As Integer = 1 To StructElement.Count - 1
                        Dim StructStörung As XElement = StructElement.ElementAt(i)

                        Störungen.Add(New HMIAlarms With {.AlarmText = StructName & " " & StructStörung.Descendants(SiemensNamespace + "MultiLanguageText").Value,
                        .Meldeklasse = "Störung", .Name = StructStörung.FirstAttribute.Value, .Datentyp = StructStörung.FirstAttribute.Value,
                        .ID = ID})
                        ID = ID + 1
                        counter = counter + 1
                    Next



                    '   Console.WriteLine(StructElement.Count)
                Else
                    Dim TypName As String

                    TypName = Störung.FirstAttribute.Value
                    '   Dim DatetypName As String = Störung.LastAttribute.ToString

                    '   Try
                    '   Console.WriteLine(Datentypen)

                    Dim LO_Type = (From Element In Datentypen Where Element.Typname = Störung.LastAttribute.Value Select Element)
                    'Catch ex As Exception
                    '    MsgBox("Datenty nicht vorhanden")
                    'End Try
                    If Not LO_Type.Count = 0 Then
                        MsgBox("Datentyp: " & Störung.LastAttribute.Value & " nicht gefunden")
                    End If


                    For Each i As HMIAlarms In LO_Type

                        Störungen.Add(New HMIAlarms With {.AlarmText = TypName & " " & i.AlarmText,
                       .Meldeklasse = "Störung", .Name = i.Name, .Datentyp = i.Datentyp,
                       .ID = ID})
                        ID = ID + 1


                    Next


                End If

                'Dim MKAttributes As IEnumerable(Of XAttribute) = Meldung.Attributes.ToList




            End If


        Next

        For Each i As HMIAlarms In Störungen
            Console.WriteLine(i.AlarmText)
        Next


        Console.WriteLine()
    End Sub

    Private Sub GetDatatyp(ByVal Pfad As String)
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

        '   HMIDatentypen.Add(New Datentypen With {.TypenName = TypName})

        ' HMIDatentypen.Last.TypStörungen.Add(LO_TypStörungen.First)


        Console.WriteLine()



        ' schliessen
        '  XMLFile.Save(Pfad, SaveOptions.None)



    End Sub


    Public Sub Write_Excel()
        Dim Fred() As Integer = {1, 3, 2, 3, 4, 5, 8}

        CreateWorkbook()
        '   excelApp.Run()
        ExcelDatenEinfügen()

        ExcelSpeichern("D:\TestExcel_1.xlsx")
        ' Excel._Worksheet = (Excel.Worksheet)
        'Property ExcelFile As String
        '   Property ExcelBlatt As Byte


    End Sub

    Dim excelApp As Excel.Application = Nothing
    Dim wkbk As Excel.Workbook
    Dim sheet As Excel.Worksheet
    Sub CreateWorkbook()



        ' Start Excel and create a workbook and worksheet.
        excelApp = New Excel.Application
        wkbk = excelApp.Workbooks.Add()
        sheet = CType(wkbk.Sheets.Add(), Excel.Worksheet)
        sheet.Name = "DiscreteAlarms"

        ' Write a column of values.
        ' In the For loop, both the row index and array index start at 1.
        ' Therefore the value of 4 at array index 0 is not included.

        ' Suppress any alerts and save the file. Create the directory 
        ' if it does not exist. Overwrite the file if it exists.




    End Sub


    Sub ExcelDatenEinfügen()
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
    Public Acknowledgementtag As String = " < No value>"
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