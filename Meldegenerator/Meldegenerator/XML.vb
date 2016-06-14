Imports System.IO
Imports System.Environment

Imports Excel = Microsoft.Office.Interop.Excel


''' <summary>
''' Öffnent das exportierte XML file liest die benötigten Tags aus, und erstellt ein Importierbares Excel
''' Erteller: Manfred Baminger
''' Benötigte Parameter
''' 
''' </summary>
''' 
Public Class XML




    Public Event StatusChaged()
    Private Sub _StatusChanged(ByVal _Status As String)
        Status = _Status
        RaiseEvent StatusChaged()
    End Sub


    'MAB:  benötigte Values
    Property CPUnummer As Integer = 1
    Property DBNummer As Integer = 260
    Property CPUName As String = ""

    'MAB:  Rückgabewerte
    Property Status As String
    Property HMIVariableDatentyp As String = ""
    Property HMIVariablenName As String = ""

    Property inWords As Boolean



    ''' <summary>
    ''' MBA:
    ''' in disem Container werden alle Störungen und Meldungen zusammengefasst
    ''' </summary>
    ''' <remarks><seealso cref="HMIAlarms"/></remarks>
    Private Meldungen As New List(Of HMIAlarms)
    ''' <summary>
    ''' Ein Container für alle Datentypen
    ''' </summary>
    Private Datentypen As New List(Of HMIAlarms)


    'MAB:  Variablen declaration class global
    Private TagName As String
    Private AddressWord As Integer = -1
    Private AddressBit As Integer = 7
    Private AddressBitforArray As Integer
    Private AddressTag As String

    Dim TempADTAG As String
    ''' <summary>
    ''' MBA:
    ''' Runs XML read and Excel erstellen
    ''' Es müssen die Propertys :
    ''' <remarks><seealso cref="CPUnummer"/>
    ''' <seealso cref="DBNummer"/>
    ''' <seealso cref="_CPUName"/>
    ''' </remarks>
    ''' mit Werten versorgt sein.
    ''' </summary>
    Public Sub RunXML()
        'MBA:  Neue Startwerte zuweisen (nötig bei mehreren CPUs)
        AddressWord = -1
        AddressBit = 7
        ID = CPUnummer * 10000


        'TagName bilden (Der HMI Variablen Name)
        TagName = "Trigger_AT_" & CPUName & "_DB"
        AddressTag = """" & TagName & DBNummer & """"
        TempADTAG = AddressTag
        'MAB:  Proof Directory file name and open XML file
        Dim XMLFile As XDocument
        _StatusChanged("Load XML")


        'Ablöschen bei neuer CPU, damit bei öffterem Durchlauf keine Doppelten Werte entstehen
        Meldungen.Clear()
        Datentypen.Clear()

        'Datentypen auswerten und zur weiteren Verwendung in Container legen
        For Each file As String In Directory.GetFiles(GetFolderPath(SpecialFolder.MyDocuments) & "\Meldegenerator_XML\Datentypen")
            GetDatatyp(file)
        Next

        Try
            XMLFile = XDocument.Load(GetFolderPath(SpecialFolder.MyDocuments) & "\Meldegenerator_XML\Meldungen.xml")


            'MAB:  Gibt das XElement Intervace_Sections zurück
            'MAB:  (Alle Nodes vor dem Namespace, Alle folgenden Nodes müssen mit dem Namespace angesprochenwerden.)
            Dim Interface_Sections As XElement =
            (From el In XMLFile.<Document>.<SW.DataBlock>.<AttributeList>.<Interface>
             Select el).First

            'MAB:  ruft die Eigentliche XML bearbeitung auf
            GetHMIMeldungen(Interface_Sections)

        Catch ex As Exception
            _StatusChanged("Fehler beim XML öffnen")

        End Try



        'MAB:  rückgabewerte Adresse und Tagname 
        If AddressWord > 0 Then
            HMIVariableDatentyp = "Array [0.." & AddressWord & "] Of Word"
            HMIVariablenName = AddressTag.Replace("""", "")
        Else
            HMIVariableDatentyp = "Array [0..1] Of Word"
            HMIVariablenName = AddressTag.Replace("""", "")
        End If

        'MAB:  erstellt eine Excel mappe und  Areitsblatt 
        CreateWorkbook()

        'MAB:  Schreibt die Meldungen in das Excel  File
        Write_Excel()



    End Sub



    ''' <summary>
    ''' MBA:
    ''' Mit jedem Durchlauf wird das AddressBit um eins erhöht 
    ''' und beim 8.bit wird das Address Word um eins erhöht
    ''' somit kann ich die Tatsächliche Array Pos errechnen
    ''' 
    ''' Soll ein ein Adresswort hochgezählt werden wird das Adress bi manuell auf den Wert 6 geschrieben
    '''    
    ''' </summary>
    Private Sub CountDBAdresse()

        _StatusChanged("Calculate DB Address")



        If AddressBit >= 15 Then
            AddressBit = 0

        Else
            AddressBit = AddressBit + 1
        End If

        If AddressBit = 8 Then
            AddressWord = AddressWord + 1

        End If

        If inWords = False Then
            AddressBitforArray = AddressBit + ((AddressWord) * 16)

        Else
            AddressBitforArray = AddressBit
            AddressTag = TempADTAG & "[" & AddressWord & "]"
        End If




    End Sub

    'MAB:  Die start ID für die Jeweilige CPU
    Dim ID As Integer = CPUnummer * 10000
    'MAB:  der namespace der für alle untergeordneten Nodes nötig ist
    Dim SiemensNamespace As XNamespace = "http://www.siemens.com/automation/Openness/SW/Interface/v1" ' must match declaration In document

    ''' <summary>
    ''' MBA:
    ''' Abarbeitung XML
    ''' </summary>
    ''' <param name="Interface_Sections">Der Node mit dem Namen InterfaceSections</param>
    Private Sub GetHMIMeldungen(ByVal Interface_Sections As XElement)
        _StatusChanged("XML initialisieren")




        Dim SelectionsElemente As IEnumerable(Of XElement) =
        From element In Interface_Sections.Elements(SiemensNamespace + "Sections")
        Select element

        Dim SectionElement As XElement = SelectionsElemente.ElementAt(0)


        'MBA:  im Node Section Liegen die "Meldeklassen" jeder untergeordnete Node repräsentiert eine MEldeklasse
        Dim Section = From element In SelectionsElemente.Elements(SiemensNamespace + "Section") Select element


        'MBA:  Nur die <Member> Elemente sind interressant hier werden nur die Elemente des Typs "Member" in einer Meldeklasse (Node) ausgewählt. 
        Dim Meldeklassen = (From element In Section.Elements(SiemensNamespace + "Member") Select element)

        ' für alle Meldeklassen Die Meldungen generieren
        For i As Integer = 0 To Meldeklassen.Count - 1

            Dim AktuelleMeldeklasse = Meldeklassen.ElementAt(i).Elements
            MeldungenGenerieren(AktuelleMeldeklasse)
            AddressBit = 6
            CountDBAdresse()
        Next

    End Sub


    'TODO: evt. Struct und Datentyp in ein eigenenes Sub (is bei array und normal immer das gleiche)
    Private Sub MeldungenGenerieren(ByVal Meldeklasse As IEnumerable(Of XElement))
        'MBA:  Dieses Bit wird benötigt um nach einer Structur oder datentyps, eine Boolsche Variable ins nächste Word zu heben.
        Dim MeldungsBoolafter_others As Boolean = False

        Dim MeldungAlarmtext As String = Nothing
        Dim MeldungStructName As String = Nothing
        Dim Meldungcounter As Integer = 0
        _StatusChanged("Meldungen aus XML lesen")


        'holt den Klassennamen aus dem ersten element des XML
        Dim Meldeklassenname As String = Meldeklasse.First.Value

        For Each Meldung As XElement In Meldeklasse


            'Nurr untergeordnete "Member" Elemete aud den gewählten "Meber" elementen abarbeiten
            If Meldung.Name = "{" & SiemensNamespace.ToString & "}Member" Then

                'Wenn der Member vom Datentyp Bool ist dann:
                ' Boolsche Meldungen generieren
                If Meldung.@Datatype.ToString = "Bool" Then

                    'Meldeword hochzählen wenn das bool nach einem Struct oder Datentyp oder array folgt
                    If MeldungsBoolafter_others = True Then
                        AddressBit = 6
                        CountDBAdresse()
                        MeldungsBoolafter_others = False
                    End If
                    CountDBAdresse()
                    'die Werte am ende eines "Containers" einfügen
                    Meldungen.Add(New HMIAlarms With {.AlarmText = Meldung.Descendants(SiemensNamespace + "MultiLanguageText").Value,
                                      .Meldeklasse = Meldeklassenname, .Name = Meldung.FirstAttribute.Value & ID, .Datentyp = Meldung.@Datatype.ToString,
                                      .ID = ID, .TriggerTag = AddressTag, .TrigerBit = AddressBitforArray})

                    ID = ID + 1
                    'Wenn der Member vom Datentyp Struct ist dann:
                    'alle Meldungen im Struct lesen generieren
                ElseIf Meldung.@Datatype.ToString = "Struct" Then
                    AddressBit = 6
                    CountDBAdresse()

                    'benötigt sollte nach dem Struct eine Bool Variable folgen (Adressierung)
                    MeldungsBoolafter_others = True

                    'Structname aus dem Element lesen
                    MeldungStructName = Meldung.FirstAttribute.Value

                    'Die Elemente im Node Struct "holen"
                    Dim StructElement = (From element In Meldung.Nodes Select element)

                    For i As Integer = 1 To StructElement.Count - 1
                        CountDBAdresse()
                        Dim StructMeldung As XElement = StructElement.ElementAt(i)

                        'Werte hinzufügen
                        Meldungen.Add(New HMIAlarms With {.AlarmText = MeldungStructName & " " & StructMeldung.Descendants(SiemensNamespace + "MultiLanguageText").Value,
                         .Meldeklasse = Meldeklassenname, .Name = Meldung.FirstAttribute.Value & ID, .Datentyp = StructMeldung.@Datatype.ToString,
                        .ID = ID, .TriggerTag = AddressTag, .TrigerBit = AddressBitforArray})
                        ID = ID + 1

                        Meldungcounter = Meldungcounter + 1

                    Next


                    'Wenn der Member vom Datentyp Array ist dann:
                    'alle Meldungen im Struct lesen generieren
                ElseIf Meldung.@Datatype.ToString Like "Array*" Then
                    AddressBit = 6
                    CountDBAdresse()
                    MeldungsBoolafter_others = True

                    Dim GETDatatyp As String = Meldung.@Datatype

                    Dim ArrayBeginn As String = Meldung.@Datatype
                    ArrayBeginn = ArrayBeginn.Substring(6)
                    ArrayBeginn = ArrayBeginn.Remove(1)


                    Dim ArrayEnde As String = Meldung.@Datatype

                    ArrayEnde = ArrayEnde.Substring(ArrayEnde.LastIndexOf("."))
                    ArrayEnde = ArrayEnde.Remove(ArrayEnde.IndexOf("]"))
                    ArrayEnde = ArrayEnde.Replace(".", "")




                    GETDatatyp = StrReverse(GETDatatyp)
                    GETDatatyp = GETDatatyp.Remove(GETDatatyp.LastIndexOf("fo"))
                    GETDatatyp = StrReverse(GETDatatyp)
                    GETDatatyp = GETDatatyp.Replace(" ", "")


                    Select Case GETDatatyp
                        Case "Byte"
                            For Arraynummer As Integer = CInt(ArrayBeginn) To CInt(ArrayEnde)
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

                        ' 
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

                    'MBA:
                    ' Datentypen auswerten
                    ' Hier wird der Name des Datentyps mit den Datentypen im Container verglichen,
                    ' bei übereinstimmung des namens werden die Meldungen hinzugefügt.
                Else
                    MeldungsBoolafter_others = True
                    AddressBit = 6
                    CountDBAdresse()
                    Dim TypName As String

                    TypName = Meldung.FirstAttribute.Value



                    Const quote As String = """"
                    Dim TyponeHochkomma As String = Meldung.LastAttribute.Value.Replace(quote, "")

                    Dim LO_Type = (From Element In Datentypen Where Element.Typname = TyponeHochkomma)

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

    'MBA:
    'XML Datentypen in Container legen
    'es Liegen alle Datentypen gesammelt in dem Container, es wird nur nach dem Namen ausgewählt.
    Private Sub GetDatatyp(ByVal Pfad As String)
        _StatusChanged("Datentypen auslesen")

        Dim XMLFile As XDocument
        XMLFile = XDocument.Load(Pfad)



        Dim SiemensNamespace As XNamespace = "http://www.siemens.com/automation/Openness/SW/Interface/v1" ' must match declaration In document



        Dim Name As XElement = (From el In XMLFile.<Document>.<SW.ControllerDatatype>.<AttributeList>
                                Select el).Last.LastNode



        Dim Interface_Sections As XElement =
            (From el In XMLFile.<Document>.<SW.ControllerDatatype>.<AttributeList>.<Interface>
             Select el).First




        Dim SelectionsElemente As IEnumerable(Of XElement) =
        From element In Interface_Sections.Elements(SiemensNamespace + "Sections")
        Select element

        Dim SectionElement As XElement = SelectionsElemente.ElementAt(0)



        Dim Selection = From element In SelectionsElemente.Elements(SiemensNamespace + "Section") Select element

        Dim Typklassen = (From element In Selection.Elements(SiemensNamespace + "Member") Select element)

        For i As Integer = 0 To Typklassen.Count - 1

            Dim Typmeldungen As XElement = Typklassen.ElementAt(i)



            Datentypen.Add(New HMIAlarms With {.AlarmText = Typmeldungen.Descendants(SiemensNamespace + "MultiLanguageText").Value,
                                 .Name = Typmeldungen.FirstAttribute.Value, .Datentyp = Typmeldungen.LastAttribute.Value, .Typname = Name.Value})



        Next




    End Sub

    'Excel befüllen und speichern
    Public Sub Write_Excel()

        System.IO.Directory.CreateDirectory(GetFolderPath(SpecialFolder.MyDocuments) & "\Meldegenerator_HMI_Alarms")



        ExcelDatenEinfügen()

        ExcelSpeichern(GetFolderPath(SpecialFolder.MyDocuments) & "\Meldegenerator_HMI_Alarms\HMIAlarms_" & CPUName & ".xlsx")



    End Sub





    'constanten für Excel Header
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


    'MBA:
    'Ein neues Excel erstellen

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

    ''' <summary>
    ''' Alle Meldungen die im Container ligen werden in das Excel geschrieben.
    ''' </summary>
    Sub ExcelDatenEinfügen()
        _StatusChanged("Daten In Excel File schreiben")

        Dim i As Integer = 1
        sheet.Cells(1, 1) = Column0_A
        sheet.Cells(1, 2) = Column1_B
        If inWords = False Then
            sheet.Cells(1, 3) = Column2_C
        Else
            sheet.Cells(1, 3) = "Alarm text [de-DE], Alarm text 1"
        End If

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


    'Excel Speichern und Schliessen. (Prozess)
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

            wkbk.SaveAs(filePath)


            wkbk.Close()
            sheet = Nothing
            wkbk = Nothing

            excelApp.Quit()
            excelApp = Nothing

            _StatusChanged("Fertig")
        Catch ex As System.Runtime.InteropServices.COMException
            _StatusChanged("Fehler beim EXCEL speichern")
            MsgBox("Das Excel File konnte nicht gespeichert werden, es ist vielleicht geöffnet.")
        End Try

        ' MsgBox("Fertig")
    End Sub

End Class



'Die "Vorlage" einer Meldung 
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