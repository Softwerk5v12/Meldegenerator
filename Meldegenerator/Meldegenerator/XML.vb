Imports System.Linq
Imports System.Xml.Linq
Imports System.IO
Imports System.Xml.Schema


'Imports <xmlns='http://www.siemens.com/automation/Openness/SW/Interface/v1'>

Public Class XML
    Friend XMLFile As XDocument
    Public Sub LoadXML()

        XMLFile = XDocument.Load("D:\01_Lokale_Projekte\Openness\XMLs\VM.xml")




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


        If (SectionElement.Name.Namespace Is XNamespace.None) Then
                Console.WriteLine("The element el2 is in no namespace.")
            Else
                Console.WriteLine("The element el2 is in a namespace.")
                'a = el2.Name.NamespaceName
                '  Dim Node1 As XNode = el2.FirstNode

            End If


            ' Console.WriteLine(SelectionsElemente.)

            '   Dim AlleMeldungen As IEnumerable(Of XElement) = (From element In SelectionElement.Descendants(SiemensNamespace + "Member").Attributes Where element.Value = "M_").First

            '   Dim LO_Meldungen As List(Of XElement) = SelectionElement.Descendants(SiemensNamespace + "MultiLanguageText").ToList

            '  Dim LO_Meldungen = (From element In SelectionsElemente(SiemensNamespace + "Member") Select element.@Name = "_M")
            Dim Selection = From element In SelectionsElemente.Elements(SiemensNamespace + "Section") Select element

            Dim Meldeklassen = (From element In Selection.Elements(SiemensNamespace + "Member") Select element)
            '   Console.WriteLine(SelectionElement.Descendants(SiemensNamespace + "MultiLanguageText").Skip(1).Take(20).Value)


            Dim Meldungen As New List(Of HMIAlarms)
            Dim Störungen As New List(Of HMIAlarms)


            Dim Meldeklasse = (From element In Meldeklassen Where element.FirstAttribute.Value = "M_")



            For Each Meldung As XElement In Meldeklasse.Elements
                ' Console.WriteLine(Meldeklassen.Elements)

                Console.WriteLine(Meldung.Name)
                '  If Meldung.Name Then
                '   If Meldung.Parent.FirstAttribute = "M_" Then

                If Meldung.Name = "{" & SiemensNamespace.ToString & "}Member" Then


                    Meldungen.Add(New HMIAlarms With {.AlarmText = Meldung.Descendants(SiemensNamespace + "MultiLanguageText").Value,
                                          .Meldeklasse = "Meldungen", .Name = Meldung.FirstAttribute.ToString, .Datentyp = Meldung.@Datatype.ToString})


                End If

            Next

            Dim Störklasse = (From element In Meldeklassen Where element.FirstAttribute.Value = "S_")


            For Each Störung As XElement In Störklasse.Elements
                ' Console.WriteLine(Meldeklassen.Elements)

                Console.WriteLine(Störung.Name)
                '  If Meldung.Name Then
                '   If Meldung.Parent.FirstAttribute = "M_" Then

                If Störung.Name = "{" & SiemensNamespace.ToString & "}Member" Then


                Dim Alarmtext As String = Nothing
                Dim StructName As String = Nothing


                If Störung.@Datatype.ToString = "Struct" Then

                    StructName = Störung.FirstAttribute.ToString

                    Console.WriteLine(StructName)

                    ' Dim StructElement As XElement = (From element In Störklasse Where element.FirstAttribute.Value = "LSp_5")



                End If

                'Dim MKAttributes As IEnumerable(Of XAttribute) = Meldung.Attributes.ToList

                Störungen.Add(New HMIAlarms With {.AlarmText = Störung.Descendants(SiemensNamespace + "MultiLanguageText").Value,
                                          .Meldeklasse = "Meldungen", .Name = Störung.FirstAttribute.ToString, .Datentyp = Störung.@Datatype.ToString})


                End If


            Next

        For Each i As HMIAlarms In Meldungen
            Console.WriteLine(i.AlarmText)
        Next


        Console.WriteLine()


        '   Next




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