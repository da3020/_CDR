Attribute VB_Name = "SUB_OzonArchive"
Option Explicit

' ===== РАСКЛАДКА =====
Const ROWS_PER_COLUMN As Long = 5

Const START_X As Double = 10    ' мм
Const START_Y As Double = 280   ' мм (верх страницы)

Const ROW_GAP As Double = 10    ' расстояние между строками (мм)
Const COL_GAP As Double = 20    ' расстояние между столбцами (мм)

' ===== ТЕКСТ ЕСЛИ ФАЙЛ НЕ НАЙДЕН =====
Const MISSING_TEXT_SIZE As Double = 20 ' pt
' =========





Sub CollectFlagsToPrint()

    Dim sourceFolder As String
    Dim ordersFile As String
    Dim outputFile As String

    sourceFolder = "D:\_CDR\archive\"
    ordersFile = "D:\_CDR\orders.txt"
    outputFile = "D:\_CDR\На печать_001.cdr"

    Dim outputDoc As Document
    Set outputDoc = Application.CreateDocument

    Application.Optimization = True
    Application.EventsEnabled = False

    Dim orders As Collection
    Set orders = ReadUtf8Lines(ordersFile)

    Dim index As Long
    index = 0

    Dim line As Variant
    For Each line In orders

        If line <> "" Then

            Dim parts() As String
            parts = Split(line, "_")

            Dim fileName As String
            Dim groupName As String

            fileName = parts(0)
            groupName = parts(1)

            Dim placedShape As Shape
            Dim fullPath As String
            fullPath = sourceFolder & fileName

            If FileExists(fullPath) Then

                Dim doc As Document
                Set doc = Application.OpenDocument(fullPath)

                Dim shp As Shape
                Set shp = FindGroupByName(doc, groupName)

                If Not shp Is Nothing Then
                    shp.Copy
                    Set placedShape = outputDoc.ActiveLayer.Paste
                Else
                    Set placedShape = CreateMissingText(outputDoc, _
                        fileName & "_" & groupName & " НЕ НАЙДЕНА ГРУППА")
                End If

                doc.Close

            Else
                Set placedShape = CreateMissingText(outputDoc, _
                    fileName & " НЕ НАЙДЕН")
            End If

            Call PlaceInGrid(placedShape, index)
            index = index + 1

        End If
    Next line

    outputDoc.SaveAs outputFile

    Application.Optimization = False
    Application.EventsEnabled = True

    MsgBox "Готово. Разложено: " & index & " объектов."

End Sub

Sub PlaceInGrid(s As Shape, index As Long)

    Dim row As Long
    Dim col As Long

    row = index Mod ROWS_PER_COLUMN
    col = index \ ROWS_PER_COLUMN

    Dim x As Double
    Dim y As Double

    x = START_X + col * (s.SizeWidth + COL_GAP)
    y = START_Y - row * (s.SizeHeight + ROW_GAP)

    s.SetPosition x, y

End Sub




Function FileExists(path As String) As Boolean
    FileExists = (Dir(path) <> "")
End Function

Function CreateMissingText(doc As Document, txt As String) As Shape

    Dim s As Shape
    Set s = doc.ActiveLayer.CreateArtisticText(0, 0, txt)

    With s.Text.Story.TextRange
        .FONTSIZE = MISSING_TEXT_SIZE
        .Fill.UniformColor.RGBAssign 255, 0, 0
    End With

    Set CreateMissingText = s

End Function






Function FindGroupByName(doc As Document, groupName As String) As Shape

    Dim p As Page
    Dim s As Shape

    For Each p In doc.Pages
        For Each s In p.Shapes.All
            If s.Type = cdrGroupShape Then
                If s.Name = groupName Then
                    Set FindGroupByName = s
                    Exit Function
                End If
            End If
        Next s
    Next p

    Set FindGroupByName = Nothing

End Function


Function ReadUtf8Lines(filePath As String) As Collection

    Dim lines As New Collection
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    With stm
        .Type = 2 ' text
        .CharSet = "utf-8"
        .Open
        .LoadFromFile filePath

        Do Until .EOS
            lines.Add Trim(.ReadText(-2)) ' -2 = read line
        Loop

        .Close
    End With

    Set ReadUtf8Lines = lines

End Function


