Option Explicit

' =========================================================
' ================== НАСТРОЙКИ ============================
' =========================================================



' --- Раскладка ---
Const ROWS_PER_COLUMN As Long = 5      ' сколько объектов в столбце
Const START_X As Double = 10            ' мм, левый край
Const START_Y As Double = 280           ' мм, верх страницы
Const ROW_GAP As Double = 10            ' мм, между строками
Const COL_GAP As Double = 20            ' мм, между столбцами

' --- Текст для отсутствующих файлов ---
Const MISSING_TEXT_SIZE As Double = 200  ' pt
Const MISSING_TEXT_COLOR_R As Long = 255
Const MISSING_TEXT_COLOR_G As Long = 0
Const MISSING_TEXT_COLOR_B As Long = 0

' =========================================================
' ================== ОСНОВНОЙ МАКРОС ======================
' =========================================================

Sub CollectFlagsToPrint()

' --- Пути ---
Dim SOURCE_FOLDER As String
Dim ORDERS_FILE As String
Dim OUTPUT_FILE As String

SOURCE_FOLDER = "D:\_CDR\archive\"
ORDERS_FILE = "D:\_CDR\orders.txt"
OUTPUT_FILE = "D:\_CDR\На печать_001.cdr"



    Dim outputDoc As Document
    Set outputDoc = Application.CreateDocument

    Application.Optimization = True
    Application.EventsEnabled = False
    Application.Visible = False

    Dim orders As Collection
    Set orders = ReadUtf8Lines(ORDERS_FILE)

    Dim index As Long
    index = 0

    Dim line As Variant
    For Each line In orders

        If Trim(line) <> "" Then

            Dim parts() As String
            parts = Split(line, "_")

            If UBound(parts) < 1 Then GoTo NextLine

            Dim fileName As String
            Dim groupName As String

            fileName = parts(0)
            groupName = parts(1)

            Dim placedShape As Shape
            Dim fullPath As String
            fullPath = SOURCE_FOLDER & fileName

            ' ---------- ЕСЛИ ФАЙЛ СУЩЕСТВУЕТ ----------
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
                        fileName & "_" & groupName & " — ГРУППА НЕ НАЙДЕНА")
                End If

                doc.Close

            ' ---------- ЕСЛИ ФАЙЛА НЕТ ----------
            Else
                Set placedShape = CreateMissingText(outputDoc, _
                    fileName & " — ФАЙЛ НЕ НАЙДЕН")
            End If

            ' ---------- РАСКЛАДКА ----------
            PlaceInGrid placedShape, index
            index = index + 1

        End If

NextLine:
    Next line

    outputDoc.SaveAs OUTPUT_FILE

    Application.Visible = True
    Application.Optimization = False
    Application.EventsEnabled = True

    MsgBox "Готово." & vbCrLf & _
           "Размещено объектов: " & index, vbInformation

End Sub

' =========================================================
' ================== РАСКЛАДКА ============================
' =========================================================

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

' =========================================================
' ============ ТЕКСТ ПРИ ОШИБКАХ ==========================
' =========================================================

Function CreateMissingText(doc As Document, txt As String) As Shape

    Dim s As Shape
    Set s = doc.ActiveLayer.CreateArtisticText(0, 0, txt)

    With s.Text.Story.Characters.All
        .Size = MISSING_TEXT_SIZE
        .Fill.UniformColor.RGBAssign 255, 0, 0
    End With

    Set CreateMissingText = s

End Function


' =========================================================
' ================= ПОИСК ГРУППЫ ==========================
' =========================================================

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

' =========================================================
' ============ ПРОВЕРКА ФАЙЛА =============================
' =========================================================

Function FileExists(path As String) As Boolean
    FileExists = (Dir(path) <> "")
End Function

' =========================================================
' ============ ЧТЕНИЕ UTF-8 ===============================
' =========================================================

Function ReadUtf8Lines(filePath As String) As Collection

    Dim lines As New Collection
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    With stm
        .Type = 2            ' text
        .CharSet = "utf-8"
        .Open
        .LoadFromFile filePath

        Do Until .EOS
            lines.Add Trim(.ReadText(-2)) ' читать построчно
        Loop

        .Close
    End With

    Set ReadUtf8Lines = lines

End Function


