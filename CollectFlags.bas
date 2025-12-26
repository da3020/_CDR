Option Explicit

' =========================================================
' Version: 1.3.1
' CorelDRAW: 2021 (25.0.0.230)
'
' Изменения:
' - поддержка ведущих нулей в именах групп
' - поиск группы по артикулу + размеру
' - информативный текст при отсутствии группы
' =========================================================


' ================== НАСТРОЙКИ ============================

Const ROWS_PER_COLUMN As Long = 5
Const START_X As Double = 10
Const START_Y As Double = 280
Const ROW_GAP As Double = 10
Const COL_GAP As Double = 20

Const ERROR_TEXT_SIZE As Double = 200
Const ERR_R As Long = 255
Const ERR_G As Long = 0
Const ERR_B As Long = 0


' ================== ОСНОВНОЙ МАКРОС ======================

Sub CollectFlagsToPrint()

    Dim ORDERS_FILE As String
    Dim INDEX_FILE As String
    Dim OUTPUT_FILE As String

    ORDERS_FILE = "D:\_CDR\orders.txt"
    INDEX_FILE = "D:\_CDR\_FLAG_INDEX.txt"
    OUTPUT_FILE = "D:\_CDR\На печать_001.cdr"

    Dim indexMap As Object
    Set indexMap = LoadIndex(INDEX_FILE)

    Dim outputDoc As Document
    Set outputDoc = Application.CreateDocument

    Application.Optimization = True
    Application.EventsEnabled = False
    Application.Visible = False

    Dim orders As Collection
    Set orders = ReadUtf8Lines(ORDERS_FILE)

    Dim i As Long
    i = 0

    Dim line As Variant
    For Each line In orders

        line = Trim(line)
        If line = "" Then GoTo NextLine

        Dim baseArticle As String
        Dim suffix As String
        ParseArticle CStr(line), baseArticle, suffix

        Dim sizePart As String
        sizePart = GetSizePartBySuffix(suffix)

        Dim placedShape As Shape

        ' -------- ФАЙЛ НЕ НАЙДЕН --------
        If Not indexMap.Exists(baseArticle) Then

            Set placedShape = CreateErrorText(outputDoc, _
                baseArticle & " — ФАЙЛ НЕ НАЙДЕН")

        ' -------- ДУБЛИКАТЫ --------
        ElseIf indexMap(baseArticle).Count > 1 Then

            Set placedShape = CreateErrorText(outputDoc, _
                baseArticle & " — НАЙДЕНЫ ДУБЛИКАТЫ")

        ' -------- ФАЙЛ НАЙДЕН --------
        Else

            Dim path As String
            path = indexMap(baseArticle)(1)

            Dim doc As Document
            Set doc = Application.OpenDocument(path)

            Dim shp As Shape
            Set shp = FindGroupByArticleAndSize(doc, baseArticle, sizePart)

            If Not shp Is Nothing Then
                shp.Copy
                Set placedShape = outputDoc.ActiveLayer.Paste
            Else
                Set placedShape = CreateErrorText(outputDoc, _
                    baseArticle & " — " & suffix & " " & _
                    baseArticle & ":" & sizePart & _
                    " ГРУППА НЕ НАЙДЕНА")
            End If

            doc.Close
        End If

        PlaceInGrid placedShape, i
        i = i + 1

NextLine:
    Next line

    outputDoc.SaveAs OUTPUT_FILE

    Application.Visible = True
    Application.Optimization = False
    Application.EventsEnabled = True

    MsgBox "Готово. Объектов: " & i, vbInformation

End Sub


' ================== ИНДЕКС ===============================

Function LoadIndex(path As String) As Object

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lines As Collection
    Set lines = ReadUtf8Lines(path)

    Dim line As Variant
    For Each line In lines

        Dim parts() As String
        parts = Split(line, "|")
        If UBound(parts) <> 1 Then GoTo NextLine

        Dim key As String
        key = parts(0)

        If Not dict.Exists(key) Then
            dict.Add key, New Collection
        End If

        dict(key).Add parts(1)

NextLine:
    Next line

    Set LoadIndex = dict

End Function


' ================== УТИЛИТЫ ===============================

Sub ParseArticle(src As String, ByRef baseArticle As String, ByRef suffix As String)

    Dim i As Long
    baseArticle = ""
    suffix = ""

    For i = 1 To Len(src)
        If Mid(src, i, 1) Like "[0-9]" Then
            baseArticle = baseArticle & Mid(src, i, 1)
        Else
            suffix = UCase(Mid(src, i))
            Exit For
        End If
    Next i

End Sub


Function GetSizePartBySuffix(suffix As String) As String

    Select Case suffix
        Case "S": GetSizePartBySuffix = "60x40"
        Case "M": GetSizePartBySuffix = "105x70"
        Case "L": GetSizePartBySuffix = "225x150"
        Case Else: GetSizePartBySuffix = "135x90"
    End Select

End Function


' ================== ПОИСК ГРУППЫ ==========================

Function FindGroupByArticleAndSize(doc As Document, _
                                   article As String, _
                                   sizePart As String) As Shape

    Dim p As Page
    Dim s As Shape

    Dim targetArticle As Long
    targetArticle = CLng(article)

    For Each p In doc.Pages
        For Each s In p.Shapes.All

            If s.Type = cdrGroupShape Then

                Dim grpName As String
                grpName = s.Name   ' например 0264:225x150

                If InStr(grpName, ":") > 0 Then

                    Dim parts() As String
                    parts = Split(grpName, ":")

                    If UBound(parts) = 1 Then

                        Dim grpArticle As Long
                        On Error Resume Next
                        grpArticle = CLng(parts(0))
                        On Error GoTo 0

                        If grpArticle = targetArticle _
                           And parts(1) = sizePart Then

                            Set FindGroupByArticleAndSize = s
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next s
    Next p

    Set FindGroupByArticleAndSize = Nothing

End Function


' ================== РАСКЛАДКА ============================

Sub PlaceInGrid(s As Shape, index As Long)

    Dim row As Long
    Dim col As Long

    row = index Mod ROWS_PER_COLUMN
    col = index \ ROWS_PER_COLUMN

    s.SetPosition _
        START_X + col * (s.SizeWidth + COL_GAP), _
        START_Y - row * (s.SizeHeight + ROW_GAP)

End Sub


' ================== ТЕКСТ ОШИБОК =========================

Function CreateErrorText(doc As Document, txt As String) As Shape

    Dim s As Shape
    Set s = doc.ActiveLayer.CreateArtisticText(0, 0, txt)

    With s.Text.Story.Characters.All
        .Size = ERROR_TEXT_SIZE
        .Fill.UniformColor.RGBAssign ERR_R, ERR_G, ERR_B
    End With

    Set CreateErrorText = s

End Function


' ================== UTF-8 ================================

Function ReadUtf8Lines(filePath As String) As Collection

    Dim lines As New Collection
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    With stm
        .Type = 2
        .CharSet = "utf-8"
        .Open
        .LoadFromFile filePath
        Do Until .EOS
            lines.Add Trim(.ReadText(-2))
        Loop
        .Close
    End With

    Set ReadUtf8Lines = lines

End Function
