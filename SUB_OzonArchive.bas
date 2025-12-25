Option Explicit

' =========================================================
' Version: 1.2.1
' CorelDRAW: 2021 (25.0.0.230)
'
' Функции:
' - orders.txt = список артикулов (8457, 4855M, 7898L ...)
' - поиск CDR по артикулу в имени (с ведущими нулями)
' - рекурсивный обход всех подпапок
' - суффикс > размер группы
' - дубликаты > визуальная ошибка
'
' КРИТИЧНО:
' - Форматирование текста: Characters.All.Size
' - Пути объявлять через Dim
' =========================================================


' ================== НАСТРОЙКИ ============================

' --- Раскладка ---
Const ROWS_PER_COLUMN As Long = 5
Const START_X As Double = 10
Const START_Y As Double = 280
Const ROW_GAP As Double = 10
Const COL_GAP As Double = 20

' --- Текст ошибок ---
Const ERROR_TEXT_SIZE As Double = 200
Const ERR_R As Long = 255
Const ERR_G As Long = 0
Const ERR_B As Long = 0


' ================== ОСНОВНОЙ МАКРОС ======================

Sub CollectFlagsToPrint()

    ' --- Пути (UNC разрешён, экранирование не нужно) ---
    Dim SOURCE_FOLDER As String
    Dim ORDERS_FILE As String
    Dim OUTPUT_FILE As String

    SOURCE_FOLDER = "\\Keenetic-5026\ugreen\STORE\! СУБЛИМАЦИЯ\! ! ! ФЛАГИ\"
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

        line = Trim(line)
        If line = "" Then GoTo NextLine

        Dim baseArticle As String
        Dim suffix As String
        ParseArticle CStr(line), baseArticle, suffix

        Dim matches As Collection
        Set matches = FindCdrFilesByArticle(SOURCE_FOLDER, baseArticle)

        Dim placedShape As Shape

        ' ---- ФАЙЛ НЕ НАЙДЕН ----
        If matches.Count = 0 Then

            Set placedShape = CreateErrorText(outputDoc, _
                baseArticle & " — ФАЙЛ НЕ НАЙДЕН")

        ' ---- НАЙДЕНЫ ДУБЛИКАТЫ ----
        ElseIf matches.Count > 1 Then

            Set placedShape = CreateErrorText(outputDoc, _
                baseArticle & " — НАЙДЕНЫ ДУБЛИКАТЫ")

        ' ---- НАЙДЕН РОВНО ОДИН ФАЙЛ ----
        Else

            Dim doc As Document
            Set doc = Application.OpenDocument(matches(1))

            Dim groupName As String
            groupName = GetGroupName(baseArticle, suffix)

            Dim shp As Shape
            Set shp = FindGroupByName(doc, groupName)

            If Not shp Is Nothing Then
                shp.Copy
                Set placedShape = outputDoc.ActiveLayer.Paste
            Else
                Set placedShape = CreateErrorText(outputDoc, _
                    baseArticle & " — ГРУППА НЕ НАЙДЕНА")
            End If

            doc.Close
        End If

        PlaceInGrid placedShape, index
        index = index + 1

NextLine:
    Next line

    outputDoc.SaveAs OUTPUT_FILE

    Application.Visible = True
    Application.Optimization = False
    Application.EventsEnabled = True

    MsgBox "Готово." & vbCrLf & _
           "Размещено объектов: " & index, vbInformation

End Sub


' ================== РАЗБОР АРТИКУЛА ======================

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


' ================== ПОИСК CDR (РЕКУРСИВНО) ===============

Function FindCdrFilesByArticle(rootFolder As String, article As String) As Collection

    Dim result As New Collection
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ScanFolderRecursive fso.GetFolder(rootFolder), article, result

    Set FindCdrFilesByArticle = result

End Function


Sub ScanFolderRecursive(fld As Object, article As String, result As Collection)

    Dim f As Object
    For Each f In fld.Files
        If LCase(GetExtension(f.Name)) = "cdr" Then
            If IsArticleMatch(f.Name, article) Then
                result.Add f.path
            End If
        End If
    Next f

    Dim subFld As Object
    For Each subFld In fld.SubFolders
        ScanFolderRecursive subFld, article, result
    Next subFld

End Sub


Function GetExtension(fileName As String) As String
    GetExtension = Mid(fileName, InStrRev(fileName, ".") + 1)
End Function


Function IsArticleMatch(fileName As String, article As String) As Boolean

    Dim pos As Long
    pos = InStr(fileName, "_")
    If pos = 0 Then Exit Function

    Dim namePart As String
    namePart = Left(fileName, pos - 1)

    On Error Resume Next
    namePart = CStr(CLng(namePart)) ' убираем ведущие нули
    On Error GoTo 0

    IsArticleMatch = (namePart = article)

End Function


' ================== СОПОСТАВЛЕНИЕ ГРУПП ==================

Function GetGroupName(article As String, suffix As String) As String

    Select Case suffix
        Case "S"
            GetGroupName = article & ":60x40"
        Case "M"
            GetGroupName = article & ":105x70"
        Case "L"
            GetGroupName = article & ":225x150"
        Case Else
            GetGroupName = article & ":135x90"
    End Select

End Function


' ================== РАСКЛАДКА ============================

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


' ================== ПОИСК ГРУППЫ =========================

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


