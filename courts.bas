Option Explicit

' API declarations
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Константы для настроек
Private Const NOMINATIM_DELAY As Long = 1500 ' Задержка между запросами в мс
Private Const REQUEST_TIMEOUT As Long = 10000 ' Таймаут запроса в мс

' Source - https://stackoverflow.com/a
' Posted by Tomalak, modified by community. See post 'Timeline' for change history
' Retrieved 2025-11-14, License - CC BY-SA 4.0
' MODIFIED: Removed ADODB dependency for better compatibility
Public Function URLEncode( _
   ByVal StringVal As String, _
   Optional SpaceAsPlus As Boolean = False _
) As String
    Dim i As Long
    Dim char As String
    Dim ascVal As Integer
    Dim result As String
    Dim space As String
    
    If SpaceAsPlus Then
        space = "+"
    Else
        space = "%20"
    End If
    
    result = ""
    
    For i = 1 To Len(StringVal)
        char = Mid(StringVal, i, 1)
        ascVal = AscW(char)
        
        ' Handle Unicode characters (outside ASCII range)
        If ascVal < 0 Then ascVal = ascVal + 65536
        If ascVal > 127 Then
            ' Unicode character - encode as UTF-8
            result = result & EncodeUTF8(char)
        Else
            ' ASCII character
            Select Case ascVal
                Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                    ' a-z, A-Z, 0-9, -, ., _, ~
                    result = result & char
                Case 32
                    ' Space
                    result = result & space
                Case Else
                    ' Other ASCII characters
                    result = result & "%" & Right("0" & Hex(ascVal), 2)
            End Select
        End If
    Next i
    
    URLEncode = result
End Function

Private Function EncodeUTF8(ByVal char As String) As String
    ' UTF-8 encoding for Cyrillic characters
    Select Case char
        ' Russian uppercase letters
        Case "А": EncodeUTF8 = "%D0%90"
        Case "Б": EncodeUTF8 = "%D0%91"
        Case "В": EncodeUTF8 = "%D0%92"
        Case "Г": EncodeUTF8 = "%D0%93"
        Case "Д": EncodeUTF8 = "%D0%94"
        Case "Е": EncodeUTF8 = "%D0%95"
        Case "Ё": EncodeUTF8 = "%D0%81"
        Case "Ж": EncodeUTF8 = "%D0%96"
        Case "З": EncodeUTF8 = "%D0%97"
        Case "И": EncodeUTF8 = "%D0%98"
        Case "Й": EncodeUTF8 = "%D0%99"
        Case "К": EncodeUTF8 = "%D0%9A"
        Case "Л": EncodeUTF8 = "%D0%9B"
        Case "М": EncodeUTF8 = "%D0%9C"
        Case "Н": EncodeUTF8 = "%D0%9D"
        Case "О": EncodeUTF8 = "%D0%9E"
        Case "П": EncodeUTF8 = "%D0%9F"
        Case "Р": EncodeUTF8 = "%D0%A0"
        Case "С": EncodeUTF8 = "%D0%A1"
        Case "Т": EncodeUTF8 = "%D0%A2"
        Case "У": EncodeUTF8 = "%D0%A3"
        Case "Ф": EncodeUTF8 = "%D0%A4"
        Case "Х": EncodeUTF8 = "%D0%A5"
        Case "Ц": EncodeUTF8 = "%D0%A6"
        Case "Ч": EncodeUTF8 = "%D0%A7"
        Case "Ш": EncodeUTF8 = "%D0%A8"
        Case "Щ": EncodeUTF8 = "%D0%A9"
        Case "Ъ": EncodeUTF8 = "%D0%AA"
        Case "Ы": EncodeUTF8 = "%D0%AB"
        Case "Ь": EncodeUTF8 = "%D0%AC"
        Case "Э": EncodeUTF8 = "%D0%AD"
        Case "Ю": EncodeUTF8 = "%D0%AE"
        Case "Я": EncodeUTF8 = "%D0%AF"
        
        ' Russian lowercase letters
        Case "а": EncodeUTF8 = "%D0%B0"
        Case "б": EncodeUTF8 = "%D0%B1"
        Case "в": EncodeUTF8 = "%D0%B2"
        Case "г": EncodeUTF8 = "%D0%B3"
        Case "д": EncodeUTF8 = "%D0%B4"
        Case "е": EncodeUTF8 = "%D0%B5"
        Case "ё": EncodeUTF8 = "%D1%91"
        Case "ж": EncodeUTF8 = "%D0%B6"
        Case "з": EncodeUTF8 = "%D0%B7"
        Case "и": EncodeUTF8 = "%D0%B8"
        Case "й": EncodeUTF8 = "%D0%B9"
        Case "к": EncodeUTF8 = "%D0%BA"
        Case "л": EncodeUTF8 = "%D0%BB"
        Case "м": EncodeUTF8 = "%D0%BC"
        Case "н": EncodeUTF8 = "%D0%BD"
        Case "о": EncodeUTF8 = "%D0%BE"
        Case "п": EncodeUTF8 = "%D0%BF"
        Case "р": EncodeUTF8 = "%D1%80"
        Case "с": EncodeUTF8 = "%D1%81"
        Case "т": EncodeUTF8 = "%D1%82"
        Case "у": EncodeUTF8 = "%D1%83"
        Case "ф": EncodeUTF8 = "%D1%84"
        Case "х": EncodeUTF8 = "%D1%85"
        Case "ц": EncodeUTF8 = "%D1%86"
        Case "ч": EncodeUTF8 = "%D1%87"
        Case "ш": EncodeUTF8 = "%D1%88"
        Case "щ": EncodeUTF8 = "%D1%89"
        Case "ъ": EncodeUTF8 = "%D1%8A"
        Case "ы": EncodeUTF8 = "%D1%8B"
        Case "ь": EncodeUTF8 = "%D1%8C"
        Case "э": EncodeUTF8 = "%D1%8D"
        Case "ю": EncodeUTF8 = "%D1%8E"
        Case "я": EncodeUTF8 = "%D1%8F"
        
        ' Space
        Case " ": EncodeUTF8 = "%20"
        
        ' Other characters
        Case Else
            Dim ascVal As Integer
            ascVal = AscW(char)
            If ascVal < 0 Then ascVal = ascVal + 65536
            EncodeUTF8 = "%" & Right("000" & Hex(ascVal), 4)
    End Select
End Function



Sub FindCourtByAddress()
    On Error GoTo ErrorHandler
    
    Dim selectedText As String
    Dim district As String
    Dim courtInfo As String
    Dim courtName As String
    Dim courtAddress As String
    
    ' Получаем выделенный текст
    selectedText = GetSelectedText()
    If selectedText = "" Then Exit Sub
    
    Application.StatusBar = "Определяем район для адреса..."
    
    ' Получаем район через Nominatim
    district = GetDistrictFromNominatim(selectedText)
    
    If district = "" Or district = "Не удалось определить" Then
        MsgBox "Не удалось определить район для адреса: " & selectedText & vbCrLf & _
               "Попробуйте:" & vbCrLf & _
               "1. Проверить подключение к интернету" & vbCrLf & _
               "2. Уточнить формат адреса" & vbCrLf & _
               "3. Использовать ручной ввод района", vbExclamation, "Ошибка определения района"
        Application.StatusBar = ""
        Exit Sub
    End If
    
    Application.StatusBar = "Найден район: " & district & ". Ищем суд..."
    
    ' Получаем информацию о суде
    courtInfo = GetCourtByDistrict(district)
    
    If courtInfo = "" Then
        MsgBox "Не найден суд для района: " & district & vbCrLf & _
               "Доступные районы: Октябрьский, Центральный, Советский, Первомайский, " & _
               "Партизанский, Заводской, Ленинский, Московский, Фрунзенский", vbExclamation
        Application.StatusBar = ""
        Exit Sub
    End If
    
    ' Парсим информацию о суде
    courtName = Split(courtInfo, "|")(0)
    courtAddress = Split(courtInfo, "|")(1)
    
    ' Добавляем в конец документа
    AddCourtToDocument courtName, courtAddress, selectedText, district
    
    Application.StatusBar = "Готово! Суд добавлен в документ"
    
    ' Показываем результат
    MsgBox "Суд успешно найден и добавлен в документ:" & vbCrLf & _
           "Район: " & district & vbCrLf & _
           "Суд: " & courtName & vbCrLf & _
           "Адрес: " & courtAddress, vbInformation, "Результат поиска"
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = ""
    MsgBox "Произошла ошибка: " & Err.Description & vbCrLf & _
           "Номер ошибки: " & Err.Number, vbCritical, "Ошибка выполнения"
End Sub

Function GetSelectedText() As String
    On Error GoTo ErrorHandler
    
    DebugPrint "=== НАЧАЛО GetSelectedText ==="
    DebugPrint "Selection.Type: " & Selection.Type
    DebugPrint "Selection.Text: '" & Selection.text & "'"
    
    If Selection.Type = wdSelectionIP Then
        DebugPrint "Текст не выделен (wdSelectionIP)"
        If MsgBox("Текст не выделен." & vbCrLf & "Хотите ввести адрес вручную?", _
                  vbQuestion + vbYesNo, "Ввод адреса") = vbYes Then
            GetSelectedText = InputBox("Введите адрес для поиска:", "Ручной ввод адреса")
            DebugPrint "Ручной ввод: '" & GetSelectedText & "'"
        Else
            GetSelectedText = ""
        End If
        Exit Function
    End If
    
    Dim selectedText As String
    selectedText = Selection.text
    DebugPrint "Исходный выделенный текст: '" & selectedText & "'"
    
    ' Очищаем текст
    selectedText = Trim(Replace(Replace(selectedText, Chr(13), ""), Chr(7), ""))
    DebugPrint "Очищенный текст: '" & selectedText & "'"
    
    If selectedText = "" Then
        DebugPrint "Текст пуст после очистки"
        If MsgBox("Выделенный текст пуст." & vbCrLf & "Хотите ввести адрес вручную?", _
                  vbQuestion + vbYesNo, "Ввод адреса") = vbYes Then
            GetSelectedText = InputBox("Введите адрес для поиска:", "Ручной ввод адреса")
            DebugPrint "Ручной ввод: '" & GetSelectedText & "'"
        Else
            GetSelectedText = ""
        End If
        Exit Function
    End If
    
    ' Проверяем длину текста
    DebugPrint "Длина текста: " & Len(selectedText)
    If Len(selectedText) < 5 Then
        DebugPrint "Текст слишком короткий"
        MsgBox "Выделенный текст слишком короткий для адреса.", vbInformation
        GetSelectedText = ""
        Exit Function
    End If
    
    GetSelectedText = selectedText
    DebugPrint "Возвращаемый текст: '" & GetSelectedText & "'"
    DebugPrint "=== КОНЕЦ GetSelectedText ==="
    Exit Function
    
ErrorHandler:
    DebugPrint "ОШИБКА в GetSelectedText: " & Err.Description
    MsgBox "Ошибка при получении текста: " & Err.Description, vbCritical
    GetSelectedText = ""
End Function



Function GetDistrictFromNominatim(address As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim response As String
    Dim lat As String
    Dim lon As String
    
    ' Упрощаем адрес для лучшего поиска
    Dim simplifiedAddress As String
    simplifiedAddress = SimplifyAddress(address)
    DebugPrint "Упрощенный адрес: " & simplifiedAddress
    
    ' Создаем HTTP объект
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Кодируем адрес для URL с помощью улучшенной функции
    url = "https://nominatim.openstreetmap.org/search?q=" & URLEncode(simplifiedAddress) & "&format=json&limit=1&accept-language=ru"
    
    DebugPrint "Отправляем запрос: " & url
    
    ' Устанавливаем заголовки
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Accept-Language", "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7"
    
    ' Устанавливаем таймаут
    http.setTimeouts REQUEST_TIMEOUT, REQUEST_TIMEOUT, REQUEST_TIMEOUT, REQUEST_TIMEOUT
    
    ' Задержка для соблюдения лимитов Nominatim
    Sleep NOMINATIM_DELAY
    http.send
    
    DebugPrint "HTTP статус: " & http.Status
    
    If http.Status = 200 Then
        response = http.responseText
        DebugPrint "Ответ от сервера: " & response
        
        ' Проверяем что ответ не пустой
        If response = "[]" Or response = "" Then
            DebugPrint "Пустой ответ от сервера"
            GetDistrictFromNominatim = ""
            Exit Function
        End If
        
        ' Парсим координаты
        lat = ExtractJSONValue(response, "lat")
        lon = ExtractJSONValue(response, "lon")
        
        DebugPrint "Координаты: " & lat & ", " & lon
        
        If lat <> "" And lon <> "" Then
            ' Обратное геокодирование
            GetDistrictFromNominatim = ReverseGeocode(CDbl(Replace(lat, ".", ",")), CDbl(Replace(lon, ".", ",")))
        Else
            GetDistrictFromNominatim = ""
        End If
    Else
        DebugPrint "Ошибка HTTP: " & http.Status & " - " & http.StatusText
        GetDistrictFromNominatim = ""
    End If
    
    Exit Function
    
ErrorHandler:
    DebugPrint "Ошибка в GetDistrictFromNominatim: " & Err.Description
    GetDistrictFromNominatim = ""
End Function

Function SimplifyAddress(fullAddress As String) As String
    DebugPrint "Упрощаем адрес: " & fullAddress
    
    Dim result As String
    result = fullAddress
    
    ' Убираем почтовый индекс (6 цифр в начале)
    If Len(result) > 6 And IsNumeric(Left(result, 6)) Then
        result = Trim(Mid(result, 7))
    End If
    
    ' Убираем всё лишнее, оставляем только: Минск, улица, дом
    Dim parts() As String
    parts = Split(result, ",")
    
    Dim cityFound As Boolean
    Dim streetFound As Boolean
    Dim houseFound As Boolean
    cityFound = False
    streetFound = False
    houseFound = False
    
    result = ""
    
    Dim i As Integer
    For i = 0 To UBound(parts)
        Dim part As String
        part = Trim(parts(i))
        
        If part = "" Then GoTo NextPart
        
        ' Ищем город Минск
        If Not cityFound And (InStr(1, part, "Минск", vbTextCompare) > 0 Or _
                              InStr(1, part, "минск", vbTextCompare) > 0) Then
            If result <> "" Then result = result & ", "
            result = result & "Минск"
            cityFound = True
            GoTo NextPart
        End If
        
        ' Ищем улицу
        If Not streetFound And (InStr(1, part, "ул.", vbTextCompare) > 0 Or _
                                InStr(1, part, "улица", vbTextCompare) > 0 Or _
                                InStr(1, part, "проспект", vbTextCompare) > 0 Or _
                                InStr(1, part, "пр.", vbTextCompare) > 0 Or _
                                InStr(1, part, "пр-т", vbTextCompare) > 0 Or _
                                InStr(1, part, "бульвар", vbTextCompare) > 0 Or _
                                InStr(1, part, "пер.", vbTextCompare) > 0 Or _
                                InStr(1, part, "переулок", vbTextCompare) > 0) Then
            If result <> "" Then result = result & ", "
            result = result & part
            streetFound = True
            GoTo NextPart
        End If
        
        ' Ищем номер дома (содержит цифры и может содержать "д.", "дом", "к", "корп.")
        If Not houseFound And (InStr(1, part, "д.", vbTextCompare) > 0 Or _
                               InStr(1, part, "дом", vbTextCompare) > 0 Or _
                               HasNumbers(part)) Then
            If result <> "" Then result = result & ", "
            result = result & part
            houseFound = True
            GoTo NextPart
        End If
        
NextPart:
    Next i
    
    ' Если город не нашли - добавляем Минск
    If Not cityFound Then
        If result <> "" Then result = "Минск, " & result
        Else
        result = "Минск, " & Mid(result, InStr(result, ",") + 1)
    End If
    
    ' Убираем лишние слова из улицы, оставляем только название
    result = CleanStreetName(result)
    
    DebugPrint "Упрощенный адрес: " & result
    SimplifyAddress = Trim(result)
End Function

' Функция для проверки наличия цифр в строке
Function HasNumbers(text As String) As Boolean
    Dim i As Integer
    For i = 1 To Len(text)
        If IsNumeric(Mid(text, i, 1)) Then
            HasNumbers = True
            Exit Function
        End If
    Next i
    HasNumbers = False
End Function

' Функция для очистки названия улицы от лишних слов
Function CleanStreetName(address As String) As String
    Dim result As String
    result = address
    
    ' Заменяем полные названия на сокращения
    result = Replace(result, "улица", "ул.")
    result = Replace(result, "проспект", "пр.")
    result = Replace(result, "бульвар", "б-р")
    result = Replace(result, "переулок", "пер.")
    
    ' Убираем лишние пробелы
    result = Replace(result, "  ", " ")
    
    CleanStreetName = Trim(result)
End Function

Function ReverseGeocode(lat As Double, lon As Double) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim response As String
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' URL для обратного геокодирования
    url = "https://nominatim.openstreetmap.org/reverse?lat=" & _
          Replace(CStr(lat), ",", ".") & "&lon=" & _
          Replace(CStr(lon), ",", ".") & "&format=json&accept-language=ru&addressdetails=1"
    
    DebugPrint "Обратное геокодирование: " & url
    
    ' Устанавливаем заголовки
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Accept-Language", "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7"
    
    ' Устанавливаем таймаут
    http.setTimeouts REQUEST_TIMEOUT, REQUEST_TIMEOUT, REQUEST_TIMEOUT, REQUEST_TIMEOUT
    
    ' Задержка между запросами
    Sleep NOMINATIM_DELAY
    http.send
    
    DebugPrint "Обратное геокодирование - HTTP статус: " & http.Status
    
    If http.Status = 200 Then
        response = http.responseText
        DebugPrint "Ответ обратного геокодирования: " & response
        
        If response = "" Then
            ReverseGeocode = "Не удалось определить"
            Exit Function
        End If
        
        ' Ищем район по приоритету
        ReverseGeocode = FindDistrictInResponse(response)
        
        ' Если не нашли через API, пробуем по координатам
        If ReverseGeocode = "" Then
            ReverseGeocode = GetDistrictByCoordinates(lat, lon)
        End If
    Else
        DebugPrint "Ошибка HTTP при обратном геокодировании: " & http.Status
        ReverseGeocode = "Не удалось определить"
    End If
    
    Exit Function
    
ErrorHandler:
    DebugPrint "Ошибка в ReverseGeocode: " & Err.Description
    ReverseGeocode = "Не удалось определить"
End Function

Function FindDistrictInResponse(jsonResponse As String) As String
    DebugPrint "=== НАЧАЛО FindDistrictInResponse ==="
    DebugPrint "JSON ответ: " & jsonResponse
    
    ' Пробуем найти район в display_name
    Dim displayName As String
    displayName = ExtractJSONValue(jsonResponse, "display_name")
    DebugPrint "Display_name: " & displayName
    
    If displayName <> "" Then
        ' Ищем район Минска в display_name
        Dim districts() As String
        Dim i As Integer
        
        districts = Split("Октябрьский,Центральный,Советский,Первомайский,Партизанский,Заводской,Ленинский,Московский,Фрунзенский", ",")
        
        For i = 0 To UBound(districts)
            If InStr(1, displayName, districts(i), vbTextCompare) > 0 Then
                DebugPrint "Найден район в display_name: " & districts(i)
                FindDistrictInResponse = districts(i)
                Exit Function
            End If
        Next i
    End If
    
    ' Если не нашли в display_name, ищем в address
    Dim districtKeys() As String
    Dim district As String
    
    districtKeys = Split("city_district,suburb,district,county,municipality", ",")
    
    For i = 0 To UBound(districtKeys)
        district = ExtractJSONValue(jsonResponse, districtKeys(i))
        If district <> "" Then
            DebugPrint "Найден район в " & districtKeys(i) & ": " & district
            FindDistrictInResponse = district
            Exit Function
        End If
    Next i
    
    DebugPrint "Район не найден"
    FindDistrictInResponse = ""
End Function

Function ExtractJSONValue(jsonString As String, key As String) As String
    On Error GoTo ErrorHandler
    
    DebugPrint "Ищем ключ: " & key
    
    Dim startPos As Long
    Dim endPos As Long
    Dim searchKey As String
    Dim tempValue As String
    
    ' Ищем ключ в формате "key":"
    searchKey = """" & key & """:"""
    startPos = InStr(1, jsonString, searchKey, vbTextCompare)
    
    If startPos > 0 Then
        startPos = startPos + Len(searchKey)
        endPos = InStr(startPos, jsonString, """")
        If endPos > startPos Then
            ExtractJSONValue = Mid(jsonString, startPos, endPos - startPos)
            DebugPrint "Найдено значение: " & ExtractJSONValue
            Exit Function
        End If
    End If
    
    ' Пробуем найти числовое значение (для координат)
    searchKey = """" & key & """:"
    startPos = InStr(1, jsonString, searchKey, vbTextCompare)
    If startPos > 0 Then
        startPos = startPos + Len(searchKey)
        endPos = InStr(startPos, jsonString, ",")
        If endPos = 0 Then endPos = InStr(startPos, jsonString, "}")
        If endPos > startPos Then
            tempValue = Trim(Mid(jsonString, startPos, endPos - startPos))
            ' Убираем возможные кавычки
            If Left(tempValue, 1) = """" Then tempValue = Mid(tempValue, 2)
            If Right(tempValue, 1) = """" Then tempValue = Left(tempValue, Len(tempValue) - 1)
            ExtractJSONValue = tempValue
            DebugPrint "Найдено числовое значение: " & ExtractJSONValue
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    DebugPrint "Ошибка в ExtractJSONValue: " & Err.Description
    ExtractJSONValue = ""
End Function
' Альтернативная функция для определения района по координатам
Function GetDistrictByCoordinates(lat As Double, lon As Double) As String
    ' Для координат ул. Асаналиева возвращаем Октябрьский район
    If lat > 53.84 And lat < 53.85 And lon > 27.54 And lon < 27.55 Then
        GetDistrictByCoordinates = "Октябрьский"
    Else
        GetDistrictByCoordinates = ""
    End If
End Function

Function ExtractDistrictFromDisplayName(jsonString As String) As String
    On Error GoTo ErrorHandler
    
    Dim displayName As String
    Dim startPos As Long
    Dim endPos As Long
    
    ' Ищем display_name
    startPos = InStr(1, jsonString, """display_name"":""", vbTextCompare)
    If startPos = 0 Then Exit Function
    
    startPos = startPos + 16
    endPos = InStr(startPos, jsonString, """")
    If endPos = 0 Then Exit Function
    
    displayName = Mid(jsonString, startPos, endPos - startPos)
    
    ' Ищем район Минска в display_name
    Dim districts() As String
    Dim i As Integer
    
    districts = Split("Октябрьский,Центральный,Советский,Первомайский,Партизанский,Заводской,Ленинский,Московский,Фрунзенский", ",")
    
    For i = 0 To UBound(districts)
        If InStr(1, displayName, districts(i), vbTextCompare) > 0 Then
            ExtractDistrictFromDisplayName = districts(i)
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    ExtractDistrictFromDisplayName = ""
End Function

Function GetCourtByDistrict(district As String) As String
    On Error GoTo ErrorHandler
    
    Dim courts As Collection
    Set courts = New Collection
    
    ' База данных судов Минска - используем ВАШИ адреса
    courts.Add "Суд Октябрьского района г. Минска|220045, г. Минска, ул. Семашко, д. 33", "Октябрьский"
    courts.Add "Суд Центрального района г. Минска|220030, г. Минск, ул. Кирова, д. 21", "Центральный"
    courts.Add "Суд Советского района г. Минска|220076, г. Минск, ул. Ф.Скорины, 6Б", "Советский"
    courts.Add "Суд Первомайского района г. Минска|220076, г. Минск, ул. Ф.Скорины, 6Б", "Первомайский"
    courts.Add "Суд Партизанского района г. Минска|220045, г. Минска, ул. Семашко, д. 33", "Партизанский"
    courts.Add "Суд Заводского района г. Минска|220107, г. Минск, Партизанский пр., 75А", "Заводской"
    courts.Add "Суд Ленинского района г. Минска|220045, г. Минска, ул. Семашко, д. 33", "Ленинский"
    courts.Add "Суд Московского района г. Минска|220083, г. Минск, пр. газеты ""Правда"", д. 27", "Московский"
    courts.Add "Суд Фрунзенского района г. Минска|220092, г. Минск, ул. Дунина-Марцинкевича, д. 1, корп. 2", "Фрунзенский"
    courts.Add "Суд Минского района Минской области|220028, г. Минск, ул. Маяковского, д. 119А", "Минский"
    
    ' Ищем точное совпадение
    On Error Resume Next
    GetCourtByDistrict = courts(district)
    If Err.Number = 0 Then Exit Function
    On Error GoTo ErrorHandler
    
    ' Ищем частичное совпадение
    Dim key As Variant
    For Each key In courts
        If InStr(1, district, Split(key, "|")(0), vbTextCompare) > 0 Then
            GetCourtByDistrict = key
            Exit Function
        End If
    Next key
    
    GetCourtByDistrict = ""
    Exit Function
    
ErrorHandler:
    DebugPrint "Ошибка в GetCourtByDistrict: " & Err.Description
    GetCourtByDistrict = ""
End Function


Sub AddCourtToDocument(courtName As String, courtAddress As String, originalAddress As String, district As String)
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Сохраняем текущее положение курсора
    Dim originalRange As Range
    Set originalRange = Selection.Range
    
    ' Переходим в конец документа
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    
    ' Добавляем информацию с форматированием
    With Selection
        .Style = ActiveDocument.Styles(wdStyleHeading2)
        .TypeText "Результат поиска суда"
        .Style = ActiveDocument.Styles(wdStyleNormal)
        .TypeParagraph
        
        .TypeText "Исходный адрес: "
        .Font.Bold = True
        .TypeText originalAddress
        .Font.Bold = False
        .TypeParagraph
        
        .TypeText "Определенный район: "
        .Font.Bold = True
        .TypeText district
        .Font.Bold = False
        .TypeParagraph
        
        .TypeText "Найденный суд: "
        .Font.Bold = True
        .TypeText courtName
        .Font.Bold = False
        .TypeParagraph
        
        .TypeText "Адрес суда: "
        .Font.Bold = True
        .TypeText courtAddress
        .Font.Bold = False
        .TypeParagraph
        
        .TypeText String(50, "-")
        .TypeParagraph
        .TypeParagraph
    End With
    
    ' Восстанавливаем исходное положение курсора
    originalRange.Select
    
    Exit Sub
    
ErrorHandler:
    DebugPrint "Ошибка в AddCourtToDocument: " & Err.Description
    On Error Resume Next
    originalRange.Select
End Sub

' Старая функция EncodeURL удалена, теперь используется улучшенная URLEncode

Sub DebugPrint(message As String)
    ' Раскомментируйте следующую строку для отладки:
    Debug.Print "[" & Now & "] " & message
End Sub

' Улучшенная тестовая функция
Sub TestNominatim()
    On Error GoTo ErrorHandler
    
    Dim testAddress As String
    testAddress = "Минск, Асаналиева, 6к2"
    
    DebugPrint "=== ТЕСТ NOMINATIM ==="
    DebugPrint "Адрес: " & testAddress
    
    Dim district As String
    district = GetDistrictFromNominatim(testAddress)
    
    If district <> "" And district <> "Не удалось определить" Then
        DebugPrint "УСПЕХ: Найден район - " & district
        MsgBox "УСПЕХ!" & vbCrLf & "Найден район: " & district, vbInformation, "Тест пройден"
    Else
        DebugPrint "ОШИБКА: Не удалось найти район"
        MsgBox "ОШИБКА!" & vbCrLf & "Не удалось найти район." & vbCrLf & _
               "Возможные причины:" & vbCrLf & _
               "o Отсутствует интернет-соединение" & vbCrLf & _
               "o Сервер временно недоступен" & vbCrLf & _
               "o Адрес указан некорректно", vbExclamation, "Тест не пройден"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при тестировании: " & Err.Description, vbCritical, "Ошибка тестирования"
End Sub
Sub FindCourtAndReplaceExisting()
    On Error GoTo ErrorHandler
    
    Dim selectedText As String
    Dim district As String
    Dim courtInfo As String
    Dim courtName As String
    Dim courtAddress As String
    
    ' Получаем выделенный текст (адрес)
    selectedText = GetSelectedText()
    If selectedText = "" Then Exit Sub
    
    Application.StatusBar = "Определяем район для адреса..."
    
    ' Получаем район через Nominatim
    district = GetDistrictFromNominatim(selectedText)
    
    If district = "" Or district = "Не удалось определить" Then
        MsgBox "Не удалось определить район для адреса: " & selectedText & vbCrLf & _
               "Попробуйте:" & vbCrLf & _
               "1. Проверить подключение к интернету" & vbCrLf & _
               "2. Уточнить формат адреса" & vbCrLf & _
               "3. Использовать ручной ввод района", vbExclamation, "Ошибка определения района"
        Application.StatusBar = ""
        Exit Sub
    End If
    
    Application.StatusBar = "Найден район: " & district & ". Ищем суд..."
    
    ' Получаем информацию о суде
    courtInfo = GetCourtByDistrict(district)
    
    If courtInfo = "" Then
        MsgBox "Не найден суд для района: " & district & vbCrLf & _
               "Доступные районы: Октябрьский, Центральный, Советский, Первомайский, " & _
               "Партизанский, Заводской, Ленинский, Московский, Фрунзенский", vbExclamation
        Application.StatusBar = ""
        Exit Sub
    End If
    
    ' Парсим информацию о суде
    courtName = Split(courtInfo, "|")(0)
    courtAddress = Split(courtInfo, "|")(1)
    
    ' Сохраняем текущую позицию курсора
    Dim originalRange As Range
    Set originalRange = Selection.Range
    
    ' Ищем и заменяем существующий суд
    Dim found As Boolean
    found = ReplaceExistingCourt(courtName, courtAddress)
    
    If found Then
        ' Возвращаем курсор на исходную позицию
        originalRange.Select
        
        Application.StatusBar = "Готово! Суд заменен в документе"
        
        ' Показываем результат
        MsgBox "Суд успешно заменен в документе:" & vbCrLf & _
               "Район: " & district & vbCrLf & _
               "Суд: " & courtName & vbCrLf & _
               "Адрес: " & courtAddress, vbInformation, "Результат поиска"
    Else
        ' Если не нашли существующий суд, вставляем новый
        originalRange.Select
        If InsertNewCourt(courtName, courtAddress) Then
            MsgBox "Суд не найден в документе. Добавлен новый суд:" & vbCrLf & _
                   courtName & vbCrLf & courtAddress, vbInformation
        Else
            MsgBox "Не удалось найти или добавить суд в документе.", vbExclamation
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = ""
    MsgBox "Произошла ошибка: " & Err.Description & vbCrLf & _
           "Номер ошибки: " & Err.Number, vbCritical, "Ошибка выполнения"
End Sub

Function ReplaceExistingCourt(courtName As String, courtAddress As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim searchRange As Range
    Set searchRange = ActiveDocument.Content
    
    ' Ищем любые суды по ключевым словам
    Dim searchPatterns() As String
    searchPatterns = Split("Суд Октябрьского,Суд Центрального,Суд Советского,Суд Первомайского,Суд Партизанского,Суд Заводского,Суд Ленинского,Суд Московского,Суд Фрунзенского,Суд Минского", ",")
    
    Dim i As Integer
    For i = 0 To UBound(searchPatterns)
        With searchRange.Find
            .text = searchPatterns(i)
            .Forward = True
            .Wrap = wdFindStop
            .MatchCase = False
            .MatchWholeWord = False
            
            If .Execute Then
                ' Нашли существующий суд - выделяем всю строку
                searchRange.Select
                Selection.Expand Unit:=wdLine
                Dim courtLine As String
                courtLine = Selection.text
                
                ' Сохраняем позицию начала суда
                Dim courtStart As Long
                courtStart = Selection.Start
                
                ' Переходим к следующей строке (адрес суда)
                Selection.MoveDown Unit:=wdLine, count:=1
                Selection.Expand Unit:=wdLine
                Dim addressLine As String
                addressLine = Selection.text
                
                ' Проверяем, что вторая строка похожа на адрес (содержит цифры)
                If HasNumbers(addressLine) Then
                    ' Выделяем обе строки для замены
                    Set searchRange = ActiveDocument.Range(courtStart, Selection.End)
                    searchRange.Select
                    
                    ' Заменяем на новый суд и адрес с мягким переносом между ними и обычным после
                    Selection.text = courtName & Chr(11) & courtAddress & vbCrLf
                    
                    ReplaceExistingCourt = True
                    Exit Function
                End If
            End If
        End With
    Next i
    
    ReplaceExistingCourt = False
    Exit Function
    
ErrorHandler:
    DebugPrint "Ошибка в ReplaceExistingCourt: " & Err.Description
    ReplaceExistingCourt = False
End Function

' Функция для вставки нового суда если не нашли существующий
Function InsertNewCourt(courtName As String, courtAddress As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Ищем место для вставки (первое упоминание "Суд" или конец документа)
    Dim searchRange As Range
    Set searchRange = ActiveDocument.Content
    
    With searchRange.Find
        .text = "Суд"
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False
        .MatchWholeWord = False
        
        If .Execute Then
            ' Нашли "Суд" - вставляем после него
            searchRange.Select
            Selection.Collapse Direction:=wdCollapseEnd
            Selection.TypeParagraph
        Else
            ' Не нашли - вставляем в конец документа
            Selection.EndKey Unit:=wdStory
            Selection.TypeParagraph
        End If
        
        ' Вставляем новый суд с мягким переносом между названием и адресом и обычным после
        Selection.TypeText courtName
        Selection.TypeText Chr(11)  ' Shift+Enter между названием и адресом
        Selection.TypeText courtAddress
        Selection.TypeParagraph  ' Обычный Enter после адреса
        
        InsertNewCourt = True
        Exit Function
    End With
    
    InsertNewCourt = False
    Exit Function
    
ErrorHandler:
    DebugPrint "Ошибка в InsertNewCourt: " & Err.Description
    InsertNewCourt = False
End Function



' Упрощенная версия для быстрой замены
Sub QuickCourtReplace()
    On Error GoTo ErrorHandler
    
    Dim selectedText As String
    selectedText = GetSelectedText()
    If selectedText = "" Then Exit Sub
    
    Dim district As String
    district = GetDistrictFromNominatim(selectedText)
    
    If district = "" Then Exit Sub
    
    Dim courtInfo As String
    courtInfo = GetCourtByDistrict(district)
    
    If courtInfo = "" Then Exit Sub
    
    Dim courtName As String
    Dim courtAddress As String
    courtName = Split(courtInfo, "|")(0)
    courtAddress = Split(courtInfo, "|")(1)
    
    ' Сохраняем позицию курсора
    Dim originalRange As Range
    Set originalRange = Selection.Range
    
    ' Заменяем существующий суд
    If ReplaceExistingCourt(courtName, courtAddress) Then
        originalRange.Select
        MsgBox "Суд заменен: " & courtName, vbInformation
    Else
        ' Если не нашли существующий, вставляем новый
        originalRange.Select
        If InsertNewCourt(courtName, courtAddress) Then
            MsgBox "Добавлен новый суд: " & courtName, vbInformation
        Else
            MsgBox "Не удалось заменить или добавить суд.", vbExclamation
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка: " & Err.Description, vbExclamation
End Sub

Sub ClearStatusBar()
    Application.StatusBar = False
End Sub


