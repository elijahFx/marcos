' API declarations для разных версий Office
#If VBA7 Then
    ' Для Office 2010 и новее
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongLong)
#Else
    ' Для Office 2007 и старше
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Константы для настроек
Private Const NOMINATIM_DELAY As Long = 1500 ' Задержка между запросами в мс
Private Const REQUEST_TIMEOUT As Long = 10000 ' Таймаут запроса в мс

Sub AddPostalIndexToAddress()
    On Error GoTo ErrorHandler
    
    Dim selectedText As String
    Dim postalIndex As String
    Dim originalRange As Range
    
    ' Получаем выделенный текст
    selectedText = GetSelectedText()
    If selectedText = "" Then Exit Sub
    
    Application.StatusBar = "Определяем почтовый индекс для адреса..."
    
    ' Получаем почтовый индекс через Nominatim
    postalIndex = GetPostalIndexFromNominatim(selectedText)
    
    If postalIndex = "" Or postalIndex = "Не удалось определить" Then
        MsgBox "Не удалось определить почтовый индекс для адреса: " & selectedText & vbCrLf & _
               "Попробуйте:" & vbCrLf & _
               "1. Проверить подключение к интернету" & vbCrLf & _
               "2. Уточнить формат адреса" & vbCrLf & _
               "3. Использовать ручной ввод индекса", vbExclamation, "Ошибка определения индекса"
        Application.StatusBar = ""
        Exit Sub
    End If
    
    ' Сохраняем позицию выделения
    Set originalRange = Selection.Range
    
    ' Заменяем выделенный текст на индекс + запятая + адрес
    Selection.text = postalIndex & ", " & selectedText
    
    ' Восстанавливаем выделение на новый текст
    originalRange.Select
    
    Application.StatusBar = "Готово! Почтовый индекс добавлен"
    
    ' Показываем результат
    MsgBox "Почтовый индекс успешно добавлен:" & vbCrLf & _
           "Адрес: " & selectedText & vbCrLf & _
           "Индекс: " & postalIndex, vbInformation, "Результат"
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = ""
    MsgBox "Произошла ошибка: " & Err.Description & vbCrLf & _
           "Номер ошибки: " & Err.Number, vbCritical, "Ошибка выполнения"
End Sub

Function GetPostalIndexFromNominatim(address As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim response As String
    
    ' Упрощаем адрес для лучшего поиска
    Dim simplifiedAddress As String
    simplifiedAddress = SimplifyAddress(address)
    DebugPrint "Упрощенный адрес для поиска индекса: " & simplifiedAddress
    
    ' Создаем HTTP объект
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Кодируем адрес для URL
    url = "https://nominatim.openstreetmap.org/search?q=" & URLEncode(simplifiedAddress) & "&format=json&limit=1&accept-language=ru&addressdetails=1"
    
    DebugPrint "Отправляем запрос для индекса: " & url
    
    ' Устанавливаем заголовки
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Accept-Language", "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7"
    
    ' Устанавливаем таймаут
    http.setTimeouts REQUEST_TIMEOUT, REQUEST_TIMEOUT, REQUEST_TIMEOUT, REQUEST_TIMEOUT
    
    ' Задержка для соблюдения лимитов Nominatim
    Call Sleep(NOMINATIM_DELAY)
    http.send
    
    DebugPrint "HTTP статус для индекса: " & http.Status
    
    If http.Status = 200 Then
        response = http.responseText
        DebugPrint "Ответ от сервера для индекса: " & response
        
        ' Проверяем что ответ не пустой
        If response = "[]" Or response = "" Then
            DebugPrint "Пустой ответ от сервера для индекса"
            GetPostalIndexFromNominatim = ""
            Exit Function
        End If
        
        ' Ищем почтовый индекс в ответе
        GetPostalIndexFromNominatim = ExtractPostalCodeFromResponse(response)
        
    Else
        DebugPrint "Ошибка HTTP для индекса: " & http.Status & " - " & http.StatusText
        GetPostalIndexFromNominatim = ""
    End If
    
    Exit Function
    
ErrorHandler:
    DebugPrint "Ошибка в GetPostalIndexFromNominatim: " & Err.Description
    GetPostalIndexFromNominatim = ""
End Function

Function ExtractPostalCodeFromResponse(jsonResponse As String) As String
    DebugPrint "=== НАЧАЛО ExtractPostalCodeFromResponse ==="
    
    ' Ищем display_name где обычно содержится полный адрес с индексом
    Dim displayName As String
    displayName = ExtractJSONValue(jsonResponse, "display_name")
    DebugPrint "Display_name для поиска индекса: " & displayName
    
    If displayName <> "" Then
        ' Простой поиск 6 цифр подряд в display_name
        Dim i As Long
        For i = 1 To Len(displayName) - 5
            Dim possibleIndex As String
            possibleIndex = Mid(displayName, i, 6)
            If IsNumeric(possibleIndex) Then
                ' Проверяем что это действительно индекс (не часть номера дома и т.д.)
                If IsValidPostalIndex(possibleIndex, displayName, i) Then
                    ExtractPostalCodeFromResponse = possibleIndex
                    DebugPrint "Найден индекс в display_name: " & ExtractPostalCodeFromResponse
                    Exit Function
                End If
            End If
        Next i
    End If
    
    ' Если не нашли в display_name, пробуем поискать в полном ответе
    For i = 1 To Len(jsonResponse) - 5
        Dim possibleIndexInJson As String
        possibleIndexInJson = Mid(jsonResponse, i, 6)
        If IsNumeric(possibleIndexInJson) Then
            ' Проверяем контекст - должен быть почтовым индексом, а не случайными цифрами
            If IsValidPostalIndex(possibleIndexInJson, jsonResponse, i) Then
                ExtractPostalCodeFromResponse = possibleIndexInJson
                DebugPrint "Найден индекс в полном JSON: " & ExtractPostalCodeFromResponse
                Exit Function
            End If
        End If
    Next i
    
    DebugPrint "Индекс не найден в ответе"
    ExtractPostalCodeFromResponse = ""
End Function

Function IsValidPostalIndex(possibleIndex As String, context As String, position As Long) As Boolean
    ' Проверяем что найденные 6 цифр - это почтовый индекс, а не что-то другое
    
    ' Почтовый индекс обычно находится в начале адреса или после слова "индекс"
    Dim beforeChar As String
    Dim afterChar As String
    
    ' Получаем символы до и после предполагаемого индекса
    If position > 1 Then
        beforeChar = Mid(context, position - 1, 1)
    Else
        beforeChar = ""
    End If
    
    If position + 6 <= Len(context) Then
        afterChar = Mid(context, position + 6, 1)
    Else
        afterChar = ""
    End If
    
    ' Индекс обычно окружен не-цифровыми символами или находится в начале/конце строки
    If beforeChar <> "" And IsNumeric(beforeChar) Then
        IsValidPostalIndex = False ' Часть более длинного числа
        Exit Function
    End If
    
    If afterChar <> "" And IsNumeric(afterChar) Then
        IsValidPostalIndex = False ' Часть более длинного числа
        Exit Function
    End If
    
    ' Проверяем типичные контексты для почтового индекса
    Dim lowerContext As String
    lowerContext = LCase(context)
    
    ' Если индекс находится рядом с ключевыми словами - это хороший признак
    If InStr(1, lowerContext, "индекс", vbTextCompare) > 0 Or _
       InStr(1, lowerContext, "postcode", vbTextCompare) > 0 Or _
       InStr(1, lowerContext, "postal", vbTextCompare) > 0 Then
        IsValidPostalIndex = True
        Exit Function
    End If
    
    ' Или если индекс находится в начале адреса (типично для Беларуси)
    If position < 10 Then
        IsValidPostalIndex = True
        Exit Function
    End If
    
    ' По умолчанию считаем валидным если окружен не-цифровыми символами
    IsValidPostalIndex = True
End Function

' Упрощенная версия для быстрого добавления индекса
Sub QuickAddPostalIndex()
    On Error GoTo ErrorHandler
    
    Dim selectedText As String
    selectedText = GetSelectedText()
    If selectedText = "" Then Exit Sub
    
    Dim postalIndex As String
    postalIndex = GetPostalIndexFromNominatim(selectedText)
    
    If postalIndex = "" Or postalIndex = "Не удалось определить" Then
        ' Предлагаем ручной ввод если автоматический не сработал
        If MsgBox("Не удалось автоматически определить индекс." & vbCrLf & _
                 "Хотите ввести индекс вручную?", vbQuestion + vbYesNo, "Ручной ввод") = vbYes Then
            postalIndex = InputBox("Введите почтовый индекс (6 цифр):", "Ручной ввод индекса")
            If postalIndex = "" Then Exit Sub
            If Len(postalIndex) <> 6 Or Not IsNumeric(postalIndex) Then
                MsgBox "Индекс должен состоять из 6 цифр!", vbExclamation
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    
    ' Заменяем выделенный текст
    Selection.text = postalIndex & ", " & selectedText
    
    MsgBox "Индекс добавлен: " & postalIndex, vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка: " & Err.Description, vbExclamation
End Sub

' Функция для ручного добавления индекса
Sub AddPostalIndexManually()
    On Error GoTo ErrorHandler
    
    Dim selectedText As String
    selectedText = GetSelectedText()
    If selectedText = "" Then Exit Sub
    
    Dim postalIndex As String
    postalIndex = InputBox("Введите почтовый индекс (6 цифр) для адреса:" & vbCrLf & selectedText, "Ручной ввод индекса")
    
    If postalIndex = "" Then Exit Sub
    
    ' Проверяем формат индекса
    If Len(postalIndex) <> 6 Or Not IsNumeric(postalIndex) Then
        MsgBox "Индекс должен состоять из 6 цифр!", vbExclamation
        Exit Sub
    End If
    
    ' Заменяем выделенный текст
    Selection.text = postalIndex & ", " & selectedText
    
    MsgBox "Индекс " & postalIndex & " успешно добавлен к адресу", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка: " & Err.Description, vbExclamation
End Sub

' Альтернативная функция задержки если Sleep не работает
Sub Delay(milliseconds As Long)
    Dim start As Double
    start = Timer
    Do While Timer < start + milliseconds / 1000
        DoEvents
    Loop
End Sub

