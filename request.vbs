' ///////////////////////////////////////////////
' R.Rabinovich 2019 (roman@web-line.ru)
'
' скрипт на вход ожидает следующие параметры:
' -i=[Integer] - ИНН или ОГРН (целое число)
' -f=[String] - путь к папке, куда будет сохраняться файл (по умолчанию - в ту же, где и скрипт)
' -t=[String] - токен (API-ключ) с сайта dadata.ru (длина 40 символов)
' //////////////////////////////////////////////

' LibraryInclude
Dim goFS     : Set goFS = CreateObject("Scripting.FileSystemObject")
Dim gsLibDir : gsLibDir = ".\"
ExecuteGlobal goFS.OpenTextFile(goFS.BuildPath(gsLibDir, "jsonParser.vbs")).ReadAll()


' глобальные переменные
Dim tokenLength: tokenLength = 40
Dim token                                                       ' токен на сайте dadata.ru
Dim argumentsParseError                                         ' текст ошибки парсинга запроса
Dim requestedValue                                              ' запрашиваемый ИНН или ОГРН
Dim destinationFolder : destinationFolder = ""                  ' папка, в которую будет сохраняться результат

' URL, к которому будем обращаться с POST-запросом
Dim url : url = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party"

If CheckArguments = True Then
   HandleRequest()
Else
   WScript.Echo argumentsParseError
End If


' ф-ция запрашивает данные ИНН на заданном сайте
Function HandleRequest()
    Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")

    Dim postData
    postData = "{""query"": """ + requestedValue  + """, ""branch_type"" : ""MAIN"" }"

    oXMLHTTP.Open "POST", url, False
    oXMLHTTP.setRequestHeader "Content-Type", "application/json"
    oXMLHTTP.setRequestHeader "Accept", "application/json"
    oXMLHTTP.setRequestHeader "Authorization", "Token " + token
    oXMLHTTP.SetRequestHeader "Content-Length", CStr(Len(PostData))
    oXMLHTTP.Send postData

    Do While oXMLHTTP.readystate <> 4: WScript.Sleep 200: Loop

    If oXMLHTTP.Status <> 200 Then
        WScript.Echo "Can't fetch data from " + url + " (status: " + CStr(oXMLHTTP.Status) + ")"
        Exit Function
    End if

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' нет папки, в которую будем писать, тогда создадим
    If len(destinationFolder) <> 0 AND fso.FolderExists(destinationFolder) = False Then
         fso.CreateFolder(destinationFolder)
    End If

    ' пути к 2 сохраняемым файлам: оригинал в json, результат парсинга - в txt
    Dim jsonFilePath: jsonFilePath = destinationFolder + "/" + requestedValue + ".json"
    Dim txtFilePath : txtFilePath = destinationFolder + "/" + requestedValue + ".txt"

    ' бьем старые файлы, если имеются
    If fso.FileExists(jsonFilePath) = True Then
        fso.DeleteFile(jsonFilePath)
    End If

    ' бьем старые файлы, если имеются
    If fso.FileExists(txtFilePath) = True Then
        fso.DeleteFile(txtFilePath)
    End If

    ' для будущих поколений сохраним оригинальный JSON, пришедший с сервера
    Set jsonFile = fso.CreateTextFile(jsonFilePath, True)
    jsonFile.Write oXMLHTTP.responseText
    jsonFile.Close

    ' начинаем парсить JSON для сохранения в TXT
    Dim json : Set json = New JSON
    json.LoadJSON(oXMLHTTP.responseText)

    ' ассоциативный массив, заполняемый в ходе парсинга данными организации
    Set orgData = CreateObject("Scripting.Dictionary")

    'Наименование компании одной строкой (как показывается в списке подсказок)
    orgData.add "value", json.GetProp("suggestions.value", false, 0)

    'Наименование компании одной строкой (полное)
    orgData.add "unrestricted_value", json.GetProp("suggestions.unrestricted_value", false, 0)

    ' Адрес.
    ' - адрес организации для юридических лиц;
    ' - город проживания для индивидуальных предпринимателей.

    'Адрес одной строкой
    ' стандартизован, поэтому может отличаться от записанного в ЕГРЮЛ.
    orgData.add "data.address.value", json.GetProp("suggestions.data.address.value", false, 0)

    'Адрес одной строкой (полный, от региона)
    'стандартизован, поэтому может отличаться от записанного в ЕГРЮЛ.
    orgData.add "data.address.unrestricted_value", json.GetProp("suggestions.data.address.unrestricted_value", false, 0)

    'гранулярный адрес. Может отсутствовать
    orgData.add "data.address.data", json.GetProp("suggestions.data.address.data", false, 0)

    'адрес одной строкой как в ЕГРЮЛ
    orgData.add "data.address.data.source", json.GetProp("suggestions.data.address.data.source", false, 0)

    'Количество филиалов
    orgData.add "data.branch_count", json.GetProp("suggestions.data.branch_count", false, 0)

    'Тип подразделения
    'MAIN   — головная организация
    'BRANCH — филиал
    orgData.add "data.branch_type", json.GetProp("suggestions.data.branch_type", false, 0)

    'ИНН
    orgData.add "data.inn", json.GetProp("suggestions.data.inn", false, 0)

    'КПП
    orgData.add "data.kpp", json.GetProp("suggestions.data.kpp", false, 0)

    'ОГРН
    orgData.add "data.ogrn", json.GetProp("suggestions.data.ogrn", false, 0)

    'Дата выдачи ОГРН
    orgData.add "data.ogrn_date", json.GetProp("suggestions.data.ogrn_date", false, 0)

    'Уникальный идентификатор в Дадате
    orgData.add "data.hid", json.GetProp("suggestions.data.hid", false, 0)

    'ФИО руководителя
    orgData.add "data.management.name", json.GetProp("suggestions.data.management.name", false, 0)

    'должность руководителя
    orgData.add "data.management.post", json.GetProp("suggestions.data.management.post", false, 0)

    'полное наименование с ОПФ
    orgData.add "full_with_opf", json.GetProp("full_with_opf", false, 0)

    'краткое наименование с ОПФ
    orgData.add "short_with_opf", json.GetProp("short_with_opf", false, 0)

    'наименование на латинице (не заполняется)
    orgData.add "latin", json.GetProp("latin", false, 0)

    'полное наименование
    orgData.add "full", json.GetProp("suggestions.data.name.full", false, 0)

    'краткое наименование
    orgData.add "short", json.GetProp("short", false, 0)

    'Код ОКПО (не заполняется)
    orgData.add "data.okpo", json.GetProp("suggestions.data.okpo", false, 0)

    'Код ОКВЭД
    orgData.add "data.okved", json.GetProp("suggestions.data.okved", false, 0)

    'Версия справочника ОКВЭД (2001 или 2014)
    orgData.add "data.okved_type", json.GetProp("suggestions.data.okved_type", false, 0)

    'ОКОПФ
    orgData.add "data.opf.code", json.GetProp("suggestions.data.opf.code", false, 0)

    'полное название ОПФ
    orgData.add "data.opf.full", json.GetProp("suggestions.data.opf.full", false, 0)

    'краткое название ОПФ
    orgData.add "data.opf.short", json.GetProp("suggestions.data.opf.short", false, 0)

    'версия справочника ОКВЭД (2001 или 2014)
    orgData.add "data.opf.type", json.GetProp("suggestions.data.opf.type", false, 0)

    'дата актуальности сведений
    orgData.add "data.state.actuality_date", json.GetProp("suggestions.data.state.actuality_date", false, 0)

    'дата регистрации
    orgData.add "data.state.registration_date", json.GetProp("suggestions.data.state.registration_date", false, 0)

    'дата ликвидации
    orgData.add "data.state.liquidation_date", json.GetProp("suggestions.data.state.liquidation_date", false, 0)

    'статус организации
    'ACTIVE       — действующая
    'LIQUIDATING  — ликвидируется
    'LIQUIDATED   — ликвидирована
    'REORGANIZING — в процессе присоединения к другому юрлицу,
    '             с последующей ликвидацией
    orgData.add "data.state.status", json.GetProp("suggestions.data.state.status", false, 0)

    'Тип организации
    'LEGAL      — юридическое лицо
    'INDIVIDUAL — индивидуальный предприниматель
    orgData.add "data.type", json.GetProp("suggestions.data.type", false, 0)

    ' собственно сваливаем все, что насобирали, в файл
    Dim textToFile
    For Each key In orgData
        textToFile = textToFile + (key + "=" + orgData.Item(key) & vbCrLf)
    Next

    Set txtFile = fso.CreateTextFile(txtFilePath, True)
    txtFile.Write textToFile
    txtFile.Close

End Function

' ф-ция проверяет наличие и корректность параметров запроса
Function CheckArguments()
    Dim argsCount : argsCount = WScript.Arguments.count

    If argsCount < 1 Then
        CheckArguments = False
        argumentsParseError = "Minimum arguments number should be more or equal to 1"
        Exit Function
    End If

    For Each arg In WScript.Arguments
        If Left(arg, 3) = "-i=" Then
            requestedValue = Mid(arg, 4)
        End If

        If Left(arg, 3) = "-t=" Then
            token = Mid(arg, 4)
        End If

        If Left(arg, 3) = "-f=" Then
            destinationFolder = Mid(arg, 4)
        End If
    Next

    If requestedValue = "" Then
        argumentsParseError = "Neither INN, nor OGRN were requested"
        CheckArguments = False
        Exit Function
    End If

    If token = "" Then
        argumentsParseError = "You should specify dadata.ru token to perform requests"
        CheckArguments = False
        Exit Function
    End if

    If Len(token) <> tokenLength Then
        argumentsParseError = "Dadata.ru token must be " + tokenLength + " characters long"
        CheckArguments = False
        Exit Function
    End If

    If IsNumeric(requestedValue) = False Then
        argumentsParseError = "INN or OGRN should be positive integer number"
        CheckArguments = False
        Exit Function
    End If

    If requestedValue <= 0 Then
        argumentsParseError = "INN or OGRN should be positive integer number"
        CheckArguments = False
        Exit Function
    End If

    If CheckDestinationFolder(destinationFolder) = False Then
        argumentsParseError = "Wrong destination folder"
        CheckArguments = False
        Exit Function
    End If

    CheckArguments = True

End Function


' ф-ция проверяет наличие папки назначения и пытается ее создать при отсутствии
Function CheckDestinationFolder(destinationFolder)

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' нет папки, в которую будем писать, тогда создадим
    If len(destinationFolder) <> 0 AND fso.FolderExists(destinationFolder) = False Then
        fso.CreateFolder(destinationFolder)

        ' папка не создана
        If fso.FolderExists(destinationFolder) = False Then
            CheckDestinationFolder = False
            Exit Function
        End If
    End If

    CheckDestinationFolder = True

End Function
