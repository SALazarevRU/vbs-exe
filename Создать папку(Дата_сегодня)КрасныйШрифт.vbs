On Error Resume Next

Set FSO = CreateObject("Scripting.FileSystemObject")

' Имя папки: ДД.ММ.ГГГГ
MyDate = Right("0" & Day(Date), 2) & "." & Right("0" & Month(Date), 2) & "." & Year(Date)

' Путь к папке со скриптом
MyPath = FSO.GetParentFolderName(WScript.ScriptFullName)

' Полный путь к новой папке
NewFold = MyPath & "\" & MyDate

' Проверяем, существует ли папка
If Not FSO.FolderExists(NewFold) Then
    ' Создаём папку
    FSO.CreateFolder(NewFold)
    
    ' Запускаем HTA-окно об успехе (без пути)
    LaunchHTA "Папка создана!", "Успех"
Else
    ' Запускаем HTA-окно о том, что папка уже есть (без пути)
    LaunchHTA "Папка уже существует!", "Внимание"
End If

' Удаляем текущий скрипт
FSO.DeleteFile WScript.ScriptFullName, True

' Освобождаем объект
Set FSO = Nothing
WScript.Quit 0

' Функция для запуска HTA-окна
Sub LaunchHTA(message, title)
    Dim tempHTA, fso, ts
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Создаём временный HTA-файл
    tempHTA = fso.GetTempName & ".hta"
    tempHTA = fso.GetSpecialFolder(2) & "\" & tempHTA  ' TEMP
    
    Set ts = fso.CreateTextFile(tempHTA, True)
    
    ts.Write GetHTAContent(message, title)
    ts.Close
    
    ' Запускаем HTA
    CreateObject("WScript.Shell").Run tempHTA, 1, False
    
    ' Можно удалить HTA после показа, но лучше оставить — Windows сама почистит TEMP
End Sub

' Возвращает HTML/HTA-код окна
Function GetHTAContent(msg, ttl)
    GetHTAContent = "<html>" & vbCrLf & _
    "<head>" & vbCrLf & _
    "  <title>" & ttl & "</title>" & vbCrLf & _
    "  <HTA:APPLICATION" & vbCrLf & _
    "    APPLICATIONNAME=""Folder Notification""" & vbCrLf & _
    "    BORDER=""thin""" & vbCrLf & _
    "    CAPTION=""yes""" & vbCrLf & _
    "    INNERBORDER=""yes""" & vbCrLf & _
    "    SCROLL=""no""" & vbCrLf & _
    "    SINGLEINSTANCE=""yes""" & vbCrLf & _
    "    WINDOWSTATE=""normal""" & vbCrLf & _
    "    SYSMENU=""yes""" & vbCrLf & _
    "    MAXIMIZEBOX=""no""" & vbCrLf & _
    "    MINIMIZEBOX=""no""" & vbCrLf & _
    "    RESIZE=""no""" & vbCrLf & _
    "    WINDOWWIDTH=600" & vbCrLf & _
    "    WINDOWHEIGHT=400" & vbCrLf & _
    "    SHOWINTASKBAR=""no""" & vbCrLf & _
    "  >" & vbCrLf & _
    "  <style>" & vbCrLf & _
    "    body { margin: 0; padding: 0; width: 600px; height: 400px;" & _
    "      background-image: url('C:\\Users\\Хозяин\\Desktop\\vbs-exe-main\\01.01.2026\\vbs-exe-main\\more3.gif'); display: flex; justify-content: center; align-items: center; }" & vbCrLf & _
    "    .container { background-color: rgba(255,255,255,0.9); border-radius: 12px;" & _
    "      padding: 30px 40px; text-align: center; max-width: 80%; }" & vbCrLf & _
    "    .message { font-size: 20px; color: #FF0000; margin: 0 0 20px 0; }" & vbCrLf & _
    "    button { padding: 12px 24px; background-color: #0078d4; color: white;" & _
    "      border: none; border-radius: 6px; cursor: pointer; font-size: 16px;" & _
    "      font-weight: bold; transition: background-color 0.3s; }" & vbCrLf & _
    "    button:hover { background-color: #005a9e; }" & vbCrLf & _
    "  </style>" & vbCrLf & _
    "  <script>" & vbCrLf & _
    "    window.resizeTo(900, 643);" & vbCrLf & _
    "    window.moveTo((screen.width-900)/2, (screen.height-643)/2);" & vbCrLf & _
    "    // setTimeout(function() { window.close(); }, 3000);" & vbCrLf & _
    "  </script>" & vbCrLf & _
    "</head>" & vbCrLf & _
    "<body>" & vbCrLf & _
    "  <div class='container'>" & vbCrLf & _
    "    <div class='message'>" & msg & "</div>" & vbCrLf & _
    "  </div>" & vbCrLf & _
    "</body>" & vbCrLf & _
    "</html>"
End Function
