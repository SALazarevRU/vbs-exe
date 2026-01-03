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
    
    ' Запускаем HTA-окно об успехе
    LaunchHTA "Папка создана: " , "Успех"
Else
    ' Запускаем HTA-окно о том, что папка уже есть
    LaunchHTA "Папка " & " уже существует!", "Внимание"
End If


' Удаляем текущий скрипт
FSO.DeleteFile WScript.ScriptFullName, True


' Освобождаем объект
Set FSO = Nothing
WScript.Quit 0

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
    "  <!-- Принудительный режим IE8 для стабильности HTA -->" & vbCrLf & _
    "  <meta http-equiv='X-UA-Compatible' content='IE=8'>" & vbCrLf & _
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
    "    WINDOWWIDTH=900" & vbCrLf & _
    "    WINDOWHEIGHT=643" & vbCrLf & _
    "    SHOWINTASKBAR=""no""" & vbCrLf & _
    "  >" & vbCrLf & _
    "  <style>" & vbCrLf & _
    "    body { margin: 0; padding: 0; width: 900px; height: 643px;" & _
    "      background-image: url('C:\\Users\\Хозяин\\Desktop\\vbs-exe-main\\01.01.2026\\vbs-exe-main\\more3.gif');" & _
    "      display: flex; justify-content: center; align-items: center; }" & vbCrLf & _
    "    .container { padding: 40px; width: 100%; max-width: 900px; }" & vbCrLf & _
    "    .message { font-size: 22px; color: #0000FF; line-height: 1.5;" & _
    "      margin-top: 0; margin-bottom: 170px; padding-top: 0; padding-bottom: 0;" & _
    "      width: 100%;" & _
    "      text-align: center;" & _
    "      display: block;" & _
    "      /* Начальное состояние: прозрачно */" & vbCrLf & _
    "      filter: alpha(opacity=0);" & _
    "      -ms-filter: 'progid:DXImageTransform.Microsoft.Alpha(Opacity=0)';" & _
    "      opacity: 0;" & _
    "    }" & vbCrLf & _
    "  </style>" & vbCrLf & _
    "  <script>" & vbCrLf & _
    "    window.resizeTo(900, 643);" & vbCrLf & _
    "    window.moveTo((screen.width-900)/2, (screen.height-643)/2);" & vbCrLf & _
    "    function getElementByClass(className) {" & vbCrLf & _
    "      var all = document.all || document.getElementsByTagName('*');" & vbCrLf & _
    "      for (var i = 0; i < all.length; i++) {" & vbCrLf & _
    "        var cls = all[i].className;" & vbCrLf & _
    "        if (cls && cls.indexOf(className) !== -1) {" & vbCrLf & _
    "          return all[i];" & vbCrLf & _
    "        }" & vbCrLf & _
    "      }" & vbCrLf & _
    "      return null;" & vbCrLf & _
    "    }" & vbCrLf & _
    "    function fadeIn(el, duration) {" & vbCrLf & _
    "      var start = new Date().getTime();" & vbCrLf & _
    "      var step = function() {" & vbCrLf & _
    "        var timePassed = new Date().getTime() - start;" & vbCrLf & _
    "        var progress = timePassed / duration;" & vbCrLf & _
    "        var opacity = progress < 1 ? progress : 1;" & vbCrLf & _
    "        el.style.filter = 'alpha(opacity=' + (opacity * 100) + ')';" & vbCrLf & _
    "        el.style.opacity = opacity;" & vbCrLf & _
    "        if (progress < 1) {" & vbCrLf & _
    "          setTimeout(step, 10);" & vbCrLf & _
    "        }" & vbCrLf & _
    "      };" & vbCrLf & _
    "      step();" & vbCrLf & _
    "    }" & vbCrLf & _
    "    window.onload = function() {" & vbCrLf & _
    "      var msgElement = getElementByClass('message');" & vbCrLf & _
    "      if (msgElement) {" & vbCrLf & _
    "        fadeIn(msgElement, 2000);" & vbCrLf & _
    "      } else {" & vbCrLf & _
    "        alert('Ошибка: элемент с классом ""message"" не найден!');" & vbCrLf & _
    "      }" & vbCrLf & _
    "    };" & vbCrLf & _
    "    setTimeout(function() { window.close(); }, 3000);" & vbCrLf & _
    "  </script>" & vbCrLf & _
    "</head>" & vbCrLf & _
    "<body>" & vbCrLf & _
    "  <div class='container'>" & vbCrLf & _
    "    <div class='message'>" & msg & "</div>" & vbCrLf & _
    "  </div>" & vbCrLf & _
    "</body>" & vbCrLf & _
    "</html>"
End Function
