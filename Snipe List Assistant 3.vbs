' Set up variables
Set objShell = CreateObject("WScript.Shell")
currentDate = Date()

' Display input box with today's date as default
inputDate = InputBox("Enter snipe date (mm/dd/yyyy) or click Cancel to exit:" & vbCrLf & vbCrLf & _
            "Today is: " & FormatDateTime(Date(), vbShortDate), _
            "Snipe List Calculator", _
            FormatDateTime(Date(), vbShortDate))

' Exit if user clicks Cancel
If inputDate = "" Then
    WScript.Quit
End If

' Validate date input
On Error Resume Next
currentDate = CDate(inputDate)
If Err.Number <> 0 Then
    MsgBox "Invalid date format. Please use mm/dd/yyyy.", vbExclamation, "Error"
    WScript.Quit
End If
On Error GoTo 0

' Calculate all dates
dayOfWeek = WeekdayName(Weekday(currentDate))
threeDayExactly = DateAdd("d", 11, DateAdd("m", 1, currentDate))
lastCalledBefore = DateAdd("d", -4, currentDate)
fortyFiveDayExactly = currentDate
lastLoadBefore = DateAdd("d", -105, currentDate)

' Create HTML with exact dimensions and all features
html = "<!DOCTYPE html><html><head>" & _
       "<title>Snipe List Assistant</title>" & _
       "<style>" & _
       "body { font-family: comic sans ms, comic sans; font-size: 14px; padding: 20px; margin: 0; overflow: hidden; }" & _
       "h2 { color: #333; margin-bottom: 5px; }" & _
       "h3 { color: #555; margin: 20px 0 10px 0; border-bottom: 1px solid #eee; padding-bottom: 5px; }" & _
       ".date-container { margin-bottom: 15px; }" & _
       ".date-label { font-weight: bold; margin-right: 10px; min-width: 150px; display: inline-block; }" & _
       ".date-value { color: #0066CC; cursor: pointer; text-decoration: underline; padding: 2px 5px; }" & _
       ".date-value:hover { background: #E6F2FF; }" & _
       "#status { color: green; margin-top: 15px; font-style: italic; }" & _
       ".separator { border-top: 2px solid #ddd; margin: 20px 0; }" & _
       "</style>" & _
       "<script>" & _
       "function copyToClipboard(text) {" & _
       "  var textarea = document.createElement('textarea');" & _
       "  textarea.value = text;" & _
       "  document.body.appendChild(textarea);" & _
       "  textarea.select();" & _
       "  document.execCommand('copy');" & _
       "  document.body.removeChild(textarea);" & _
       "  document.getElementById('status').innerHTML = 'Copied: ' + text;" & _
       "  setTimeout(function(){ document.getElementById('status').innerHTML = ''; }, 2000);" & _
       "}" & _
       "window.resizeTo(396,488);" & _
       "</script>" & _
       "</head><body>" & _
       "<h2>Snipe List Dates</h2>" & _
       "<p><b>Snipe date:</b> " & dayOfWeek & " " & FormatDateTime(currentDate, vbShortDate) & "</p>" & _
       "<h3>45 Day List</h3>" & _
       "<div class='date-container'>" & _
       "<div><span class='date-label'>Fall Off Date ""Is Exactly"":</span> <span class='date-value' onclick=""copyToClipboard('" & Replace(FormatDateTime(fortyFiveDayExactly, vbShortDate), "'", "\'") & "')"">" & FormatDateTime(fortyFiveDayExactly, vbShortDate) & "</span></div>" & _
       "<div><span class='date-label'>Last Load Date ""Is Before"":</span> <span class='date-value' onclick=""copyToClipboard('" & Replace(FormatDateTime(lastLoadBefore, vbShortDate), "'", "\'") & "')"">" & FormatDateTime(lastLoadBefore, vbShortDate) & "</span></div>" & _
       "</div>" & _
       "<div class='separator'></div>" & _
       "<h3>3 Day List</h3>" & _
       "<div class='date-container'>" & _
       "<div><span class='date-label'>3 Day Fall Off ""Is Exactly"":</span> <span class='date-value' onclick=""copyToClipboard('" & Replace(FormatDateTime(threeDayExactly, vbShortDate), "'", "\'") & "')"">" & FormatDateTime(threeDayExactly, vbShortDate) & "</span></div>" & _
       "<div><span class='date-label'>Last Called ""Is Before"":</span> <span class='date-value' onclick=""copyToClipboard('" & Replace(FormatDateTime(lastCalledBefore, vbShortDate), "'", "\'") & "')"">" & FormatDateTime(lastCalledBefore, vbShortDate) & "</span></div>" & _
       "</div>" & _
       "<div id='status'></div>" & _
       "<p><i>Click any date to copy it to clipboard</i></p>" & _
       "</body></html>"

' Create temporary HTML file
Set fso = CreateObject("Scripting.FileSystemObject")
tempFile = fso.GetSpecialFolder(2) & "\ProspectTaggingResults.html"
Set file = fso.CreateTextFile(tempFile, True)
file.Write html
file.Close

' Open in default browser with exact dimensions
objShell.Run "mshta " & Chr(34) & tempFile & Chr(34)

' Also copy all dates to clipboard in plain text with new format
plainText = "Snipe List Dates" & vbCrLf & _
            "Using date: " & dayOfWeek & " " & FormatDateTime(currentDate, vbShortDate) & vbCrLf & vbCrLf & _
            "=== 45 DAY LIST ===" & vbCrLf & _
            "Fall Off Date ""Is Exactly"": " & FormatDateTime(fortyFiveDayExactly, vbShortDate) & vbCrLf & _
            "Last Load Date ""Is Before"": " & FormatDateTime(lastLoadBefore, vbShortDate) & vbCrLf & vbCrLf & _
            "=== 3 DAY LIST ===" & vbCrLf & _
            "3 Day Fall Off ""Is Exactly"": " & FormatDateTime(threeDayExactly, vbShortDate) & vbCrLf & _
            "Last Called ""Is Before"": " & FormatDateTime(lastCalledBefore, vbShortDate)

objShell.Run "cmd /c echo " & Chr(34) & plainText & Chr(34) & " | clip", 0, True

' Clean up after 10 seconds
WScript.Sleep 10000
On Error Resume Next
fso.DeleteFile tempFile
On Error GoTo 0