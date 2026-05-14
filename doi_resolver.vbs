Set args = WScript.Arguments
If args.Count = 0 Then WScript.Quit

Dim rawUrl, cleanUrl
rawUrl = args(0)

' Strip away the custom protocol trigger prefix
cleanUrl = Replace(rawUrl, "doi:", "")

' Check if Obsidian or the text already provided a full URL layout
If InStr(1, cleanUrl, "doi.org", vbTextCompare) > 0 Then
    ' If it already contains doi.org, launch it directly
    CreateObject("WScript.Shell").Run """" & cleanUrl & """"
Else
    ' If it is just the raw numeric string, attach the resolver URL path
    CreateObject("WScript.Shell").Run "https://doi.org" & cleanUrl
End If
