Set args = WScript.Arguments
If args.Count = 0 Then WScript.Quit

Dim rawUrl, cleanUrl
rawUrl = args(0)

' 1. Standardise incoming data by stripping the protocol trigger prefix
cleanUrl = Replace(rawUrl, "doi:", "")

' 2. Handle redundant URL fragments left behind by custom app formats
' (Transforms anomalies like "https//doi.org" or double prefixes into a clean string)
cleanUrl = Replace(cleanUrl, "https://doi.org/", "")
cleanUrl = Replace(cleanUrl, "http://doi.org", "")
cleanUrl = Replace(cleanUrl, "doi.org/", "")

' 3. Safely launch the final formatted target URL in your system default browser
Dim finalTarget
finalTarget = "https://doi.org/" & cleanUrl

CreateObject("WScript.Shell").Run """" & finalTarget & """"
