# VbaWebScrapping
Sub LEI_validate()
Dim objIEBrowser
Dim i As Integer
i = 2
Do Until IsEmpty(Cells(i, 1))
Set objIEBrowser = CreateObject("InternetExplorer.Application")
objIEBrowser.Visible = False

objIEBrowser.Navigate2 "openleis.com/legal_entities/" & Range("B" & i)
'Application.Wait Now + TimeValue("00:00:04")
Do While objIEBrowser.Busy Or objIEBrowser.ReadyState <> 4
Loop
Dim sDD As String
sDD = Trim(objIEBrowser.Document.getElementsByTagName("h1")(0).innerText)
Range("C" & i).Value = sDD
objIEBrowser.Quit
i = i + 1
Loop
ErrMsg:
MsgBox "Done!"
End Sub

