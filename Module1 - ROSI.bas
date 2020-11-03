Attribute VB_Name = "Module1"
Sub ROSI_5AD()

Dim browser As Object
Dim Workbook_Name As String
Dim Worksheet_Name As String
Dim studentid As Range
Dim gradsessioncd As String
Dim postcd As String
Dim ROSI_message2 As String
Dim POST_cd As String
Dim person_id As String


Workbook_Name = ActiveWorkbook.Name
Worksheet_Name = ActiveSheet.Name

Set studentid = Application.InputBox("Click the first student number, and drag down to the last student number to highlight all the student numbers. Do not include the header.", _
Title:="Select Range of Student Numbers", _
Type:=8)

Set browser = CreateObject("internetexplorer.application")

'production environment
'browser.Navigate "https://javnat-admin-qa.easi.utoronto.ca/ROSI/rosi"

'live environment
browser.Navigate "https://admin.rosi.utoronto.ca/ROSI/rosi"
browser.Visible = True

On Error Resume Next
browser.document.all("JavNat_PF12").Click

While browser.Busy
DoEvents
Wend

browser.document.all("field2").Value = "5 A D"
browser.document.all("JavNat_ENTR").Click

While browser.Busy
DoEvents
Wend

For Each Cell In studentid

Cell.Activate

gradsessioncd = "20209"
POST_cd = ActiveCell.Offset(0, 4).Value
person_id = ActiveCell.Value

If POST_cd <> "" Then

browser.document.all("#ACTION").Value = "A"
browser.document.all("SGA006.SESSION_CD").Value = gradsessioncd
browser.document.all("SGA006.PERSON_ID").Value = person_id
browser.document.all("SGA006.POST_CD").Value = POST_cd
browser.document.all("JavNat_ENTR").Click

While browser.Busy
DoEvents
Wend

Application.Wait (Now + TimeValue("0:00:01"))
ROSI_message2 = browser.document.all("JavNat_Message").innerhtml
Cells(ActiveCell.Row, 23).Value = ROSI_message2
End If

If POST_cd = "" Then
ROSI_message2 = "No POST_CD entered"
Cells(ActiveCell.Row, 23).Value = ROSI_message2
End If

If Cells(ActiveCell.Row, 23).Value = " Degree Request already exists" Then
Cells(ActiveCell.Row, 1).Font.ColorIndex = 46
Else: Cells(ActiveCell.Row, 1).Font.ColorIndex = 10
End If

If Cells(ActiveCell.Row, 23).Value = "6195 - Access to this POSt code is denied" Or Cells(ActiveCell.Row, 23).Value = "6556 - Student is not presently active in a POSt" Or Cells(ActiveCell.Row, 23).Value = "No POST_CD entered" Or Cells(ActiveCell.Row, 23).Value = "6205 - POSt is not being offered in this session" Then
Cells(ActiveCell.Row, 1).Font.ColorIndex = 3
End If



Next

MsgBox ("Graduation dates have been added to ROSI! Review the students with student numbers in blue/red/orange.")


End Sub

