'!!!!! NOTE: The code below comes from QOLAddings.xlam. I have included it here so it can be publically viewed but to get the addin working just download the .xlam file instead. This code is located 
'!!!!! Workbook module of QOLAddins.xlam.

Private Sub Workbook_Open()

Call Ribbon_Creation


End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
Call Ribbon_Deletion

End Sub
Private Sub Ribbon_Creation()

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////// _
'_____________________________ Creating Custom Ribbon Objects attached to Macros _________________________ _
'//////////////////// This sub modifies the standard ribbon file for office to incorporate new buttons which activate macros _
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Modified from script on here: https://stackoverflow.com/questions/37974447/get-name-or-id-of-a-ribbon-controls
'NOTES: Orginially, this script had been written to try and create a custom ribbon file in 2 different places before modifying the excel default _
        ribbon file. However, I cannot get excel to load the custom ribbon files, so I will just have to settle with modifying the default. However, _
        I have left the old code in case it becomes a problem down the road
Dim hFile As Long
Dim path As String, fileName As String, ribbonXML As String, user As String
If Not Left(Application.OperatingSystem, 1) Like "W" Then 'if its windows or not. changes the slashes
    Var_Slash = "/"
Else
    Var_Slash = "\"
End If

hFile = FreeFile
user = Environ("Username")
path = "C:" & Var_Slash & "Users" & Var_Slash & user & Var_Slash & "AppData" & Var_Slash & "Local" & Var_Slash & "Microsoft" & Var_Slash & "Office" & Var_Slash
fileName = "Excel.officeUI"

ribbonXML = "<mso:customUI      xmlns:mso='http:" & Var_Slash & Var_Slash & "schemas.microsoft.com" & Var_Slash & "office" & Var_Slash & "2009" & Var_Slash & "07" & Var_Slash & "customui'>" & vbNewLine
ribbonXML = ribbonXML + "  <mso:ribbon>" & vbNewLine
ribbonXML = ribbonXML + "    <mso:qat/>" & vbNewLine
ribbonXML = ribbonXML + "    <mso:tabs>" & vbNewLine
ribbonXML = ribbonXML + "      <mso:tab id='reportTab' label='QOL Features' insertBeforeQ='mso:TabFormat'>" & vbNewLine
ribbonXML = ribbonXML + "        <mso:group id='reportGroup' label='File Manipulation' autoScale='true'>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='SaveBackup' label='Save Backup' " & vbNewLine
ribbonXML = ribbonXML + "imageMso='FileSave'      onAction='BackingUpFiles_Addin.SaveWorkbookBackUp'/>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='CloseFiles' label='Close Other Excel Files' " & vbNewLine
ribbonXML = ribbonXML + "imageMso='FileClose'      onAction='BackingUpFiles_Addin.Save_And_Close_All_Workbooks_Except_This_One'/>" & vbNewLine
ribbonXML = ribbonXML + "        </mso:group>" & vbNewLine
ribbonXML = ribbonXML + "        <mso:group id='EmailGroup' label='Write To Email (Old Outlook App Only)' autoScale='true'>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='AutoEmailSelect' label='Current Selection' " & vbNewLine
ribbonXML = ribbonXML + "imageMso='CreateEmail'      onAction='BackingUpFiles_Addin.AutoEmail_Copy_CurrentSelection_ToEmail'/>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='AutoEmailPrint' label='Print Area' " & vbNewLine
ribbonXML = ribbonXML + "imageMso='GroupPrintPreviewPrint'      onAction='BackingUpFiles_Addin.AutoEmail_Copy_PrintArea_ToEmail'/>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='AutoEmailPrintPDF' label='Print Area as PDF' " & vbNewLine
ribbonXML = ribbonXML + "imageMso='PublishToPdfOrEdoc'      onAction='BackingUpFiles_Addin.AutoEmail_Attach_PDF_Of_PrintArea_ToEmail'/>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='AutoEmailSelectXLSX' label='Cur. Sel. as .xlsx' " & vbNewLine
ribbonXML = ribbonXML + "imageMso='FileSaveAsExcelXlsx'      onAction='BackingUpFiles_Addin.AutoEmail_Attach_Xlsx_Of_Selection_ToEmail'/>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='AutoEmailPrintXlsx' label='Print Area as Xlsx' " & vbNewLine
ribbonXML = ribbonXML + "imageMso='FileSaveAsExcelXlsx'      onAction='BackingUpFiles_Addin.AutoEmail_Attach_Xlsx_Of_PrintArea_ToEmail'/>" & vbNewLine
ribbonXML = ribbonXML + "        </mso:group>" & vbNewLine
ribbonXML = ribbonXML + "      </mso:tab>" & vbNewLine
ribbonXML = ribbonXML + "    </mso:tabs>" & vbNewLine
ribbonXML = ribbonXML + "  </mso:ribbon>" & vbNewLine
ribbonXML = ribbonXML + "</mso:customUI>"

ribbonXML = Replace(ribbonXML, """", "")

Open path & fileName For Output Access Write As hFile
Print #hFile, ribbonXML
Close hFile
End Sub
Private Sub Ribbon_Deletion()
'This sub works in conjunction with Ribbon_Open. The idea is that when you open this addin, it will stay in the _
ribbon until you close the last workbook, at which point it will trigger this script to run, removing the addin _
from the ribbon the next time you open excel
Dim hFile As Long
Dim path As String, fileName As String, ribbonXML As String, user As String
If Not Left(Application.OperatingSystem, 1) Like "W" Then 'if its windows or not. changes the slashes
    Var_Slash = "/"
Else
    Var_Slash = "\"
End If
hFile = FreeFile
user = Environ("Username")
path = "C:" & Var_Slash & "Users" & Var_Slash & user & Var_Slash & "AppData" & Var_Slash & "Local" & Var_Slash & "Microsoft" & Var_Slash & "Office" & Var_Slash
fileName = "Excel.officeUI"

ribbonXML = "<mso:customUI           xmlns:mso=""http:" & Var_Slash & Var_Slash & "schemas.microsoft.com" & Var_Slash & "office" & Var_Slash & "2009" & Var_Slash & "07" & Var_Slash & "customui"">" & _
"<mso:ribbon></mso:ribbon></mso:customUI>"

Open path & fileName For Output Access Write As hFile
Print #hFile, ribbonXML
Close hFile
End Sub



