
Sub vba_extract_sheet_for_IMPORT_FILE()
    
    '### Turn off various settings while the program is running to reduce flickering and otherwise increase performance. ###
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' ### Declare string variables to be used.  ###
    Dim projectTitle, startDate, endDate, quoteNumber, projectCells, phaseOneCells, phaseTwoCells, cells, fullFileName, fileSuffix, fullFilePath, recipientAddress, phaseCount, recipientName, worksheetName As String
   
    ' ### Assign variables values from the designated sheet. ###
    projectTitle = Range("E7").Value
    startDate = Range("X60").Value
    endDate = Range("X61").Value
    quoteNumber = Range("E6").Value
    phaseCount = Range("X62").Value
    
    ' ### Defining the cells referred to in different ranges pertaining to a phased or unphased project.  ###
    ' ### It just tells the program which cells we need to check, and the program chooses the correct range ###
    ' ### at runtime by concatenating (connecting) the string-type ranges into one range at runtime. ###
    phaseOneCells = "G68, G69, G70, G71, R68"
    phaseTwoCells = "G76, G77, G78, G79, R76"
    projectCells = "E6, E7, E8, F13, V6, V7, V8, V9, AB7, R8, R9, X60, X61, X62, X63"
    
    ' ### Defining the various strings used to make the filepath names and other important variables appear correctly. ###
    
    fileSuffix = "_IMPORT_FILE.xlsx"
    cleanedProjectTitle = Replace(projectTitle, "/", "")
    fullFileName = quoteNumber & fileSuffix
    worksheetName = "IMPORTS"
    fullFilePath = Application.ThisWorkbook.Path & "/" & fullFileName
    recipientAddress = "emailName@your_domain.com"   ' <----------<----------<----| EMAIL RECIPIENT ADDRESS
    recipientName = "Sam"   ' <----------<--------<----------<----------<-------| EMAIL RECIPIENT NAME
    
     ' ### Declare Outlook email objects. ###
    Dim EmailApp As Outlook.Application
    Dim NewEmailItem As Outlook.MailItem
    
    Set EmailApp = New Outlook.Application
    Set NewEmailItem = EmailApp.CreateItem(olMailItem)
    
'====================================================================================================================
' ### If the phase name cell for phase 2 is empty, we can assume it is a single phase project, so set the boolean ###
' ### to true and tell the program to only check for emptiness of the cells pertaining to the main project.       ###
' ### Otherwise it is set to multiphase and we check for the other cells too.                                     ###
'====================================================================================================================
    
    If IsEmpty(Range("H35").Value) Then
        Dim hasOnePhase As Boolean
        hasOnePhase = True
        cells = projectCells
        
    Else
        Dim hasTwoPhases As Boolean
        hasTwoPhases = True
        cells = projectCells & ", " & phaseOneCells & ", " & phaseTwoCells

    End If
    
'====================================================================================================================================================================
'### Check for empty fields, color them red, display an appropriate prompt and exit the script if there are empty required (*) fields when the button is clicked. ###
'====================================================================================================================================================================
    
    Set requiredCells = Range(cells)
    Dim hasUnfilledCells As Boolean
    
    For Each Item In requiredCells
        If IsEmpty(Item) Then
            ' ### This RGB() is the color the UN-filled cells turn to. ###
            Item.Interior.Color = RGB(255, 80, 80)
            hasUnfilledCells = True
            
        Else
            ' ### This RGB() is the color the subsequently-filled cells return to after being corrected when turned red. ###
            Item.Interior.Color = RGB(174, 170, 170)
            
        End If
    Next Item
    
    If hasUnfilledCells Then
    
        Application.ScreenUpdating = True
        
        MsgBox "The sheet was not extracted because some mandatory fields are not filled out.  See them in red. ", Title:="Unfilled Required Fields"
        
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        
        Exit Sub
    End If
    
'==========================================================================

    ' ### Declare "Workbook" object called "wb", Add the Workbook object stored in the wb variable to the "Workbooks" collection. ###
    
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    ' ### Save the copied sheet as "projectTitle_IMPORT_FILE.xls". Use the generic path to grab the current sheet's location dynamically so this script is portable. ###
    wb.SaveAs fullFilePath
    
    ' ### Copy the specified sheet's (IMPORTS) rows from this macro's workbook based on the number of phases, and place it in the new workbook stored in the "wb" variable.  Currently only handles 2 phases.  ###
    Workbooks.Open fullFilePath
    If hasOnePhase Then
        ThisWorkbook.Sheets(worksheetName).Range("A1:T2").Copy Workbooks(fullFileName).Worksheets("Sheet1").Range("A1:T2")
        
    ElseIf hasTwoPhases Then
        ThisWorkbook.Sheets(worksheetName).Range("A1:T4").Copy Workbooks(fullFileName).Worksheets("Sheet1").Range("A1:T4")
        
    Else
        MsgBox "There was an Error copying the sheet.  The operation was not completed.", Title:="Error Copying Sheet!"
        Exit Sub
        
    End If
    
    wb.SaveAs fullFilePath
    
    ' ### Close the new workbook, since it will otherwise automatically open and disrupt the file's subsequent import to the external application.
    
    Workbooks(fullFileName).Close
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    '________________________________________
    '======= SEND EMAIL NOTIFICATION ========
    '----------------------------------------
   
    NewEmailItem.To = recipientAddress
    NewEmailItem.Subject = fullFileName & " is ready For import."
    
    ' ### Body of the Email: underscores indicate continuation on the next line. Ampersands concatenate strings. ###
    
    NewEmailItem.HTMLBody = "Hi " & recipientName & "," & vbNewLine & vbNewLine & _
                            "this is an email notification from Excel to inform you that " & fullFileName & " is ready to be imported.  You can find it in the designated quote folder For quote #" & quoteNumber & "." & " " & " " & _
                            "The project is scheduled to start on " & startDate & " and scheduled to end on " & endDate & ", " & "and the number of phases is " & phaseCount & "."
    
    ' ### Attach the file via the fullFilePath ###
    'Src = fullFilePath
    'NewEmailItem.Attachments.Add Src
    
    NewEmailItem.Display False
    NewEmailItem.Send
    
    '=========================================================
    
    ' ### Display end of process message. ###
    MsgBox "Extraction successful!  Email sent to " & recipientAddress & "."
    
    ' ### Turn screen updating back on and declare the end of the procedure. ###
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub
