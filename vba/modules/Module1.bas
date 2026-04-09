Attribute VB_Name = "Module1"
Option Explicit

'===================================================================================================
' Procedure: Import_Text_File
' Purpose:   Iterates through a selection of external files, importing the contiguous data range
'            from the first sheet of each into new worksheets within the primary workbook.
'===================================================================================================
Public Sub Import_Text_File()

    ' Disable application alerts and calculations to optimize performance and prevent flicker
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

    ' Declare object and variable types for workbook handling and iteration
    Dim TextFile As Workbook
    Dim OpenFiles() As Variant
    Dim i As Integer
    
    ' Call helper function to launch the File Open dialog and capture file paths
    OpenFiles = Get_Files()
    
    ' Iterate through the array of selected file paths
    ' Note: CountA is used to determine the bound of the variant array returned by GetOpenFilename
    For i = 1 To Application.CountA(OpenFiles)
    
        ' Instantiate the external workbook object by opening the file path
        Set TextFile = Workbooks.Open(OpenFiles(i))
        
        ' Identify the contiguous data range starting at A1 and copy to clipboard
        TextFile.Sheets(1).Range("A1").CurrentRegion.Copy
        
        ' Shift focus to the host workbook (index 1) to perform the import
        Workbooks(1).Activate
        
        ' Initialize a new worksheet to house the imported data
        Workbooks(1).Worksheets.Add
        
        ' Transfer clipboard data to the newly created ActiveSheet
        ActiveSheet.Paste
        
        ' Rename the worksheet to match the source filename for traceability
        ActiveSheet.Name = TextFile.Name
        
        ' Clear the clipboard to release memory resources
        Application.CutCopyMode = False
        
        ' Terminate the external workbook instance without saving changes
        TextFile.Close
    Next i
    
    ' Restore application state to default environment settings
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
End Sub

'===================================================================================================
' Function:  Get_Files
' Purpose:   Invokes the standard Windows File Explorer dialog to allow the user to select one
'            or more files. Returns an array of strings (file paths) or False if cancelled.
'===================================================================================================
Public Function Get_Files() As Variant

    ' Execute the GetOpenFilename method with MultiSelect enabled to return an array
    Get_Files = Application.GetOpenFilename(Title:="Select File(s) to Import", MultiSelect:=True)

End Function

