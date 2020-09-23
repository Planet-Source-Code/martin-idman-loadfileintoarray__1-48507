Attribute VB_Name = "modLoadFileIntoArray"
' Force variable declaration
Option Explicit

' Declare variables
Public arrArray() As String

' +------------------------------------------------------------------------------------------------------+
' | Function    : LoadFileIntoArray                                                                      |
' | Returns     : Number of rows added to passed array                                                   |
' | Author      : Martin Idman                                                                           |
' | Description : Loads a separated-, fixed lenght- or Excel-file into an array                          |
' | Parameters  : arrLoadFileIntoArray                 - Array to store loaded rows                      |
' |               strLoadFileIntoArrayFile             - Filename including path to load                 |
' |               intLoadFileIntoArrayType             - Type of file to load, 0 = text/fixed, 1 = Excel |
' |               lngLoadFileIntoArrayStart            - Which row to start from                         |
' |               lngLoadFileIntoArrayColumnsCount     - Number of columns in file                       |
' |               strLoadFileIntoArrayInSeparator      - Separator if separated file                     |
' |               strLoadFileIntoArrayOutSeparator     - Separator between columns in output array       |
' |               strLoadFileIntoArrayFixedDescription - Lengths for each column in fixed length file    |
' | Other       : Needs Microsoft Excel 9.0 Object Library to be added to Project References             |
' +------------------------------------------------------------------------------------------------------+
Public Function LoadFileIntoArray(ByRef arrLoadFileIntoArray() As String, _
                                 ByVal strLoadFileIntoArrayFile As String, _
                                 ByVal intLoadFileIntoArrayType As Integer, _
                                 Optional ByVal lngLoadFileIntoArrayStart As Long = 1, _
                                 Optional ByVal lngLoadFileIntoArrayColumnsCount As Long, _
                                 Optional ByVal strLoadFileIntoArrayInSeparator As String = ",", _
                                 Optional ByVal strLoadFileIntoArrayOutSeparator As String = "||", _
                                 Optional ByVal strLoadFileIntoArrayFixedDescription As String) As Long
                                 
   ' Declare variables
   Dim lngLoadFileIntoArrayRow As Long            ' Temporary row counter
   Dim lngLoadFileIntoArrayColumn As Long         ' Temporary column counter
   Dim bolLoadFileIntoArrayRowEmpty As Boolean    ' Temporary check for empty row
   Dim intLoadFileIntoArrayFileNumber As Integer  ' Temporary file number opened
   Dim strLoadFileIntoArrayIn As String           ' Temporary in string
   Dim strLoadFileIntoArrayOut As String          ' Temporary out string
   Dim lngLoadFileIntoArrayLength As Long         ' Temporary column length
   Dim lngLoadFileIntoArrayStartPosition As Long  ' Temporary column start position
        
   ' Check if file exists
   If Dir(strLoadFileIntoArrayFile) = "" Then
      ' Return 0 to calling sub
      LoadFileIntoArray = 0
      ' Exit the function
      Exit Function
   ' End check if file exists
   End If

' Check wich filetype
   Select Case intLoadFileIntoArrayType
      Case 0 ' Text file
         ' Set first available free filenumber
         intLoadFileIntoArrayFileNumber = FreeFile
         ' Open file
         Open strLoadFileIntoArrayFile For Input As #intLoadFileIntoArrayFileNumber
         
         ' Loop to start position
         For lngLoadFileIntoArrayRow = 2 To lngLoadFileIntoArrayStart
            ' Get line
            Line Input #1, strLoadFileIntoArrayIn
         ' End loop to start position
         Next lngLoadFileIntoArrayRow
         ' Reset row number
         lngLoadFileIntoArrayRow = 1
            
         ' Loop all rows until end of file
         While Not EOF(intLoadFileIntoArrayFileNumber)
            ' Reset row variable
            strLoadFileIntoArrayOut = ""
            ' Get line
            Line Input #intLoadFileIntoArrayFileNumber, strLoadFileIntoArrayIn
            ' Reset start position
            lngLoadFileIntoArrayStartPosition = 1
            ' Loop all columns
            For lngLoadFileIntoArrayColumn = 1 To lngLoadFileIntoArrayColumnsCount
               ' Check if separated file
               If strLoadFileIntoArrayFixedDescription = "" Then ' Separated
                  ' Set value to row variable
                  strLoadFileIntoArrayOut = strLoadFileIntoArrayOut + strLoadFileIntoArrayOutSeparator + Split(strLoadFileIntoArrayIn, strLoadFileIntoArrayInSeparator)(lngLoadFileIntoArrayColumn - 1)
               Else ' Fixed file length
                  ' Set current column length
                  lngLoadFileIntoArrayLength = Split(strLoadFileIntoArrayFixedDescription, ",")(lngLoadFileIntoArrayColumn - 1)
                  ' Set value to row variable
                  strLoadFileIntoArrayOut = strLoadFileIntoArrayOut + strLoadFileIntoArrayOutSeparator + Mid(strLoadFileIntoArrayIn, lngLoadFileIntoArrayStartPosition, lngLoadFileIntoArrayLength)
                  ' Add to start position counter
                  lngLoadFileIntoArrayStartPosition = lngLoadFileIntoArrayStartPosition + lngLoadFileIntoArrayLength
               ' End check if separated file
               End If
            ' End loop all columns
            Next lngLoadFileIntoArrayColumn
                  
            ' Redimension variable
            ReDim Preserve arrLoadFileIntoArray(lngLoadFileIntoArrayRow)
            ' Add to array
            arrLoadFileIntoArray(lngLoadFileIntoArrayRow) = strLoadFileIntoArrayOut
            ' Add to row counter
            lngLoadFileIntoArrayRow = lngLoadFileIntoArrayRow + 1
         ' End loop all rows
         Wend
         ' Close file
         Close #intLoadFileIntoArrayFileNumber
         
      Case 1 ' Excel file
         ' Declare Excel application object variable
         Dim objLoadFileIntoArrayExcel As Excel.Application
         ' Set Excel application object variable
         Set objLoadFileIntoArrayExcel = Excel.Application
         ' Open workbook
         objLoadFileIntoArrayExcel.Workbooks.Open strLoadFileIntoArrayFile
         ' Choose first sheet in workbook
         objLoadFileIntoArrayExcel.Worksheets(1).Activate
         ' Set start value
         lngLoadFileIntoArrayRow = lngLoadFileIntoArrayStart
         ' Loop all rows in sheet
         Do
            ' Reset emty check variable
            bolLoadFileIntoArrayRowEmpty = True
            ' Reset row variable
            strLoadFileIntoArrayOut = ""
            ' Loop columns in row
            For lngLoadFileIntoArrayColumn = 1 To lngLoadFileIntoArrayColumnsCount
            ' Set value to row variable
               strLoadFileIntoArrayOut = strLoadFileIntoArrayOut + strLoadFileIntoArrayOutSeparator + Trim(objLoadFileIntoArrayExcel.Cells(lngLoadFileIntoArrayRow, lngLoadFileIntoArrayColumn).Value)
            ' End loop columns in row
            Next lngLoadFileIntoArrayColumn
            ' Check if empty
            If lngLoadFileIntoArrayColumnsCount * Len(strLoadFileIntoArrayOutSeparator) = Len(strLoadFileIntoArrayOut) Then
               ' Set boolean to empty
               bolLoadFileIntoArrayRowEmpty = True
            Else
               ' Redimension variable
               ReDim Preserve arrLoadFileIntoArray(lngLoadFileIntoArrayRow - lngLoadFileIntoArrayStart + 1)
               ' Add to array
               arrLoadFileIntoArray(lngLoadFileIntoArrayRow - lngLoadFileIntoArrayStart + 1) = strLoadFileIntoArrayOut
               ' Set boolean to not empty
               bolLoadFileIntoArrayRowEmpty = False
            ' End check if empty
            End If
            ' Add to row counter
            lngLoadFileIntoArrayRow = lngLoadFileIntoArrayRow + 1
         ' End loop all rows
         Loop Until lngLoadFileIntoArrayRow > 65535 Or bolLoadFileIntoArrayRowEmpty = True
         ' Close workbook
         ActiveWorkbook.Close
         ' Quit Excel object
         objLoadFileIntoArrayExcel.Quit
         ' Kill Excel object
         Set objLoadFileIntoArrayExcel = Nothing
   End Select
   ' Return value to calling sub
   LoadFileIntoArray = UBound(arrLoadFileIntoArray)
End Function


