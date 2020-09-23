VERSION 5.00
Begin VB.Form frmLoadFileIntoArray 
   Caption         =   "GetFileIntoArray"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstResult 
      Height          =   2595
      Left            =   60
      TabIndex        =   3
      Top             =   1230
      Width           =   4785
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Excel file"
      Height          =   525
      Index           =   2
      Left            =   3330
      TabIndex        =   2
      Top             =   120
      Width           =   1545
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Fixed length"
      Height          =   525
      Index           =   1
      Left            =   1710
      TabIndex        =   1
      Top             =   120
      Width           =   1545
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Separated"
      Height          =   525
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label lblLabel 
      Caption         =   $"loadFileIntoArray.frx":0000
      Height          =   495
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   3870
      Width           =   4785
   End
   Begin VB.Label lblLabel 
      Caption         =   "Array after run:"
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   930
      Width           =   1245
   End
End
Attribute VB_Name = "frmLoadFileIntoArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTest_Click(Index As Integer)
   ' Declare variables
   Dim lngRows As Long
   Dim lngTemp As Long
   
   ' Clear listbox
   lstResult.Clear
   
   ' Check wich button is pressed
   Select Case Index
      Case 0 ' Comma separated (or any other separator)
         lngRows = LoadFileIntoArray(arrArray(), App.Path + "\atest.txt", 0, 1, 6)
      Case 1 ' Fixed length
         lngRows = LoadFileIntoArray(arrArray(), App.Path + "\btest.txt", 0, 1, 6, , , "2,2,2,2,2,2")
      Case 2 ' Excel file
         lngRows = LoadFileIntoArray(arrArray(), App.Path + "\ctest.xls", 1, 1, 6)
      ' End check wich button
      End Select
         
   ' Check if loaded
   If lngRows = 0 Then
      ' Display to user
      MsgBox "No file or records found!", vbInformation
      ' Exit
      Exit Sub
   ' End check if loaded
   End If
      
   ' Loop array
   For lngTemp = 1 To lngRows
      ' Add to listbox
      lstResult.AddItem arrArray(lngTemp)
   ' End loop array
   Next lngTemp
   
   ' Refresh listbox
   lstResult.Refresh
      
End Sub
