VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "AutoComplete Combo Demo"
   ClientHeight    =   1650
   ClientLeft      =   3840
   ClientTop       =   3615
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleWidth      =   3735
   Begin Project1.AutoCompleteCombo AutoCompleteCombo1 
      Height          =   315
      Left            =   525
      TabIndex        =   0
      Top             =   540
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   556
      BackColor       =   -2147483643
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      ForeColor       =   -2147483640
      Text            =   "AutoCompleteCombo1"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim iItem As Integer
Dim iCtr As Integer

With AutoCompleteCombo1
    .Clear
    

    For iItem = 65 To 90
       .AddItem Chr(iItem) & "TEST"
       .ItemData(iCtr) = iItem
       iCtr = iCtr + 1
    Next
        .ListIndex = 0
        
End With
End Sub
