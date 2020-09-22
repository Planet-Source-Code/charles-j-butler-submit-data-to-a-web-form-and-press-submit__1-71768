VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Submit Form Example"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Data"
      Height          =   435
      Left            =   2280
      TabIndex        =   2
      Top             =   3840
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Data"
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   1035
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3675
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   3615
      ExtentX         =   6376
      ExtentY         =   6482
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################
'# Submit Form Example by Charles J Butler cbutler@defonic.com
'################################################

Option Explicit
Private Sub Form_Load()
    
    '  Locate your form, we use a test form but this is intended for you web form.

    WebBrowser1.Navigate App.Path & "/form.htm"
    
End Sub

Private Sub Command1_Click()

    WebBrowser1.Document.Forms(0).elements(0).Value = "Funny video of dog" ' The first form data entry
    WebBrowser1.Document.Forms(0).elements(1).Value = "http://test.com/test.avi" ' The second form data entry
    WebBrowser1.Document.Forms(0).elements(2).Value = "1" 'This will select the index of the combobox (drop down box)
    WebBrowser1.Document.Forms("data").All("Submit").Click 'This will click the submit/post button on your form, you must specify the actual form name

End Sub





