VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Show all chars in a font by Johannes B 2001"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3960
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Columns         =   10
      Height          =   2985
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Selected char"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Font:"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub showf()
'Show current font
On Error GoTo kallena
Dim JB As Integer

List1.Clear
JB = 0
Do
JB = JB + 1
List1.AddItem Chr(JB)
Loop
Exit Sub
kallena:
Exit Sub

End Sub

Private Sub Combo1_Click()
List1.FontName = Combo1.Text
Call showf
List1.Height = "2985"
End Sub


Private Sub Form_Load()

Dim Counter1 As Long
    For Counter1 = 0 To Screen.FontCount - 1
        Combo1.AddItem Screen.Fonts(Counter1)
Next Counter1

Combo1.Text = "Arial"

End Sub

Private Sub List1_Click()
Text1.FontName = List1.FontName
Text1.Text = List1.Text
End Sub


