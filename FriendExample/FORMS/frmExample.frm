VERSION 5.00
Begin VB.Form frmExample 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selection Dialog"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3705
      TabIndex        =   2
      Top             =   2865
      Width           =   1185
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2490
      TabIndex        =   1
      Top             =   2865
      Width           =   1185
   End
   Begin VB.ListBox lstTest 
      Appearance      =   0  'Flat
      Height          =   2505
      Left            =   45
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   30
      Width           =   4815
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oCaller As cPublicClass 'private to hold the class that instantiated this form

Public Sub Initialize(ByRef mCallingClass As cPublicClass)
    Set oCaller = mCallingClass
    
    'in a more robust example, you could reset the listbox
    'with the current selection and add an incoming parameter
    'to parse, etc..........
End Sub

Private Sub cmdButton_Click(Index As Integer)
    If Index = 0 Then
        Dim tmpString As String
        Dim i As Integer
        
        ''just a string builder for the selection
        For i = 0 To lstTest.ListCount - 1
            If lstTest.Selected(i) Then
                    tmpString = tmpString & lstTest.List(i) & "|" 'just a cheap string builder for example sake
            End If
        Next
        
        
        'in the scope of this project
        'we can see the friend method
        'if the selection changed in
        'the friend, our event will fire
        
        oCaller.DefineSelection tmpString
    End If
    
    Set oCaller = Nothing
    Unload Me 'unload the form, we passed back the info if index was 0 (ok)
    
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    'just add some items
    lstTest.Clear
    For i = 0 To 30
        lstTest.AddItem "Test Item " & i
    Next
    
End Sub
