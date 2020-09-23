VERSION 5.00
Begin VB.Form frmTestFriend 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select  some items...."
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test Component Friend Method"
      Height          =   420
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   2790
   End
End
Attribute VB_Name = "frmTestFriend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//---Take a look at the object browser and you will notice
'//---that the DefineSelection Friend is not visible

Private WithEvents oTestClass As FriendExample.cPublicClass
Attribute oTestClass.VB_VarHelpID = -1

Private Sub cmdTest_Click()
    'call the ui class to get a selection
    oTestClass.GetPropertyFromUI
End Sub

Private Sub Form_Load()
    'create our new object
    Set oTestClass = New FriendExample.cPublicClass
End Sub

Private Sub oTestClass_SelectionChanged(newValue As String)
    MsgBox oTestClass.CurrentSelection
End Sub
