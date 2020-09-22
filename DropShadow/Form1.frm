VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000018&
   Caption         =   "Drop Shadow"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3315
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
   ScaleHeight     =   2550
   ScaleWidth      =   3315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Show Shadow"
      Height          =   375
      Left            =   990
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------
'                   Real Drop Shadow Form
'Shows how to add shadow to form without using any subclassing.
'All using WinAPIs. Just 8 Lines of code.
'----------------------------------------------------------------
'Written By: abhishek
'Email: abhishek007p@hotmail.com
'       binarylife9@yahoo.com
'Add me to your yahoo or msn if u like to chat about programming
'----------------------------------------------------------------

Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const CS_DROPSHADOW As Long = &H20000
Private Const GCL_STYLE     As Long = -26

Private Sub ApplyDropShadow(ByVal hWnd As Long)
    Call SetClassLong(hWnd, GCL_STYLE, GetClassLong(hWnd, GCL_STYLE) Or CS_DROPSHADOW)
End Sub

Private Sub Form_Load()
    Call ApplyDropShadow(Me.hWnd)
End Sub

'=============================================================================
'                   Love VB6? then Sign this Petition
'   A PETITION FOR THE DEVELOPMENT OF UNMANAGED VISUAL BASIC AND VBA
'
'                   http://www.classicvb.org/Petition
'Include this message in all your PSC or other code submision if possible
'pass it to others developers. i dont care about votes, instead sign petition.
'=============================================================================
'
'Why?
'Think, u build a litte app, but it needs a 24MB .NET Runtime to run.
'Think, ur code can be de-complied
'
'Think...
'
