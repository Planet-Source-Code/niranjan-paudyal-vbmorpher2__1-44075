VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Morpher2"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmHelp.frx":0000
   ScaleHeight     =   176
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   248
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Click here to vote now!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ok"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Contact me before using any parts of this program."
      ForeColor       =   &H00FF0000&
      Height          =   450
      Left            =   0
      TabIndex        =   5
      Top             =   1680
      Width           =   2340
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "nirpaudyal@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   600
      MouseIcon       =   "FrmHelp.frx":1A5AA
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1440
      Width           =   1725
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "March - May 2003"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Niranjan Paudyal"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   0
      Picture         =   "FrmHelp.frx":1A8B4
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   3735
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Part of VBMorpher2 written by Niranjan paudyal
'Last updated 16May 2003
'No parts of this program may be copied or used
'without contacting me first at
'nirpaudyal@hotmail.com (see the about box)
'If you like to use the avi file making module, or
'would like any help on that module, plese contact
'the author (contact address on mAVIDecs module)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
    ShellExecute Me.HWND, "open", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=44075&lngWId=1", vbNull, vbNull, 1
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Label4_Click()
    ShellExecute Me.HWND, "open", "mailto:nirpaudyal@hotmail.com", vbNull, vbNull, 1
End Sub
