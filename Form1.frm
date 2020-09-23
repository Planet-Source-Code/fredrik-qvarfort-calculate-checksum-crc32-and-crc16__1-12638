VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CRC-16 and CRC32 Checksum Example"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Benchmark"
      Height          =   375
      Left            =   2220
      TabIndex        =   11
      Top             =   4755
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process"
      Height          =   375
      Left            =   900
      TabIndex        =   12
      Top             =   4755
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Information"
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   3315
      Width           =   4215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   285
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time spent:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   585
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   870
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<unknown>"
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   7
         Top             =   285
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<unknown>"
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   6
         Top             =   585
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<unknown>"
         Height          =   195
         Index           =   2
         Left            =   1800
         TabIndex        =   5
         Top             =   870
         Width           =   840
      End
   End
   Begin VB.TextBox Text1 
      Height          =   320
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1720
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   320
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Text            =   "C:\Saol.txt"
      Top             =   1025
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":0024
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   15
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Checksum value (HEX):"
      Height          =   195
      Index           =   5
      Left            =   135
      TabIndex        =   14
      Top             =   1485
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File/Text:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   13
      Top             =   800
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Checksum Algorithm:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_CRC As clsCRC

Private Sub Command1_Click()

  Dim OldTimer As Single
  
'  On Error GoTo ErrorHandler
  
  'Reset the labels
  Label2(0).Caption = "<unknown>"
  Label2(1).Caption = "<unknown>"
  Label2(2).Caption = "<unknown>"
  
  'Select the algorithm from the combobox
  m_CRC.Algorithm = Combo1.ListIndex
  
  'If the text fields contain filenames we
  'want to calculate the CRC of the file given
  If (Mid$(Text1(0).Text, 2, 2) = ":\") Then
    Label2(0).Caption = FileLen(Text1(0).Text) & " bytes"
    OldTimer = Timer
    Text1(1).Text = Hex(m_CRC.CalculateFile(Text1(0).Text))
    Label2(1).Caption = Format$(Timer - OldTimer, "#0.00") & " s"
    Label2(2).Caption = Format$(FileLen(Text1(0).Text) / (Timer - OldTimer) / 1000000, "#0.00 MB/s")
    Call MsgBox("File CRC calculation successful.")
    Exit Sub
  End If

  'Calculate the CRC of the first textbox and
  'store the CRC value in the second textbox,
  'we do not time this because we will get
  'a *very* high value
  Label2(0).Caption = Len(Text1(0).Text)
  Text1(1).Text = Hex(m_CRC.CalculateString(Text1(0).Text))
  Exit Sub
  
ErrorHandler:
  Call MsgBox("Hrmm.. something went terribly wrong." & vbCrLf & vbCrLf & Err.Description, vbExclamation)

End Sub

Private Sub Command2_Click()

  Dim a As Long
  Dim CRC As Long
  Dim ByteArray() As Byte
  Dim TimerAction As Single
  
  'Select the algorithm from the combobox
  m_CRC.Algorithm = Combo1.ListIndex
  
  'Create a *large* bytearray (10MB here, you
  'might want to change this if you have a
  'really fast computer)
  ReDim ByteArray(0 To 9999999)
  
  'Calculate the CRC of the bytearray (with timer)
  TimerAction = Timer
  Call m_CRC.CalculateBytes(ByteArray())
  TimerAction = (Timer - TimerAction)
  
  'Show the result, and amaze the user
  Call MsgBox("Benchmark successful" & vbCrLf & vbCrLf & "Calculation speed: " & Format$(10000000 / TimerAction / 1000000, "#0.00 MB/s"))
  
End Sub


Private Sub Command3_Click()

End Sub


Private Sub Form_Load()

  'Create the CRC object
  Set m_CRC = New clsCRC
  
  'Preselect the CRC32 calculation
  Combo1.ListIndex = 1
  
End Sub

