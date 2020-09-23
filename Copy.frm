VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copy File"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CmnDlg 
      Left            =   6000
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton CmdBrowse 
         Caption         =   "Browse"
         Height          =   255
         Left            =   5400
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtSrc 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox TxtDest 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Text            =   "c:\temp.tmp"
         Top             =   600
         Width           =   4815
      End
      Begin VB.CommandButton CmdCopy 
         Caption         =   "Copy"
         Default         =   -1  'True
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   6015
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label3 
         Caption         =   "%"
         Height          =   255
         Left            =   5880
         TabIndex        =   9
         Top             =   1755
         Width           =   255
      End
      Begin VB.Label LblPercent 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   5280
         TabIndex        =   8
         Top             =   1755
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Source       :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Destination :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************
'Coded by   : Nabhan Ahmed
'Date       : 8-25-2003
'Description: This program shows you how to copy a file byte by byte.
'             It reads 4 kbs from the source file and write them in
'             the destination file until it reads all byte in the source
'             file. There a bar that shows the copying progress.
'*************************************************************************

Private Sub CmdBrowse_Click()
'Set filter to all files
CmnDlg.Filter = "All files|*.*"
'Show the open file window
CmnDlg.ShowOpen
'Put the file bath in the source text
TxtSrc.Text = CmnDlg.FileName
End Sub

Private Sub CmdCopy_Click()
On Error GoTo CopyErr

'Declare variables
Dim SrcFile As String
Dim DestFile As String
Dim SrcFileLen As Long
Dim nSF, nDF As Integer
Dim Chunk As String
Dim BytesToGet As Integer
Dim BytesCopied As Long

'Disable the copy button
CmdCopy.Enabled = False

'The source file the you want to copy
SrcFile = TxtSrc
'The destination file name
DestFile = TxtDest

'Get source file length
SrcFileLen = FileLen(SrcFile)
'Progress bar settings
ProgressBar1.Min = 0
ProgressBar1.Max = SrcFileLen

'Open both files
nSF = 1
nDF = 2
Open SrcFile For Binary As nSF
Open DestFile For Binary As nDF

'How many bytes to get each time
BytesToGet = 4096 '4kb
BytesCopied = 0
'Show Progress
ProgressBar1.Value = 0
'Show percentage
LblPercent.Caption = "0"
'ProgressBar1.Visible = True

'Keep copying until finishing all bytes
Do While BytesCopied < SrcFileLen
    'Check how many bytes left
    If BytesToGet < (SrcFileLen - BytesCopied) Then
        'Copy 4 KBytes
        Chunk = Space(BytesToGet)
        Get #nSF, , Chunk
    Else
        'Copy the rest
        Chunk = Space(SrcFileLen - BytesCopied)
        Get #nSF, , Chunk
    End If
    BytesCopied = BytesCopied + Len(Chunk)
    
    'Show progress
    ProgressBar1.Value = BytesCopied
    'Show Percentage
    LblPercent.Caption = Int(BytesCopied / SrcFileLen * 100)
    LblPercent.Refresh
        
    'Put data in destination file
    Put #nDF, , Chunk
Loop

'Hide progress bar
ProgressBar1.Value = 0
'ProgressBar1.Visible = False

'Skip the error message and exit sub
GoTo Ex

CopyErr:
MsgBox Err.Description, vbCritical, "Error"

Ex:
'Close files
Close #nSF
Close #nDF

'Re-enable the copy button
CmdCopy.Enabled = True
End Sub

