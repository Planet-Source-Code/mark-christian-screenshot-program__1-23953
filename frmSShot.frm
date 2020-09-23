VERSION 5.00
Begin VB.Form frmSShot 
   AutoRedraw      =   -1  'True
   Caption         =   "Screenshot"
   ClientHeight    =   4710
   ClientLeft      =   1725
   ClientTop       =   1545
   ClientWidth     =   2550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSShot.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   2550
   Begin VB.CommandButton cmdTake 
      Caption         =   "Take screenshot"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtFN 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "screenshot.bmp"
      Top             =   3360
      Width           =   2295
   End
   Begin VB.DirListBox Folder 
      Height          =   2340
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblFN 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lblSave 
      Caption         =   "Save in:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmSShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code (C) 2001 by Mark Christian
'Email: mark.christian@bigfoot.com
'Web: http://nexxus.dhs.org

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function GetImage(OutputBitmap)
'Note that wherever it says Me.[anything],
'ME can be changed to FORMNAME.[anything].
'Me is a shortcut to the current form.
'Note that the form's AutoRedraw property must be true.

'This pauses the computer to give it time to hide
'the screenshot program. Without it, frmSShot would
'appear in the screenshot.
Sleep 100
DoEvents 'This refreshes after the delay

'Declare variables
Dim wHand As Long
Dim wDC As Long
Dim nHeight As Long, nWidth As Long

wHand = GetDesktopWindow 'Get the desktop's hWnd
wDC = GetDC(wHand) 'Convert hWnd to hDC

'Get screen resolution
nHeight = Screen.Height / Screen.TwipsPerPixelY
nWidth = Screen.Width / Screen.TwipsPerPixelX

'Take snapshot
BitBlt Me.hDC, 0, 0, nWidth, nHeight, wDC, 0, 0, vbSrcCopy

'Save to file
SavePicture Me.Image, OutputBitmap

'Clear form
Me.Cls
End Function
Private Sub cmdAbout_Click()
MsgBox "Screenshot by Mark Christian." & vbNewLine & "Email: mark.christian@bigfoot.com", 64
End Sub


Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdTake_Click()
Me.Visible = False

If Right(Folder.Path, 1) = "\" Then
    OutFN = Folder.Path & txtFN.Text
Else
    OutFN = Folder.Path & "\" & txtFN.Text
End If

If Dir(OutFN, vbNormal) <> "" Then
    x = MsgBox("File already exists. Overwrite?", vbYesNo + vbExclamation)
    If x = vbNo Then Exit Sub
End If

GetImage OutFN
Me.Visible = True
End Sub


Private Sub Drive_Change()
Folder.Path = Drive.Drive
End Sub


