VERSION 5.00
Begin VB.Form frmIcon 
   Caption         =   "Extract Icon from an exe"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   2700
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   3480
      Width           =   435
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   2520
      Pattern         =   "*.exe"
      TabIndex        =   2
      Top             =   540
      Width           =   2355
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   195
      Index           =   1
      Left            =   2760
      TabIndex        =   5
      Top             =   3120
      Width           =   315
   End
   Begin VB.Shape Shape1 
      Height          =   675
      Left            =   2580
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select an exe to get associated Icon "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   195
      Index           =   0
      Left            =   2280
      TabIndex        =   4
      Top             =   180
      Width           =   2670
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const DI_MASK = &H1
Const DI_IMAGE = &H2
Const DI_NORMAL = DI_MASK Or DI_IMAGE

Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long



Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo err
    Dir1.Path = Drive1.Drive
    Exit Sub
err:
    MsgBox "Error in opening drive.Please try again", vbInformation
End Sub

Private Sub File1_Click()
On Error GoTo err
    Dim hIconHandle     As Long
    Dim sPath           As String
    
    sPath = File1.Path & "\" & File1.FileName
    
    hIconHandle = ExtractAssociatedIcon(Me.hWnd, sPath, 0)
    
    'if call is success then icon handle will be obtained
     
    p.Cls
    If hIconHandle Then
        Dim div As Long
        p.ScaleMode = vbPixels
        div = Screen.TwipsPerPixelX
        
        ScaleMode = p
        DrawIconEx p.hdc, 0, 0, hIconHandle, p.Width / div, p.Height / div, 0, 0, DI_NORMAL
        
        DestroyIcon hIconHandle
    Else
        MsgBox "Unable to extract the Associated Icon.", vbInformation
    End If
    
    
Exit Sub
err:
End Sub
