VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "3D StarField"
   ClientHeight    =   4215
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4695
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   FillStyle       =   0  'Solid
   Icon            =   "frmMain.frx":0000
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMeteor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   120
      Picture         =   "frmMain.frx":000C
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.PictureBox vBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   135
      Left            =   120
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Visible         =   0   'False
      Begin VB.Menu mnuStarCount 
         Caption         =   "Star&Count.."
      End
      Begin VB.Menu mnuSpeed 
         Caption         =   "&Speed.."
      End
      Begin VB.Menu mnuDist 
         Caption         =   "&Distance.."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrails 
         Caption         =   "&Trails"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuMax 
         Caption         =   "&Maximize"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    If App.PrevInstance = True Then End     'Only one copy allowed. (ScreenSaver)
    
    frmMain.Visible = True                  'Show the form before the loop starts.
    
    StarSpeed = 100     'Star movement speed. Used to smooth out movement.
    StarCount = 200     'Number of stars in the field.
    MaxDist = 10        'Farthest starting point possible. Set to your liking.
    TickSpeed = 50      'Loop speed in ms.
    
    Do
        CurrTick = GetTickCount
        If CurrTick - LastTick >= TickSpeed Then
            UpdateStars     'Calculate stars new positions.
            RedrawStars     'Redraw the stars.
            If DrawButton = True Then RedrawButton      'If cursors at bottom left, draw.
            
            LastTick = CurrTick
        End If
        
        DoEvents
    Loop
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Functional button clicked, bring up menu.
    If x < 30 And y > (frmMain.ScaleHeight - 15) Then PopupMenu mnuSettings, , 0, frmMain.ScaleHeight
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Check mouse position for function button on/off.
    If x < 30 And y > (frmMain.ScaleHeight - 15) Then
        DrawButton = True
    Else
        DrawButton = False
    End If
End Sub

Private Sub Form_Resize()
    'Resize the video buffer to the new form size.
    'Set scalemodes to twips to fix sizing problems, then reset to pixels.
    vBuffer.ScaleMode = vbTwips
    frmMain.ScaleMode = vbTwips
    vBuffer.Move 0, 0, Me.Width, Me.Height
    vBuffer.ScaleMode = vbPixels
    frmMain.ScaleMode = vbPixels
    
    'Midpoints used to calculate X/Y from X/Y/Z
    MidX = vBuffer.Width / 2
    MidY = vBuffer.Height / 2
End Sub

'Menu items
Private Sub mnuDist_Click()
    MaxDist = InputBox("Distance of star origin:", "Distance", MaxDist)
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuMax_Click()
    mnuMax.Checked = Not mnuMax.Checked
    
    If mnuMax.Checked = True Then
        frmMain.WindowState = vbMaximized
    Else
        frmMain.WindowState = vbNormal
    End If
End Sub

Private Sub mnuSpeed_Click()
    StarSpeed = InputBox("Speed of stars: ", "Speed", StarSpeed)
End Sub

Private Sub mnuStarCount_Click()
    StarCount = InputBox("Number of stars in the field:", "StarCount", StarCount)
End Sub

Private Sub mnuTrails_Click()
    mnuTrails.Checked = Not mnuTrails.Checked
End Sub
