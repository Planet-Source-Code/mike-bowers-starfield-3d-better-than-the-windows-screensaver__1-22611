Attribute VB_Name = "modStarfield"
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest

Public DrawButton As Boolean

Public Type tStar       'Star data
    x As Double
    y As Double
    Z As Double
    
    isMeteor As Boolean
    Speed As Long
    isAlive As Long
End Type

Public CurrTick As Long
Public LastTick As Long
Public TickSpeed As Long

Public Const MaxStars As Long = 10000
Public StarCount As Long
Public Star(MaxStars) As tStar
Public StarSpeed As Long

Public MaxDist As Long

Public MidX As Long
Public MidY As Long

Public Sub MakeStar(Index As Long)
    With Star(Index)
        If Rnd * 1000 > 999 Then        '1/1000 chance of being a meteor.
            .isMeteor = True
            .Z = MaxDist                'Set all meteors starting point to the farthest possible.
        Else
            .isMeteor = False
            .Z = Rnd * MaxDist + 1
        End If
        
        .x = (Rnd * (MidX * 2)) - MidX  'Set X position
        .y = (Rnd * (MidY * 2)) - MidY  'Set Y position
        
        .Speed = Rnd * MaxDist + 2      'Each star has a variable speed.
        .isAlive = True
    End With
End Sub

Public Sub UpdateStars()
    Dim A As Long
    
    For A = 0 To StarCount
        With Star(A)
            If .isAlive = False Then
                MakeStar A              'It's dead, so make remake it.
            Else
                If .Z <= 0.1 Then       'Restart star, its at end of movement.
                    .isAlive = False
                Else
                    .Z = .Z - .Speed / StarSpeed    'Move star.
                End If
            End If
        End With
        
        DoEvents
    Next A
End Sub

Public Sub RedrawStars()
    Dim A As Long
    Dim Shade As Long
    
    'No trails, clear the buffer.
    If frmMain.mnuTrails.Checked = False Then frmMain.vBuffer.Cls
    
    For A = 0 To StarCount
        With Star(A)
            If .Z = 0 Then .Z = 0.1
            
            Shade = 255 - (255 * (Int(.Z) / MaxDist))       'Figure out shade according to distance.
            If .isMeteor = True Then
                'Meteors are drawn using the picture in picMeteor... size is determined by distance as in real life.
                StretchBlt frmMain.vBuffer.hdc, MidX + (.x / .Z), MidY + (.y / .Z), 20 / (.Z * 2), 20 / (.Z * 2), frmMain.picMeteor.hdc, 40, 0, 40, 40, SRCAND
                StretchBlt frmMain.vBuffer.hdc, MidX + (.x / .Z), MidY + (.y / .Z), 20 / (.Z * 2), 20 / (.Z * 2), frmMain.picMeteor.hdc, 0, 0, 40, 40, SRCCOPY
            Else
                'Regular stars are just pixels.
                SetPixelV frmMain.vBuffer.hdc, MidX + (.x / .Z), MidY + (.y / .Z), RGB(Shade, Shade, Shade)
            End If
        End With
        
        DoEvents
    Next A
    
    'BitBlt everything to the form. Users don't see anything until now.
    BitBlt frmMain.hdc, 0, 0, frmMain.Width, frmMain.Height, frmMain.vBuffer.hdc, 0, 0, SRCCOPY
End Sub

Public Sub RedrawButton()
    'Draws functional button.
    With frmMain
        frmMain.Line (0, .ScaleHeight - 15)-(30, .ScaleHeight), RGB(150, 150, 150), BF
        frmMain.Line (0, .ScaleHeight - 15)-(30, .ScaleHeight - 15), RGB(200, 200, 200)
        frmMain.Line (30, .ScaleHeight - 15)-(30, .ScaleHeight), RGB(200, 200, 200)
    End With
End Sub

