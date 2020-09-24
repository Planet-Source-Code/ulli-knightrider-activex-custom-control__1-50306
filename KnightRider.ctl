VERSION 5.00
Begin VB.UserControl KnightRider 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   ForeColor       =   &H00000000&
   MaskColor       =   &H00000000&
   ScaleHeight     =   14
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   100
   Begin VB.Timer tmrTick 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   285
      Top             =   -75
   End
   Begin VB.PictureBox picBlend 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   180
      Left            =   -585
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   86
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   1320
   End
End
Attribute VB_Name = "KnightRider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Two alternative blend functions; use the other if you have problems
 Private Declare Function Blend Lib "gdi32.dll" Alias "GdiAlphaBlend" (ByVal desthDC As Long, ByVal destX As Long, ByVal destX As Long, ByVal destW As Long, ByVal destH As Long, ByVal srchDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcW As Long, ByVal srcH As Long, ByVal BLENDFUNCT As Long) As Long
'Private Declare Function Blend Lib "msimg32" Alias "AlphaBlend" (ByVal desthDC As Long, ByVal destX As Long, ByVal destX As Long, ByVal destW As Long, ByVal destH As Long, ByVal srchDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcW As Long, ByVal srcH As Long, ByVal BLENDFUNCT As Long) As Long

Public Enum eEffect
    LeftToRight = 1
    RightToLeft = 2
    Oscillating = 3
End Enum
#If False Then
Private LeftToRight, RightToLeft, Oscillating
#End If
Private myEffect    As eEffect 'Effect in use
Private myTail      As Long
Private mySpeed     As Long 'Speed
Attribute mySpeed.VB_VarDescription = "Step width."
Private Position    As Long 'Current drawing position
Attribute Position.VB_VarDescription = "Current drawing position."
Private w           As Long
Private h           As Long
Private Const pnBC  As String = "BackColor"
Private Const pnFC  As String = "ForeColor"
Private Const pnEF  As String = "Effect"
Private Const pnEN  As String = "Enabled"
Private Const pnSP  As String = "Speed"
Private Const pnTL  As String = "Tail"
Private Const erEF  As String = "No such Effect"
Private Const erZR  As String = "Illegal Zero Value"

Public Property Let BackColor(ByVal nuBackColor As OLE_COLOR)
Attribute BackColor.VB_Description = "Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest."

    picBlend.BackColor = nuBackColor
    UserControl.BackColor = nuBackColor
    PropertyChanged pnBC

End Property

Public Property Get BackColor() As OLE_COLOR

    BackColor = picBlend.BackColor

End Property

Public Property Get Effect() As eEffect
Attribute Effect.VB_Description = "Sets/returns the effect."

    Effect = myEffect

End Property

Public Property Let Effect(ByVal nuEffect As eEffect)

    If nuEffect = LeftToRight Or nuEffect = RightToLeft Or nuEffect = Oscillating Then
        myEffect = nuEffect
        Select Case myEffect
          Case LeftToRight
            Position = 0
            mySpeed = Abs(mySpeed)
          Case RightToLeft
            Position = ScaleWidth
            mySpeed = -Abs(mySpeed)
          Case Else
            Position = ScaleWidth / 2
        End Select
        PropertyChanged pnEF
      Else 'NOT NUEFFECT...
        Err.Raise 381, Me, erEF
    End If

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Sets/returns whether the control is enabled."
Attribute Enabled.VB_UserMemId = 0
Attribute Enabled.VB_MemberFlags = "200"

    Enabled = tmrTick.Enabled

End Property

Public Property Let Enabled(ByVal nuEnabled As Boolean)

    tmrTick.Enabled = (nuEnabled <> False)
    If tmrTick.Enabled = False Then
        Refresh
    End If
    PropertyChanged pnEN

End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gibt die Vordergrundfarbe zurück, die zum Anzeigen von Text und Grafiken in einem Objekt verwendet wird, oder legt diese fest."

    ForeColor = picBlend.ForeColor

End Property

Public Property Let ForeColor(ByVal nuForeColor As OLE_COLOR)

    picBlend.ForeColor() = nuForeColor
    PropertyChanged pnFC

End Property

Public Property Get Speed() As Long
Attribute Speed.VB_Description = "Sets/returns the speed. Usable values are 1 thru 10."

    Speed = Abs(mySpeed)

End Property

Public Property Let Speed(ByVal nuSpeed As Long)

    If nuSpeed Then
        mySpeed = nuSpeed * Sgn(mySpeed)
        PropertyChanged pnSP
      Else 'NUSPEED = FALSE/0
        Err.Raise 382, Me, erZR
    End If

End Property

Public Property Get Tail() As Long
Attribute Tail.VB_Description = "Sets/returns the tail length."

    Tail = 31 - myTail \ 65536

End Property

Public Property Let Tail(ByVal nuTail As Long)

    myTail = (31 - (nuTail And 31)) * 65536
    PropertyChanged pnTL

End Property

Private Sub tmrTick_Timer()

    Blend hDC, 0, 0, w, h, picBlend.hDC, 0, 0, w, h, myTail
    Line (Position, 0)-(Position + mySpeed - 1, h - 1), picBlend.ForeColor, BF
    Position = Position + mySpeed
    Select Case myEffect
      Case LeftToRight
        If Position > w Then
            Position = 0
        End If
      Case RightToLeft
        If Position < 0 Then
            Position = w
        End If
      Case Oscillating
        If Position < 0 Or Position > w Then
            mySpeed = -mySpeed
        End If
    End Select

End Sub

Private Sub UserControl_InitProperties()

    mySpeed = 2
    Effect = Oscillating
    Tail = 15

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        BackColor = .ReadProperty(pnBC, vbBlack)
        picBlend.ForeColor = .ReadProperty(pnFC, vbGreen)
        tmrTick.Enabled = .ReadProperty(pnEN, False)
        myEffect = .ReadProperty(pnEF, Oscillating)
        mySpeed = .ReadProperty(pnSP, 2)
        Tail = .ReadProperty(pnTL, 15)
    End With 'PROPBAG

End Sub

Private Sub UserControl_Resize()

    picBlend.Move 0, 0, Width, Height
    w = ScaleWidth
    h = ScaleHeight
    If myEffect Then 'this is here to avoid an error while effect is unknown
        Effect = myEffect
    End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty pnBC, picBlend.BackColor, vbBlack
        .WriteProperty pnFC, picBlend.ForeColor, vbGreen
        .WriteProperty pnEN, tmrTick.Enabled, False
        .WriteProperty pnEF, myEffect, Oscillating
        .WriteProperty pnSP, Abs(mySpeed), 2
        .WriteProperty pnTL, Tail, 15
    End With 'PROPBAG

End Sub

':) Ulli's VB Code Formatter V2.16.13 (2003-Dez-04 10:20) 28 + 169 = 197 Lines
