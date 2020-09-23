VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MIDAR's HSV to RGB Colour Converter"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pictGradient 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3555
      Index           =   2
      Left            =   8640
      ScaleHeight     =   3495
      ScaleWidth      =   3495
      TabIndex        =   32
      ToolTipText     =   "Fixed Hue. Saturation vs. Lightness"
      Top             =   780
      Width           =   3555
   End
   Begin VB.PictureBox pictGradient 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1155
      Index           =   1
      Left            =   6210
      ScaleHeight     =   1095
      ScaleWidth      =   1155
      TabIndex        =   31
      ToolTipText     =   "Fixed Hue. Saturation vs. Lightness"
      Top             =   1980
      Width           =   1215
   End
   Begin VB.PictureBox pictColourWheel 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3555
      Left            =   5040
      ScaleHeight     =   3495
      ScaleWidth      =   3495
      TabIndex        =   30
      Top             =   780
      Width           =   3555
   End
   Begin VB.PictureBox pictGradient 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1155
      Index           =   0
      Left            =   3780
      ScaleHeight     =   1095
      ScaleWidth      =   1155
      TabIndex        =   27
      ToolTipText     =   "Fixed Hue. Saturation vs. Lightness"
      Top             =   1980
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   26
      Top             =   4875
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pictGrid 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1155
      Left            =   3780
      ScaleHeight     =   1095
      ScaleWidth      =   1155
      TabIndex        =   25
      ToolTipText     =   "Fixed Saturation (equals 1). Variable Hue and Lightness."
      Top             =   3180
      Width           =   1215
   End
   Begin VB.PictureBox pictRGB 
      Height          =   1155
      Left            =   3780
      ScaleHeight     =   1095
      ScaleWidth      =   1155
      TabIndex        =   24
      ToolTipText     =   "Current Colour"
      Top             =   780
      Width           =   1215
   End
   Begin VB.Frame FrameCurrentValues 
      Caption         =   "Current Values (read only)"
      Height          =   1995
      Left            =   60
      TabIndex        =   9
      Top             =   2340
      Width           =   3615
      Begin VB.TextBox txtHTMLCode 
         Height          =   285
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "#FFFFFF"
         Top             =   1560
         Width           =   915
      End
      Begin VB.TextBox txtBlue 
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   15
         ToolTipText     =   "Any value between 0 and 255."
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox txtGreen 
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Any value between 0 and 255."
         Top             =   660
         Width           =   915
      End
      Begin VB.TextBox txtRed 
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   13
         ToolTipText     =   "Any value between 0 and 255."
         Top             =   300
         Width           =   915
      End
      Begin VB.TextBox txtLightness 
         Height          =   285
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Any value between 0 and 1."
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox txtSaturation 
         Height          =   285
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Any value between 0 and 1."
         Top             =   660
         Width           =   915
      End
      Begin VB.TextBox txtHue 
         Height          =   285
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Any value between 0 and 360."
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "HTML Code"
         Height          =   195
         Index           =   9
         Left            =   1605
         TabIndex        =   23
         Top             =   1620
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hue"
         Height          =   195
         Index           =   8
         Left            =   2220
         TabIndex        =   21
         Top             =   360
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Saturation"
         Height          =   195
         Index           =   7
         Left            =   1800
         TabIndex        =   20
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lightness"
         Height          =   195
         Index           =   6
         Left            =   1845
         TabIndex        =   19
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Red"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Green"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   17
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Blue"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   16
         Top             =   1080
         Width           =   315
      End
   End
   Begin VB.Frame FrameHSV 
      Caption         =   "HSV"
      Height          =   1455
      Left            =   60
      TabIndex        =   4
      Top             =   720
      Width           =   3615
      Begin VB.CheckBox chkHue 
         Caption         =   "Hue"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         ToolTipText     =   "Click to Toggle Black & White Mode."
         Top             =   308
         Value           =   1  'Checked
         Width           =   615
      End
      Begin MSComctlLib.Slider SliderSaturation 
         Height          =   270
         Left            =   960
         TabIndex        =   5
         ToolTipText     =   "Saturation: Any value between 0 and 1"
         Top             =   660
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   476
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelStart        =   100
         TickFrequency   =   10
         Value           =   100
      End
      Begin MSComctlLib.Slider SliderLightness 
         Height          =   270
         Left            =   960
         TabIndex        =   6
         ToolTipText     =   "Lightness: Any value between 0 and 1"
         Top             =   1020
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   476
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelStart        =   100
         TickFrequency   =   10
         Value           =   100
      End
      Begin MSComctlLib.Slider SliderHue 
         Height          =   270
         Left            =   960
         TabIndex        =   28
         ToolTipText     =   "Hue: Any value between 0 and 360 inclusive."
         Top             =   300
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   476
         _Version        =   393216
         LargeChange     =   60
         Max             =   360
         TickFrequency   =   90
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Saturation"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   8
         Top             =   705
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lightness"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   7
         Top             =   1065
         Width           =   675
      End
   End
   Begin VB.CommandButton btnClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   7380
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.PictureBox pictHeader 
      Align           =   1  'Align Top
      BackColor       =   &H80000005&
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8580
      TabIndex        =   0
      Top             =   0
      Width           =   8640
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "by Peter Wilson - http://dev.midar.com/"
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   2
         Top             =   360
         Width           =   2820
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "MIDAR's HSV to RGB Colour Converter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   780
         TabIndex        =   1
         Top             =   0
         Width           =   5460
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "Form1.frx":000C
         Top             =   60
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DrawColourWheel(ByVal HueOffSet As Single, ByVal Saturation As Single, ByVal Lightness As Single, ByStep As Single)

    Me.pictColourWheel.ScaleHeight = 100
    Me.pictColourWheel.ScaleWidth = 100
    Me.pictColourWheel.ScaleLeft = -50
    Me.pictColourWheel.ScaleTop = -50
    
    Dim sngHueTemp As Single
    Dim sngDeg As Single
    Dim sngRadians As Single
    Dim pixelX As Single, pixelX2 As Single
    Dim pixelY As Single, pixelY2 As Single
    
    Dim sngRadius As Single, sngTickLength As Single
    
    sngRadius = 40              ' << Change this for fun.
    sngTickLength = 5           ' << Change this for fun.
    Me.pictColourWheel.DrawWidth = 12  ' << Change this for fun.
    
    ' Clear the Picture Box
    Me.pictColourWheel.Cls
    
    ' Draw the colouor wheel
    ' ======================
    For sngDeg = 0 To 360 Step ByStep
        ' Convert Degrees to Radians
        sngRadians = sngDeg * (3.141593 / 180)
        
        If HueOffSet = -1 Then
            sngHueTemp = -1
        Else
            sngHueTemp = sngDeg + HueOffSet + 180
        End If
        Me.pictColourWheel.ForeColor = HSV(sngHueTemp, Saturation, Lightness)
        
        pixelX = Cos(sngRadians) * (sngRadius + sngTickLength)
        pixelY = Sin(sngRadians) * (sngRadius + sngTickLength)
        pixelX2 = Cos(sngRadians) * (sngRadius - sngTickLength)
        pixelY2 = Sin(sngRadians) * (sngRadius - sngTickLength)
        
        If sngDeg = 0 Then
            ' Just move to the point on the first time around.
            Me.pictColourWheel.Line (pixelX, pixelY)-(pixelX, pixelY)
        Else
            Me.pictColourWheel.Line (pixelX, pixelY)-(pixelX2, pixelY2)
        End If
    Next sngDeg
    
    
    
    ' *****************************************************************************
    ' Everything else in the sub, after this point is OPTIONAL - you can delete it.
    ' *****************************************************************************
    If ByStep > 4 Then Exit Sub
    
    
    ' Smooth the outside edge of the colour wheel by adjusting the lightness value.
    ' Note: This looks really cool offest -30 degrees
    ' =============================================================================
    Me.pictColourWheel.DrawWidth = 5
    For sngDeg = 0 To 360 Step ByStep
        sngRadians = sngDeg * (3.141593 / 180)
        If HueOffSet = -1 Then
            sngHueTemp = -1
        Else
            sngHueTemp = sngDeg + HueOffSet + 150
        End If
        Me.pictColourWheel.ForeColor = HSV(sngHueTemp, Saturation, Lightness * 0.7)
        pixelX = Cos(sngRadians) * (sngRadius + sngTickLength + 2)
        pixelY = Sin(sngRadians) * (sngRadius + sngTickLength + 2)
        If sngDeg = 0 Then
            Me.pictColourWheel.PSet (pixelX, pixelY)
        Else
            Me.pictColourWheel.Line -(pixelX, pixelY)
        End If
    Next sngDeg
    
    ' Smooth the inside edge of the colour wheel by adjusting the lightness value.
    ' Note: This looks really cool offest +30 degrees
    ' =============================================================================
    Me.pictColourWheel.DrawWidth = 5
    For sngDeg = 0 To 360 Step ByStep
        sngRadians = sngDeg * (3.141593 / 180)
        If HueOffSet = -1 Then
            sngHueTemp = -1
        Else
            sngHueTemp = sngDeg + HueOffSet + 210
        End If
        Me.pictColourWheel.ForeColor = HSV(sngHueTemp, Saturation, Lightness * 0.7)
        pixelX = Cos(sngRadians) * (sngRadius - sngTickLength - 2)
        pixelY = Sin(sngRadians) * (sngRadius - sngTickLength - 2)
        If sngDeg = 0 Then
            Me.pictColourWheel.PSet (pixelX, pixelY)
        Else
            Me.pictColourWheel.Line -(pixelX, pixelY)
        End If
    Next sngDeg
    
    
End Sub

Private Sub DrawColourGrid(ByVal HueOffSet As Single, ByVal Saturation As Single, ByVal Lightness As Single)

    Me.pictGrid.ScaleHeight = 360
    Me.pictGrid.ScaleWidth = 1
    Me.pictGrid.ScaleLeft = 0
    Me.pictGrid.ScaleTop = 0
    Me.pictGrid.Cls
    
    Dim sngHue As Single, sngHueTemp As Single
    Dim sngLightness As Single
    
    Dim sngX1 As Single, sngY1 As Single
    Dim sngX2 As Single, sngY2 As Single
    
    Dim sngStep As Single
    
    sngStep = 360 / 6
    
    For sngHue = 0 To 360 Step sngStep
        For sngLightness = 0 To 1 Step (1 / 6)
        
            sngX1 = sngLightness: sngY1 = sngHue
            sngX2 = sngX1 + (1 / 6): sngY2 = sngY1 + sngStep
            
            If HueOffSet = -1 Then
                sngHueTemp = -1
            Else
                sngHueTemp = sngHue + HueOffSet
            End If
        
            Me.pictGrid.ForeColor = HSV(sngHueTemp, Saturation, 1 - sngLightness)
            
            Me.pictGrid.Line (sngX1, sngY1)-(sngX2, sngY2), , BF
            
        Next sngLightness
    Next sngHue
    
End Sub

Private Sub DrawColourGradient(ByVal Hue As Single, ByVal Index As Integer, ByVal sngStep As Single)
    
    Dim sngSaturation As Single
    Dim sngLightness As Single
    
    Dim sngX1 As Single, sngY1 As Single
    Dim sngX2 As Single, sngY2 As Single
        
    Me.pictGradient(Index).ScaleHeight = 1
    Me.pictGradient(Index).ScaleWidth = 1
    Me.pictGradient(Index).ScaleLeft = 0
    Me.pictGradient(Index).ScaleTop = 0
    If Index <> 2 Then Me.pictGradient(Index).Cls
    
    For sngSaturation = 0 To (1 - sngStep) Step sngStep
        For sngLightness = 0 To 1 Step sngStep
        
            sngX1 = sngSaturation: sngY1 = sngLightness
            sngX2 = sngX1 + sngStep: sngY2 = sngY1 + sngStep
            
            Me.pictGradient(Index).ForeColor = HSV(Hue, 1 - sngSaturation, 1 - sngLightness)
            Me.pictGradient(Index).Line (sngX1, sngY1)-(sngX2, sngY2), , BF
                        
        Next sngLightness
        If Index = 2 Then Me.pictGradient(Index).Refresh
    Next sngSaturation

End Sub

Private Sub btnClose_Click()

    Unload Me
    
End Sub

Private Sub chkHue_Click()
    
    ' It's actually possible to have a null Hue value (ie. Black and White)
    If Me.chkHue.Value = vbChecked Then
        Me.SliderHue.Enabled = True
        Me.SliderSaturation.Value = 100
        Me.SliderSaturation.Enabled = True
    Else
        Me.SliderHue.Enabled = False
        Me.SliderSaturation.Value = 0
        Me.SliderSaturation.Enabled = False
    End If
    
    Call DoUpdateValues_Slow
    
End Sub

Private Sub SliderHue_Change()
    Call DoUpdateValues_Slow
End Sub


Private Sub SliderHue_Scroll()
    Call DoUpdateValues_Quick
End Sub

Private Sub SliderLightness_Change()
    Call DoUpdateValues_Slow
End Sub

Private Sub SliderLightness_Scroll()
    Call DoUpdateValues_Quick
End Sub

Private Sub SliderSaturation_Change()
    Call DoUpdateValues_Slow
End Sub

Private Sub SliderSaturation_Scroll()
    Call DoUpdateValues_Quick
End Sub

Public Sub DoUpdateValues_Slow()

    On Error GoTo errTrap
    
    Dim sngHue As Single
    Dim sngSaturation As Single
    Dim sngLightness As Single
    
    Dim sngRed As Single
    Dim sngGreen As Single
    Dim sngBlue As Single
    
    Screen.MousePointer = vbHourglass

    If Me.SliderHue.Enabled = True Then
        sngHue = Me.SliderHue.Value
    Else
        sngHue = -1 ' Black & White mode
    End If
    
    sngSaturation = (Me.SliderSaturation.Value / 100)   ' << Sliders only work in Integers, so this is an easy workaround.
    sngLightness = (Me.SliderLightness.Value / 100)     ' << Sliders only work in Integers, so this is an easy workaround.
    
    
    ' This is how to use the HSV routine as a direct replacement for RGB.
    ' ===================================================================
    Me.pictRGB.BackColor = HSV(sngHue, sngSaturation, sngLightness)
    
    
    ' Here is an alternative method of using the HSV routine, to return the separate Red, Green, Blue values.
    ' =======================================================================================================
    Call HSV2(sngRed, sngGreen, sngBlue, sngHue, sngSaturation, sngLightness)
    
    
    ' Draw the colour wheel (just for fun)
    Call DrawColourWheel(sngHue, sngSaturation, sngLightness, 1)
    
    ' Draw a Colour Gradient (just for fun)
    Call DrawColourGradient(sngHue, 0, (1 / 8))
    Call DrawColourGradient(sngHue, 1, (1 / 96))
    
    ' Draw a Colour Grid (just for fun)
    Call DrawColourGrid(sngHue, sngSaturation, sngLightness)
    
    
    ' Update Text Boxes
    ' =================
    Me.txtRed.Text = Int(sngRed * 255)
    Me.txtGreen.Text = Int(sngGreen * 255)
    Me.txtBlue.Text = Int(sngBlue * 255)
    
    Me.txtHTMLCode.Text = RGBHTML(sngRed, sngGreen, sngBlue)
    
    Me.StatusBar1.SimpleText = ""
    
ResumePoint:
    Me.txtHue.Text = sngHue
    Me.txtSaturation.Text = sngSaturation
    Me.txtLightness.Text = sngLightness
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
errTrap:
    Me.txtRed.Text = ""
    Me.txtGreen.Text = ""
    Me.txtBlue.Text = ""
    Me.txtHTMLCode.Text = ""
    
    Me.pictRGB.BackColor = &H8000000F   ' << System Colour - Button Face
    Me.pictColourWheel.Cls
    Me.pictGradient(0).Cls
    Me.pictGradient(1).Cls
    Me.pictGrid.Cls
    
    Me.StatusBar1.SimpleText = Err.Description
    Resume ResumePoint
    
End Sub
Public Sub DoUpdateValues_Quick()

    On Error GoTo errTrap
    
    Dim sngHue As Single
    Dim sngSaturation As Single
    Dim sngLightness As Single
    
    Dim sngRed As Single
    Dim sngGreen As Single
    Dim sngBlue As Single
    
    Screen.MousePointer = vbHourglass
    
    If Me.SliderHue.Enabled = True Then
        sngHue = Me.SliderHue.Value
    Else
        sngHue = -1 ' Black & White mode
    End If

    sngSaturation = (Me.SliderSaturation.Value / 100)   ' << Sliders only work in Integers, so this is an easy workaround.
    sngLightness = (Me.SliderLightness.Value / 100)     ' << Sliders only work in Integers, so this is an easy workaround.
    
    
    ' This is how to use the HSV routine as a direct replacement for RGB.
    ' ===================================================================
    Me.pictRGB.BackColor = HSV(sngHue, sngSaturation, sngLightness)
    
    
    ' Here is an alternative method of using the HSV routine, to return the separate Red, Green, Blue values.
    ' =======================================================================================================
    Call HSV2(sngRed, sngGreen, sngBlue, sngHue, sngSaturation, sngLightness)
    
    
    ' Draw the colour wheel (just for fun)
    Call DrawColourWheel(sngHue, sngSaturation, sngLightness, 5)
    
    ' Clear Colour Gradient (just for fun)
    Call DrawColourGradient(sngHue, 0, (1 / 2))
    Call DrawColourGradient(sngHue, 1, (1 / 4))
    
    ' Draw a Colour Grid (just for fun)
    Call DrawColourGrid(sngHue, sngSaturation, sngLightness)
    
    
    ' Update Text Boxes
    ' =================
    Me.txtRed.Text = Int(sngRed * 255)
    Me.txtGreen.Text = Int(sngGreen * 255)
    Me.txtBlue.Text = Int(sngBlue * 255)
    
    Me.txtHTMLCode.Text = RGBHTML(sngRed, sngGreen, sngBlue)
    
    Me.StatusBar1.SimpleText = ""
    
ResumePoint:
    Me.txtHue.Text = sngHue
    Me.txtSaturation.Text = sngSaturation
    Me.txtLightness.Text = sngLightness
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
errTrap:
    Me.txtRed.Text = ""
    Me.txtGreen.Text = ""
    Me.txtBlue.Text = ""
    Me.txtHTMLCode.Text = ""
    
    Me.pictRGB.BackColor = &H8000000F   ' << System Colour - Button Face
    Me.pictColourWheel.Cls
    Me.pictGradient(0).Cls
    Me.pictGradient(1).Cls
    
    Me.pictGrid.Cls
    
    Me.StatusBar1.SimpleText = Err.Description
    Resume ResumePoint
    
End Sub

