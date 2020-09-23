VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "PolyCube Calendar (dodecahedron)"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9495
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin Project1.XPProgressBar PB1 
      Height          =   285
      Left            =   7155
      TabIndex        =   24
      Top             =   2490
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   503
      Max             =   11
      Step_Length     =   10
      Seperator_Width =   0
      BackColor       =   16777215
      BarColor        =   16777215
   End
   Begin Project1.ucPanel ucPanel3 
      Height          =   1065
      Left            =   7140
      TabIndex        =   20
      Top             =   2895
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "           Options"
      ColorTop        =   12648447
      ColorBottom     =   8454143
      Begin VB.CheckBox chkSaturday 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Color Saturdays"
         Height          =   225
         Left            =   105
         TabIndex        =   23
         Top             =   795
         Width           =   1575
      End
      Begin VB.CheckBox chkSunday 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Color Sundays"
         Height          =   225
         Left            =   105
         TabIndex        =   22
         Top             =   555
         Width           =   1440
      End
      Begin VB.CheckBox chkLatin 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Monday First"
         Height          =   285
         Left            =   105
         TabIndex        =   21
         Top             =   285
         Width           =   1260
      End
   End
   Begin Project1.ucPanel ucPanel2 
      Height          =   1545
      Left            =   6780
      TabIndex        =   15
      Top             =   3990
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   2725
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "                   Controls"
      ColorTop        =   12648447
      ColorBottom     =   8454143
      Begin VB.CommandButton Command5 
         Caption         =   "Save Jpeg"
         Height          =   525
         Left            =   1335
         TabIndex        =   19
         Top             =   330
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save Bitmap"
         Height          =   525
         Left            =   75
         TabIndex        =   18
         Top             =   330
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Print"
         Height          =   525
         Left            =   1335
         TabIndex        =   17
         Top             =   930
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear"
         Height          =   525
         Left            =   75
         TabIndex        =   16
         Top             =   930
         Width           =   1215
      End
   End
   Begin Project1.ucPanel ucPanel1 
      Height          =   1590
      Left            =   6825
      TabIndex        =   8
      Top             =   7740
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   2805
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "                     Logo"
      ColorTop        =   12648447
      ColorBottom     =   8454143
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1335
         Picture         =   "Form1.frx":030A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   14
         Top             =   705
         Width           =   270
      End
      Begin VB.CheckBox chkLogo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Add a Logo"
         Height          =   210
         Left            =   75
         TabIndex        =   11
         Top             =   690
         Width           =   1275
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "(Ex: Just for demo)"
         Height          =   270
         Left            =   75
         TabIndex        =   13
         Top             =   405
         Width           =   2205
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   15
         Left            =   30
         TabIndex        =   12
         Top             =   435
         Width           =   30
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Logo size 16 x 16"
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Top             =   1035
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Click on Logo to Change"
         Height          =   240
         Left            =   60
         TabIndex        =   9
         Top             =   1275
         Width           =   1830
      End
   End
   Begin VB.ComboBox cboYear 
      Height          =   315
      ItemData        =   "Form1.frx":0454
      Left            =   7500
      List            =   "Form1.frx":046D
      TabIndex        =   4
      Text            =   "2009"
      Top             =   90
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make Calendar"
      Height          =   510
      Left            =   7410
      TabIndex        =   1
      Top             =   1905
      Width           =   1275
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   16000
      Left            =   11355
      ScaleHeight     =   1063
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   884
      TabIndex        =   2
      Top             =   -90
      Width           =   13320
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   165
         TabIndex        =   7
         Top             =   15165
         Width           =   1395
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   165
         TabIndex        =   6
         Top             =   14895
         Width           =   2085
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   180
         TabIndex        =   5
         Top             =   14655
         Width           =   1080
      End
   End
   Begin VB.PictureBox PicDate 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   6945
      ScaleHeight     =   131
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   0
      Top             =   5625
      Width           =   2295
   End
   Begin VB.PictureBox picRot 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   10920
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   227
      TabIndex        =   3
      Top             =   4620
      Width           =   3465
   End
   Begin VB.Image Image2 
      Height          =   2790
      Left            =   6690
      Picture         =   "Form1.frx":049B
      Stretch         =   -1  'True
      Top             =   60
      Width           =   2760
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   9285
      Left            =   45
      Stretch         =   -1  'True
      Top             =   60
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FtSize As Integer
Dim ct As Integer

Private Sub Form_Load()
   picMain.Picture = picMain.Image
   Image1.Picture = picMain.Picture
End Sub

Private Sub Form_Resize()
   picRot.ScaleHeight = 181
   picRot.ScaleWidth = 186
   PicDate.ScaleHeight = 131
   PicDate.ScaleWidth = 149
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

Private Sub DrawCalendar(MyCal As PictureBox)
   Dim x As Integer
   Dim titlewidth As Long
   Dim RowLoop As Long
   Dim ColLoop As Long
   Dim dt As Date
   Dim CurrentDate As Date
   Dim FirstDate As Date
   Dim GoFrom As Date
   Dim DateLoop As Date
   Dim Row As Long
   Dim DatePos(1 To 6, 1 To 7) As Date
   Dim strDate As String
   Dim PosX As Long
   Dim PosY As Long
   Dim rotvalue As Long
   ct = 1
   For x = 1 To 6
      strDate = x & "/01/ " & cboYear.Text
      CurrentDate = CDate(strDate)
      
      MyCal.Visible = True
      MyCal.AutoRedraw = True
      
      ' Store dates
      FirstDate = DateSerial(Year(CurrentDate), Month(CurrentDate), 1)
      If chkLatin.Value = Unchecked Then
         GoFrom = IIf(Weekday(FirstDate) = 1, FirstDate, FirstDate - Weekday(FirstDate) + 1)
      Else
         GoFrom = IIf(Weekday(FirstDate) = 1, FirstDate - 7, FirstDate - Weekday(FirstDate) + 1)
      End If
      Row = 1
      For DateLoop = GoFrom To GoFrom + 41
         If chkLatin.Value = Unchecked Then
            DatePos(Row, Weekday(DateLoop)) = DateLoop
         Else
            DatePos(Row, Weekday(DateLoop)) = DateLoop + 1
         End If
         If Weekday(DateLoop) = 7 Then Row = Row + 1
      Next DateLoop
      
      With MyCal
         .Cls
         MyCal.Scale (0, 0)-(70, 87)
         
         ' Define font sizes/colors
         .FontSize = .FontSize / 1.1
         .FontBold = True
         ' Write headers
         .FontSize = 8
         titlewidth = .TextWidth(Format(CurrentDate, "mmmm yyyy"))
         .ForeColor = vbBlack
         .CurrentY = 0
         .CurrentX = ((50 - titlewidth) / 2) + 10
         MyCal.Print Format(CurrentDate, "mmmm yyyy")
         .FontSize = 8
         .ForeColor = vbBlack
         .ForeColor = vbBlack
         .CurrentY = 13
         If chkLatin.Value = Unchecked Then
            MyCal.Print "  Su  Mo  Tu  We Th   Fr   Sa"
         Else
            MyCal.Print "  Mo  Tu  We Th   Fr   Sa  Su"
         End If
         ' Loop through stored dates and write to screen
         For RowLoop = 1 To 6
            For ColLoop = 1 To 7
               dt = DatePos(RowLoop, ColLoop) ' store date for quick access
               .CurrentY = (RowLoop + 1) * 10 + 5
               .CurrentX = ((ColLoop - 1) * 10) + (10 - .TextWidth(Day(dt))) / 2 - 1
               If chkLatin.Value = Unchecked Then
                  If ColLoop = 1 Then
                     If chkSunday.Value = Checked Then
                       .ForeColor = vbRed
                     Else
                       .ForeColor = vbBlack
                     End If
                  Else
                     .ForeColor = vbBlack
                  End If
                  If ColLoop = 7 Then
                     If chkSaturday.Value = Checked Then
                      .ForeColor = vbRed
                     Else
                       .ForeColor = vbBlack
                     End If
                  End If
               Else
                  If ColLoop = 7 Then
                     If chkSunday.Value = Checked Then
                       .ForeColor = vbRed
                     Else
                       .ForeColor = vbBlack
                     End If
                  Else
                     .ForeColor = vbBlack
                  End If
                  If ColLoop = 6 Then
                     If chkSaturday.Value = Checked Then
                      .ForeColor = vbRed
                     Else
                       .ForeColor = vbBlack
                     End If
                  End If
               End If
                  
               If Format(dt, "mmmyy") <> Format(CurrentDate, "mmmyy") Then GoTo here
               MyCal.Print Day(dt) ' print the number to the screen
here:
            Next ColLoop
         Next RowLoop
      End With
      If x = 1 Then PosX = 265: PosY = 225: rotvalue = 108
      If x = 2 Then PosX = 430: PosY = 285: rotvalue = 72
      If x = 3 Then PosX = 365: PosY = 75: rotvalue = 144
      If x = 4 Then PosX = 140: PosY = 70: rotvalue = 215
      If x = 5 Then PosX = 80: PosY = 284: rotvalue = 288
      If x = 6 Then PosX = 255: PosY = 415: rotvalue = 0
      picRot.Picture = LoadPicture()
      DoEvents
      If chkLogo.Value = Checked Then PicDate.PaintPicture Picture1, 30, 77, 16, 16, 0, 0, 16, 16
      bmp_rotate2 PicDate, picRot, rotvalue * Trans
      picRot.Picture = picRot.Image
      picMain.PaintPicture picRot, PosX, PosY, picRot.ScaleWidth, picRot.ScaleHeight, 0, 0, picRot.ScaleWidth, picRot.ScaleHeight
      PicDate.Picture = LoadPicture()
      picMain.Picture = picMain.Image
      Image1.Picture = picMain.Picture
      
      PB1.Value = ct
      ct = ct + 1
      PB1.BarColor = &HC0C0&
   Next x
   
   For x = 12 To 7 Step -1   'need to print December before the other months
      strDate = x & "/01/ " & cboYear.Text
      CurrentDate = CDate(strDate)
      
      ' Store dates
      FirstDate = DateSerial(Year(CurrentDate), Month(CurrentDate), 1)
      If chkLatin.Value = Unchecked Then
         GoFrom = IIf(Weekday(FirstDate) = 1, FirstDate, FirstDate - Weekday(FirstDate) + 1)
      Else
         GoFrom = IIf(Weekday(FirstDate) = 1, FirstDate - 7, FirstDate - Weekday(FirstDate) + 1)
      End If
      Row = 1
      For DateLoop = GoFrom To GoFrom + 41
         If chkLatin.Value = Unchecked Then
            DatePos(Row, Weekday(DateLoop)) = DateLoop
         Else
            DatePos(Row, Weekday(DateLoop)) = DateLoop + 1
         End If
         If Weekday(DateLoop) = 7 Then Row = Row + 1
      Next DateLoop
      
      With MyCal
         .Cls
         MyCal.Scale (0, 0)-(70, 87)
         
         ' Define font sizes/colors
         .FontSize = .FontSize / 1.1
         .FontBold = True
         ' Write headers
         .FontSize = 8
         titlewidth = .TextWidth(Format(CurrentDate, "mmmm yyyy"))
         .ForeColor = vbBlack
         .CurrentY = 0
         .CurrentX = ((50 - titlewidth) / 2) + 10
         MyCal.Print Format(CurrentDate, "mmmm yyyy")
         .FontSize = 8
         .ForeColor = vbBlack
         .ForeColor = vbBlack
         .CurrentY = 13
         If chkLatin.Value = Unchecked Then
            MyCal.Print "  Su  Mo  Tu  We Th   Fr   Sa"
         Else
            MyCal.Print "  Mo  Tu  We Th   Fr   Sa  Su"
         End If
         
         ' Loop through stored dates and write to screen
         For RowLoop = 1 To 6
            For ColLoop = 1 To 7
               dt = DatePos(RowLoop, ColLoop) ' store date for quick access
               .CurrentY = (RowLoop + 1) * 10 + 5
               .CurrentX = ((ColLoop - 1) * 10) + (10 - .TextWidth(Day(dt))) / 2 - 1
                If chkLatin.Value = Unchecked Then
                  If ColLoop = 1 Then
                     If chkSunday.Value = Checked Then
                       .ForeColor = vbRed
                     Else
                       .ForeColor = vbBlack
                     End If
                  Else
                     .ForeColor = vbBlack
                  End If
                  If ColLoop = 7 Then
                     If chkSaturday.Value = Checked Then
                      .ForeColor = vbRed
                     Else
                       .ForeColor = vbBlack
                     End If
                  End If
               Else
                  If ColLoop = 7 Then
                     If chkSunday.Value = Checked Then
                       .ForeColor = vbRed
                     Else
                       .ForeColor = vbBlack
                     End If
                  Else
                     .ForeColor = vbBlack
                  End If
                  If ColLoop = 6 Then
                     If chkSaturday.Value = Checked Then
                      .ForeColor = vbRed
                     Else
                       .ForeColor = vbBlack
                     End If
                  End If
               End If
               
               If Format(dt, "mmmyy") <> Format(CurrentDate, "mmmyy") Then GoTo here1
               MyCal.Print Day(dt) ' print the number to the screen
here1:
            Next ColLoop
         Next RowLoop
      End With
      
      If x = 7 Then PosX = 360: PosY = 538: rotvalue = 179
      If x = 8 Then PosX = 182: PosY = 660: rotvalue = 249
      If x = 9 Then PosX = 245: PosY = 875: rotvalue = 323
      If x = 10 Then PosX = 470: PosY = 880: rotvalue = 36
      If x = 11 Then PosX = 535: PosY = 666: rotvalue = 108
      If x = 12 Then PosX = 370: PosY = 730: rotvalue = 71
      picRot.Picture = LoadPicture()
      DoEvents
      If chkLogo.Value = Checked Then PicDate.PaintPicture Picture1, 30, 77, 16, 16, 0, 0, 16, 16
      bmp_rotate2 PicDate, picRot, rotvalue * Trans
      picRot.Picture = picRot.Image
      picMain.PaintPicture picRot, PosX, PosY, picRot.ScaleWidth, picRot.ScaleHeight, 0, 0, picRot.ScaleWidth, picRot.ScaleHeight
      PicDate.Picture = LoadPicture()
      picMain.Picture = picMain.Image
      Image1.Picture = picMain.Picture

      If ct = 12 Then
         Command2.Enabled = True
         Command3.Enabled = True
         Command5.Enabled = True
         PB1.Value = 13
      Else
         PB1.Value = ct
         ct = ct + 1
      End If
   Next x
  
   DrawPoly
   ct = 0
   PB1.Value = ct
   PB1.BarColor = vbWhite
End Sub

Private Sub DrawPoly()

   picMain.Line (132, 86)-(102, 179), &HD1D1FB
   picMain.Line (181, 86)-(132, 86), &HD1D1FB
   picMain.Line (101, 179)-(142, 209), &HD1D1FB
   picMain.Line (197, 39)-(181, 87), &HD1D1FB
   picMain.Line (293, 40)-(196, 40), &HD1D1FB
   picMain.Line (349, 209)-(293, 39), &HD1D1FB
   picMain.Line (388, 86)-(308, 86), &HD1D1FB
   picMain.Line (141, 209)-(117, 283), &HD1D1FB
   picMain.Line (403, 40)-(387, 88), &HD1D1FB
   picMain.Line (501, 40)-(402, 40), &HD1D1FB
   picMain.Line (502, 41)-(516, 89), &HD1D1FB
   picMain.Line (545, 86)-(515, 86), &HD1D1FB
   picMain.Line (546, 87)-(580, 192), &HD1D1FB
   picMain.Line (450, 284)-(581, 189), &HD1D1FB
   picMain.Line (580, 284)-(556, 209), &HD1D1FB
   picMain.Line (245, 283)-(69, 283), &HD1D1FB
   picMain.Line (38, 376)-(70, 283), &HD1D1FB
   picMain.Line (38, 377)-(79, 406), &HD1D1FB
   picMain.Line (62, 452)-(77, 405), &HD1D1FB
   picMain.Line (142, 510)-(62, 451), &HD1D1FB
   picMain.Line (141, 508)-(181, 479), &HD1D1FB
   picMain.Line (181, 480)-(245, 527), &HD1D1FB
   picMain.Line (284, 406)-(229, 573), &HD1D1FB
   picMain.Line (580, 284)-(610, 284), &HD1D1FB
   picMain.Line (643, 388)-(610, 285), &HD1D1FB
   picMain.Line (618, 407)-(643, 388), &HD1D1FB
   picMain.Line (413, 405)-(452, 530), &HD1D1FB
   picMain.Line (229, 573)-(310, 631), &HD1D1FB
   picMain.Line (619, 406)-(634, 454), &HD1D1FB
   picMain.Line (556, 508)-(633, 452), &HD1D1FB
   picMain.Line (516, 481)-(555, 509), &HD1D1FB
   picMain.Line (451, 527)-(516, 482), &HD1D1FB
   picMain.Line (180, 725)-(349, 602), &HD1D1FB
   picMain.Line (453, 528)-(724, 724), &HD1D1FB
   picMain.Line (619, 648)-(516, 725), &HD1D1FB
   picMain.Line (723, 725)-(619, 1044), &HD1D1FB
   picMain.Line (620, 1043)-(283, 1043), &HD1D1FB
   picMain.Line (181, 725)-(284, 1044), &HD1D1FB
   picMain.Line (285, 649)-(387, 724), &HD1D1FB
   picMain.Line (684, 849)-(555, 849), &HD1D1FB
   picMain.Line (491, 1043)-(452, 920), &HD1D1FB
   picMain.Line (244, 922)-(348, 846), &HD1D1FB
   
   picMain.DrawStyle = 2
   
   picMain.Line (181, 87)-(141, 211), &HFCD3C8
   picMain.Line (182, 86)-(309, 86), &HFCD3C8
   picMain.Line (246, 284)-(142, 210), &HFCD3C8
   picMain.Line (348, 208)-(244, 283), &HFCD3C8
   picMain.Line (387, 87)-(349, 208), &HFCD3C8
   picMain.Line (386, 86)-(515, 86), &HFCD3C8
   picMain.Line (515, 87)-(557, 212), &HFCD3C8
   picMain.Line (348, 208)-(452, 283), &HFCD3C8
   picMain.Line (77, 406)-(117, 284), &HFCD3C8
   picMain.Line (245, 284)-(284, 406), &HFCD3C8
   picMain.Line (283, 406)-(180, 480), &HFCD3C8
   picMain.Line (182, 481)-(76, 405), &HFCD3C8
   picMain.Line (452, 283)-(581, 283), &HFCD3C8
   picMain.Line (451, 284)-(412, 407), &HFCD3C8
   picMain.Line (283, 405)-(412, 405), &HFCD3C8
   picMain.Line (580, 284)-(620, 410), &HFCD3C8
   picMain.Line (516, 481)-(620, 405), &HFCD3C8
   picMain.Line (412, 405)-(520, 484), &HFCD3C8
   picMain.Line (244, 527)-(350, 603), &HFCD3C8
   picMain.Line (452, 527)-(347, 603), &HFCD3C8
   picMain.Line (388, 726)-(349, 603), &HFCD3C8
   picMain.Line (556, 602)-(515, 726), &HFCD3C8
   picMain.Line (553, 847)-(660, 922), &HFCD3C8
   picMain.Line (451, 924)-(413, 1044), &HFCD3C8
   picMain.Line (220, 846)-(349, 846), &HFCD3C8
   picMain.Line (388, 725)-(516, 725), &HFCD3C8
   picMain.Line (516, 725)-(555, 848), &HFCD3C8
   picMain.Line (555, 847)-(451, 923), &HFCD3C8
   picMain.Line (452, 923)-(349, 846), &HFCD3C8
   picMain.Line (388, 725)-(348, 848), &HFCD3C8
   
   picMain.DrawStyle = 0
   
   picMain.CurrentX = 20
   picMain.CurrentY = 900
   picMain.ForeColor = vbRed
   picMain.Print "Red Lines--Cut"
   picMain.Print
   picMain.CurrentX = 20
   picMain.ForeColor = vbBlue
   picMain.Print "Blue Lines--Score and Fold"
   picMain.Print
   picMain.CurrentX = 20
   picMain.ForeColor = vbBlack
   picMain.Print "Print on Cardstock"
   picMain.CurrentX = 640
   picMain.CurrentY = 640
   picMain.FontSize = 12
   picMain.FontBold = True
   picMain.Print "Glue Last"
   picMain.FontBold = False
   picMain.FontSize = 8
   picMain.Picture = picMain.Image
   Image1.Picture = picMain.Picture
   
End Sub

Private Sub Command1_Click()   'make calendar
   picMain.Picture = LoadPicture()
   DrawCalendar PicDate
End Sub

Private Sub Command2_Click()  'Save as bitmap
   BitmapSave picMain
End Sub

Private Sub Command3_Click()  'Print
   picMain.ScaleMode = vbCentimeters
   Printer.PaintPicture picMain.Picture, 0, -800
   Printer.EndDoc
End Sub

Private Sub Command4_Click()  'clear
   picMain.Picture = LoadPicture()
   Image1.Picture = LoadPicture()
   Command2.Enabled = False
   Command3.Enabled = False
   Command5.Enabled = False
   'Label4.Caption = "0"
End Sub

Private Sub Command5_Click() ' Save Jpeg
   JpegSave picMain
End Sub

Private Sub Picture1_Click()
    On Error Resume Next
    ShowOpen
    Picture1.Picture = LoadPicture(cmndlg.Filename)
End Sub

