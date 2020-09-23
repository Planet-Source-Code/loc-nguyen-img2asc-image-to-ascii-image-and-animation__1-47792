VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmASC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image to ASCII - Loc2K"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmASC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   653
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   677
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3570
      ItemData        =   "FrmASC.frx":0CCA
      Left            =   2040
      List            =   "FrmASC.frx":0D85
      TabIndex        =   26
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   3600
      Left            =   0
      Pattern         =   "*.bmp;*.gif;*.jpg;*.jpeg"
      TabIndex        =   25
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Real-&Time Processing"
      Height          =   1995
      Left            =   7260
      TabIndex        =   0
      Top             =   7680
      Width           =   2775
      Begin VB.CommandButton CmdLoad 
         Caption         =   "(1) &Load Image"
         Height          =   435
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   2535
      End
      Begin VB.CommandButton CmdMosaic 
         Caption         =   "(2) &Write Matrix"
         Height          =   435
         Left            =   120
         TabIndex        =   2
         Top             =   870
         Width           =   2535
      End
      Begin VB.CommandButton CmdASC 
         Caption         =   "(3) &Export ASCII"
         Height          =   435
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proc&ess Options"
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   7680
      Width           =   7035
      Begin VB.CommandButton CmdBrowse 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   6480
         TabIndex        =   21
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   22
         Text            =   "C:\"
         Top             =   720
         Width           =   1920
      End
      Begin VB.CheckBox Check4 
         Caption         =   "E&xport Animation (Alphabetically) HTML"
         Height          =   255
         Left            =   3720
         TabIndex        =   20
         Top             =   480
         Width           =   3135
      End
      Begin VB.OptionButton Option5 
         Caption         =   "1px:1&ch Output (426x426 Max) (Faster)"
         Height          =   255
         Left            =   3480
         TabIndex        =   19
         ToolTipText     =   "Ouput text will be bigger than input image by a factor of 6 to 7 (depending on font size)"
         Top             =   240
         Value           =   -1  'True
         Width           =   3255
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Actual Size &Output (2560x2560 Max)"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Output text will be roughly the same size as input image"
         Top             =   240
         Width           =   3255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Export As &HTML"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Generate Mosaic &Preview (Slower)"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Director&y:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Palette Optio&ns"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   8820
      Width           =   7035
      Begin VB.VScrollBar VScroll2 
         Height          =   255
         Left            =   6090
         Max             =   24
         Min             =   2
         TabIndex        =   27
         Top             =   450
         Value           =   2
         Width           =   255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmASC.frx":14E5
         Left            =   5760
         List            =   "FrmASC.frx":152E
         Style           =   1  'Simple Combo
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   420
         Width           =   615
      End
      Begin VB.CheckBox Check3 
         Caption         =   "&Aliased 8pt Optimized Palette"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "(Otherwise optimized for antialiased [XP ClearType] size 1 web font under Internet Explorer ""Medium"" text size)"
         Top             =   480
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Modular Palette &Reduction:"
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         ToolTipText     =   "Use this option to reduce the modular palette (for detail-intensive images)"
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ASCII ""&Dithering"" (Sub-Palette)"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Use this option for images with subtle color transitions."
         Top             =   240
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Modular Palette (""~4.643... bit"")"
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         ToolTipText     =   "Use this option for images with obvious color transitions."
         Top             =   240
         Value           =   -1  'True
         Width           =   3255
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Common Image Files (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg;*.jpeg"
   End
   Begin VB.PictureBox PicBG 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   7515
      Left            =   120
      ScaleHeight     =   497
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   657
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   9915
      Begin VB.PictureBox PicASC 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3360
         Index           =   1
         Left            =   0
         ScaleHeight     =   224
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   232
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   255
         LargeChange     =   4096
         Left            =   0
         SmallChange     =   1024
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   7200
         Width           =   9615
      End
      Begin VB.VScrollBar VScroll1 
         Enabled         =   0   'False
         Height          =   7215
         LargeChange     =   4096
         Left            =   9600
         SmallChange     =   1024
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   7200
         Width           =   255
      End
      Begin VB.PictureBox PicASC 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   6390
         Index           =   0
         Left            =   0
         ScaleHeight     =   426
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   426
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   6390
      End
   End
End
Attribute VB_Name = "FrmASC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'aMatrix(X - 1, b)  aMatrix(a, X - 1)
'2560 \ 6 = 426     2560 \ 12 = 213
Private aMatrix(0 To 425, 0 To 212) As Double   'Variable equivalent of the mosaic
Private wUB As Long, hUB As Long    'Mosaic-al dimensions
'Palette array, global document title
Private Atrix(0 To 24) As String, fTitle As String
Private EscAni As Boolean

Private Sub Check3_Click()
    DoPalette
End Sub

Private Sub Check4_Click()
    If Check4.Value = 1 Then
        Label1.Enabled = True
        Text1.Enabled = True
        CmdBrowse.Enabled = True
        Check2.Enabled = False
        Check2.Value = 1
        CmdMosaic.Enabled = False
    Else
        Label1.Enabled = False
        Text1.Enabled = False
        CmdBrowse.Enabled = False
        Check2.Enabled = True
        CmdMosaic.Enabled = True
    End If
End Sub

Private Sub CmdASC_Click()
    On Error GoTo ErrorHandler
    'Number of elements in aRow()'s first dimension is the same as aMatrix()'s second
    Dim aRow(0 To 212) As String, txtOut As String, prgOut As String
    Dim fI As Double, eI As Double, aI As Double
    Dim sI As Long, i As Long, j As Long, i2 As Long, h As Long, w As Long
    Dim a
    'The animation loop
    If Check4.Value = 1 Then
        a = MsgBox("The animation job may take a while (depending on the amount of files and size of each file)." & vbCrLf & "All frames must be 426x426 or smaller.  (Same-sized frames are optimal.)" & vbCrLf & vbCrLf & "Do you wish to continue?" & vbCrLf & vbCrLf & "Remember that you can cancel anytime during the operation by pressing the ""Esc"" key.", vbYesNo, "Job Confirmation")
        If a = vbNo Then Exit Sub
        Check4.Value = 0
        'If Dir(Text1.Text) <> "" Then File1.Path = Text1.Text Else GoTo ErrorHandler
        File1.Path = Text1.Text
        If File1.ListCount = 0 Then GoTo ErrorHandler
        txtOut = App.Path & "\" & "Output.htm"
        If Dir(txtOut) <> "" Then Kill txtOut   'Check the existence of old output file
        'Write output file
        Open txtOut For Append As #1
        'HTML header
        Print #1, "<html>"
        Print #1, "<head>"
        Print #1, " <title>" & Left(File1.List(0), InStr(1, File1.List(0), ".") - 1) & " et al Animation</title>"
        Print #1, " <script>"
        Print #1, "  <!-- DHTML based on code from cjr, miK, and Meph (used without permission) -->"
        Print #1, "  var max_pics=" & CStr(File1.ListCount) & ";"
        For i = 0 To List1.ListCount - 1
            Print #1, List1.List(i)
        Next i
        For i = 0 To File1.ListCount - 1
            'Insert Matrix code
            PicASC(0).Picture = LoadPicture(File1.Path & "\" & File1.List(i))
            PicASC(0).Refresh
            If PicASC(0).Width > 426 Or PicASC(0).Height > 426 Then
                MsgBox "One or more of the frames in the selected directory are too large.  The operation will now terminate.", vbCritical, "Error"
                Exit Sub
            End If
            CmdMosaic_Click
            'Insert Export code
            CmdASC.Caption = "(3) &Export ASCII (" & Format(100 * i / (File1.ListCount - 1), "0.0") & "%)"
            DoEvents
            If EscAni = True Then
                EscAni = False
                GoTo Terminate
            End If
            For h = 0 To hUB
                aRow(h) = ""
                For w = 0 To wUB
                    fI = aMatrix(w, h)
                    eI = fI - Int(fI)
                    sI = Len(Atrix(Int(fI))) - 1
                    aI = eI * sI + 1    'Considers eI (fractional value) to simulate dithering
                    aRow(h) = aRow(h) & Mid(Atrix(Int(fI)), aI, 1)
                Next w
                'Replace potentially browser-decoded characters
                aRow(h) = Replace(aRow(h), "<", "&lt;")
                aRow(h) = Replace(aRow(h), ">", "&gt;")
            Next h
            Print #1, " <div id=""text" & CStr(i + 1) & """>"
            Print #1, "  <pre><font face=""Courier New"" size=1>"
            For j = 0 To hUB
                Print #1, aRow(j)
            Next j
            Print #1, "  </font></pre>"
            Print #1, " </div>"
        Next i
        Print #1, "</body>"
        Print #1, "</html>"
Terminate:
        Close #1
        Shell "explorer.exe " & txtOut, vbNormalFocus
        CmdASC.Caption = "(3) &Export ASCII"
        Exit Sub
    End If
    For h = 0 To hUB
        For w = 0 To wUB
            fI = aMatrix(w, h)
            eI = fI - Int(fI)
            sI = Len(Atrix(Int(fI))) - 1
            aI = eI * sI + 1    'Considers eI (fractional value) to simulate dithering
            aRow(h) = aRow(h) & Mid(Atrix(Int(fI)), aI, 1)
        Next w
        If Check2.Value = 1 Then
            'Replace potentially browser-decoded characters
            aRow(h) = Replace(aRow(h), "<", "&lt;")
            aRow(h) = Replace(aRow(h), ">", "&gt;")
            CmdASC.Caption = "(3) &Export ASCII (" & Format(100 * h / hUB, "0.0") & "%)"
        End If
    Next h
    CmdASC.Caption = "(3) &Export ASCII"
    If Check2.Value = 1 Then txtOut = "Output.htm" Else txtOut = "Output.txt"
    txtOut = App.Path & "\" & txtOut
    If Dir(txtOut) <> "" Then Kill txtOut   'Check the existence of old output file
    'Write output file ("Output.htm" or "Output.txt" in application directory)
    Open txtOut For Append As #1
        If Check2.Value = 1 Then Print #1, " <pre><font face=""Courier New"" size=1>"
        For i = 0 To hUB
            Print #1, aRow(i)
        Next i
        If Check2.Value = 1 Then Print #1, " </font></pre>"
    Close #1
    'Open proper text viewer or Internet browser
    If Check2.Value = 1 Then Shell "explorer.exe " & txtOut, vbNormalFocus Else Shell "notepad.exe " & txtOut, vbNormalFocus
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred.  Please make sure you have written a matrix before exporting." & vbCrLf & vbCrLf & "If using Animation mode, please make sure the path provided is correct and that the directory contains image files.", vbCritical, "Error"
End Sub

Private Sub CmdASC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then EscAni = True
End Sub

Private Sub CmdBrowse_Click()
    FrmBrowse.Show 1
End Sub

Private Sub CmdLoad_Click()
    Dim tTitle As String
    On Error GoTo ErrorHandler
    tTitle = fTitle
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = 4
    CommonDialog1.ShowOpen
    fTitle = CommonDialog1.FileTitle
    PicASC(1).Picture = PicASC(0).Image
    PicASC(0).Picture = LoadPicture(CommonDialog1.FileName)
    PicASC(0).Left = 0
    PicASC(0).Top = 0
    'Disallow high dimensions to provide practical processing time
    If PicASC(0).Width > 2560 Or PicASC(0).Height > 2048 Then
        MsgBox "The image is too large.  Please resize it to 2560 by 2560 or smaller.", vbCritical, "Error"
        PicASC(0).Picture = PicASC(1).Image
        fTitle = tTitle
    Else
        tTitle = fTitle
    End If
    If PicASC(0).Width > 426 Or PicASC(0).Height > 426 Then Option5.Enabled = False Else Option5.Enabled = True
    If PicASC(0).Width > 640 Then HScroll1.Enabled = True Else HScroll1.Enabled = False
    If PicASC(0).Height > 480 Then VScroll1.Enabled = True Else VScroll1.Enabled = False
    'Display filename and dimensions in titlebar
    Me.Caption = App.Title & " - """ & tTitle & """ [" & PicASC(0).Width & "x" & PicASC(0).Height & "]"
    If PicASC(0).Width > 426 Or PicASC(0).Height > 426 Then Option4.Value = True
    Exit Sub
ErrorHandler:
    If Err.Number <> 32755 Then MsgBox "Invalid filename.", vbCritical, "Error"
End Sub

Private Sub CmdMosaic_Click()
    Dim w As Long, h As Long, sw As Long, sh As Long, cR As Long, cG As Long, cB As Long, nRGB As Long
    Dim rTotal As Double
    Dim cRGB As String
    Dim xFac As Integer, yFac As Integer
    'Accord to mosaic block size (depending on output mode)
    If Option4.Value = True Then
        xFac = 6
        yFac = 12
    Else
        xFac = 1
        yFac = 2
    End If
    'Resize image to usable pixels
    PicASC(0).Width = PicASC(0).Width - (PicASC(0).Width Mod xFac)
    PicASC(0).Height = PicASC(0).Height - (PicASC(0).Height Mod yFac)
    'Find the mosaic (modular) dimensions
    wUB = PicASC(0).Width / xFac - 1
    hUB = PicASC(0).Height / yFac - 1
    For h = 0 To hUB
        For w = 0 To wUB
            rTotal = 0
            'Create virtual mosaic
            For sh = 0 To yFac - 1
                For sw = 0 To xFac - 1
                    'Grayscale conversion
                    cRGB = Hex(PicASC(0).Point(w * xFac + sw, h * yFac + sh))
                    cRGB = String(6 - Len(cRGB), "0") & cRGB
                    cR = Val("&H" & Right(cRGB, 2))
                    cG = Val("&H" & Mid(cRGB, 3, 2))
                    cB = Val("&H" & Left(cRGB, 2))
                    rTotal = rTotal + (cR + cG + cB) / 3
                Next sw
            Next sh
            nRGB = rTotal / (xFac * yFac)   'Calculate mosaic element
            aMatrix(w, h) = ((nRGB + 1) / 256) * 24 'Find corresponding Atrix() value
            If Check1.Value = 1 Then
                'Draw the mosaic
                For sh = 0 To yFac - 1
                    For sw = 0 To xFac - 1
                        PicASC(0).PSet (w * xFac + sw, h * yFac + sh), RGB(nRGB, nRGB, nRGB)
                    Next sw
                Next sh
            End If
        Next w
        'Update displayed information
        If Check1.Value = 1 Then PicASC(0).Refresh
        CmdMosaic.Caption = "(2) &Write Matrix (" & Format(100 * h / hUB, "0.0") & "%)"
    Next h
    CmdMosaic.Caption = "(2) &Write Matrix"
End Sub

Private Sub Combo1_Click()
    Option3.Value = True
    DoPalette
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    VScroll2.Value = Combo1.ListIndex + 2
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Load()
    'Assign default values
    Text1.Text = Left(App.Path, 3)
    Combo1.ListIndex = 3
    VScroll2.Value = Combo1.ListIndex + 2
    fTitle = "Untitled"
    Option2.Value = True
    DoPalette
End Sub

Private Sub HScroll1_Change()
    HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
    Dim ScrVal As Long
    ScrVal = (PicASC(0).Width - 640) * HScroll1.Value / HScroll1.Max
    If Check1.Value = 0 Then
        PicASC(0).Left = -ScrVal
    Else
        PicASC(0).Left = -(6 * (ScrVal \ 6))
    End If
End Sub

Private Sub Option1_Click()
    DoPalette
End Sub

Private Sub Option2_Click()
    DoPalette
End Sub

Private Sub Option3_Click()
    DoPalette
End Sub

'Readjusts the palette according to user-specified options
Private Sub DoPalette()
    Dim i As Long, j As Long, ni As Long, a As Long, b As Long, c As Long, d As Long
    Dim PalStr As String, nPalStr As String
    If Check3.Value = 1 Then
        'Assign aliased palette ("dithered")
        PalStr = "MNH#EmRFX953$%+}?>;""^:'` "
        'Based on character density of 0 to 27 pixels omitting 1 and 24 and 25
        Atrix(0) = "M":         Atrix(1) = "N"
        Atrix(2) = "H":         Atrix(3) = "#@WKB"
        Atrix(4) = "E":         Atrix(5) = "m"
        Atrix(6) = "RQUkgqph":  Atrix(7) = "FADVdb"
        Atrix(8) = "XSZ8GPyan": Atrix(9) = "96YO04w"
        Atrix(10) = "5ITLzue":  Atrix(11) = "32&C][fv"
        Atrix(12) = "$Jjxstr":  Atrix(13) = "%7l1ioc"
        Atrix(14) = "+":        Atrix(15) = "}{="
        Atrix(16) = "?*":       Atrix(17) = "><\/)(|"
        Atrix(18) = ";_":       Atrix(19) = """!~"
        Atrix(20) = "^-,":      Atrix(21) = ":"
        Atrix(22) = "'":        Atrix(23) = "`."
        Atrix(24) = " "
    Else
        'Assign antialiased palette ("dithered")
        PalStr = "MN@#SUZT$83Jj1v+[}~;^:.` "
        'Based on graphical density determined by driver program
        'Important to note the omition of characters as "Q" and certain lowercase characters
        '   "Q" is "darker" than "M" but is far less homogenous
        '   Certain (most) lowercase letters lack homogeneity
        'Also important is this is graphical density--only accurate with antialiased (Windows XP ClearType) font
        Atrix(0) = "M":         Atrix(1) = "NHWEB"
        Atrix(2) = "@XRK":      Atrix(3) = "#%"
        Atrix(4) = "SGDP":      Atrix(5) = "UA"
        Atrix(6) = "ZF96":      Atrix(7) = "T"
        Atrix(8) = "$52OC":     Atrix(9) = "8Y4"
        Atrix(10) = "3LV":      Atrix(11) = "JI0"
        Atrix(12) = "j&uroc":   Atrix(13) = "1li"
        Atrix(14) = "v><":      Atrix(15) = "+?"
        Atrix(16) = "[]""":     Atrix(17) = "}{"
        Atrix(18) = "~":        Atrix(19) = ";)(\/|!"
        Atrix(20) = "^-":       Atrix(21) = ":,'"
        Atrix(22) = ".":        Atrix(23) = "`"
        Atrix(24) = " "
    End If
    'Assign modular palette (from "dithered" set)
    If Option1.Value = False Then
        For i = 0 To 24
            Atrix(i) = Left(Atrix(i), 1)
        Next i
    End If
    'Parse the palette into a values (modular reduction algorithm)
    'This is a mathematical distribution; it is equally effective (and perhaps easier to write) to predefine this palette
    '   Upon palette extension (unlikely, but extended ASCII may be considered), this algorithm will provide flexibility (no need for further predefinitions)
    If Option3.Value = True Then
        c = CInt(Combo1.Text)   'Palette size
        a = 25 \ c  'Palette element basic size
        b = 25 Mod c    'Remainder size (left to distribute)
        'Create the basic reduced palette
        nPalStr = String(a, Left(PalStr, 1))
        For i = 1 To c - 2
            'Append equal-interval characters
            nPalStr = nPalStr & String(a, Mid(PalStr, i * a, 1))
        Next i
        nPalStr = nPalStr & String(a, Right(PalStr, 1))
        'Homogenize the reduced palette
        If b > 1 Then
            'Use the remainder of the partitions in a (b) to determine d partitions
            d = Len(nPalStr) \ b
            For i = 0 To b - 1
                'Insert Mid(nPalStr, (i + 1) * d [the current partition] + i [the partition(s) offset]) into nPalStr at d
                nPalStr = Left(nPalStr, (i + 1) * d + i) & Mid(nPalStr, (i + 1) * d + i, 1) & Right(nPalStr, Len(nPalStr) - ((i + 1) * d + i))
            Next i
        End If
        'Append remainder Right(PalStr, 1) to offset Left(PalStr, 1) bias
        nPalStr = nPalStr & String(25 - Len(nPalStr), Right(PalStr, 1))
        'Apply the new palette
        For i = 0 To 24
            Atrix(i) = Mid(nPalStr, i + 1, 1)
        Next i
    End If
End Sub

Private Sub Option4_Click()
    Option5_Click
    Label1.Enabled = False
    Text1.Enabled = False
    CmdBrowse.Enabled = False
    Check4.Value = 0
End Sub

Private Sub Option5_Click()
    If Option5.Value = True Then Check4.Enabled = True Else Check4.Enabled = False
    Check4_Click
End Sub

Private Sub VScroll1_Change()
    VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()
    Dim ScrVal As Long
    ScrVal = (PicASC(0).Height - 480) * VScroll1.Value / VScroll1.Max
    If Check1.Value = 0 Then
        PicASC(0).Top = -ScrVal
    Else
        PicASC(0).Top = -(12 * (ScrVal \ 12))
    End If
End Sub

Private Sub VScroll2_Change()
    Combo1.ListIndex = VScroll2.Value - 2
End Sub
