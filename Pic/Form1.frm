VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   335
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   698
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   360
      Top             =   4920
   End
   Begin VB.CommandButton cmdLoadText 
      Caption         =   "Load Text"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdSaveText 
      Caption         =   "Save Text"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   3840
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox txtMain 
      Height          =   4095
      Left            =   6120
      TabIndex        =   9
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7223
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0BB0
   End
   Begin MSComDlg.CommonDialog CDText 
      Left            =   2520
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   360
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   265
      TabIndex        =   10
      Top             =   600
      Width           =   3975
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4560
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdSavePic 
      Caption         =   "Save Pic"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoadPic 
      Caption         =   "Load Pic"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Text-Box"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Write Text"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read Text"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog Dialog2 
      Left            =   1920
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Pictures (*.bmp)|*.bmp"
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   1200
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Pictures (*.bmp;*.jpg;*.jpeg)|*.bmp;*.jpg;*.jpeg"
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   10050
      Top             =   240
      Width           =   45
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   360
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "!!!Text In Pictures!!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4440
      TabIndex        =   11
      Top             =   240
      Width           =   1635
   End
   Begin VB.Label lblPassword 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'____________________________________________________________________________
'First Start the program and then type the password
'The password is "bluesoft" without the quotes.
'then click on 'Read Text' button
'You will have an idea that how much text we can
'write into a picture of this size and what effect does it have
'On the original picture.
'____________________________________________________________________________
'Note:- 32Bit Display is Must, to run this Program
'Best View results are on 1024 x 768 Pixel Display
'Because I made this on that Display settings.
'This is for sending messages to someone when you want top-secrecy
'I have made it as Optimized as I can. If you have anny more Optimization tips
'Then email me at keral82@keral.com
'©BlueSoft 2003
'Made by Keral.C.Patel.
'____________________________________________________________________________
'One thing more:- If you like the effects and want some more effects then just let me know.
'The following API's are Used for Flat buttons and Moveable Form
'____________________________________________________________________________
'Now the Logic behind this technique
'I have read about this on a Website.(Security Site)
'Here three Pixels are taken and then their RGB value is derived.
'Their RGB value is then converted into Even Numbers.
'Then Their are checked and then the binary value of the character is
'added to it.
'I think other things are self explanatory
'If you have questions then email me
'I will try my best to answer the questions regarding this
'____________________________________________________________________________

Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Const BS_FLAT = &H8000&

Private mHover As Boolean
Dim InitBTStyle As Long
Private Const GWL_STYLE = (-16)

Dim i As Integer, intcntr As Integer
Dim arrayA() As Integer, arrayB() As Integer, ln As Integer

Public Sub SetInitialBTStyle(BT As CommandButton)
    
    If GetWindowLong&(BT.hwnd, GWL_STYLE) = InitBTStyle Then Exit Sub
    
    SetWindowLong& BT.hwnd, GWL_STYLE, InitBTStyle
    BT.Refresh

End Sub

Public Sub GetInitialBTStyle(BT As CommandButton)
    
    InitBTStyle = GetWindowLong&(BT.hwnd, GWL_STYLE)

End Sub

Public Sub BTFlat(BT As CommandButton)
    
    If GetWindowLong&(BT.hwnd, GWL_STYLE) And BS_FLAT Then Exit Sub
    
    SetWindowLong BT.hwnd, GWL_STYLE, InitBTStyle Or BS_FLAT
    BT.Refresh

End Sub

Private Sub cmdClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If mHover Then SetInitialBTStyle cmdClear
Timer1.Enabled = False
End Sub

Private Sub cmdLoadPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If mHover Then SetInitialBTStyle cmdLoadPic
Timer1.Enabled = False
End Sub

Private Sub cmdLoadText_Click()

    CDText.ShowOpen

    If Trim(CDText.FileName) <> "" Then

        txtMain.LoadFile CDText.FileName

    End If

End Sub

Private Sub cmdLoadText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If mHover Then SetInitialBTStyle cmdLoadText
    Timer1.Enabled = False
End Sub

Private Sub cmdRead_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If mHover Then SetInitialBTStyle cmdRead
Timer1.Enabled = False
End Sub

Private Sub cmdSavePic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If mHover Then SetInitialBTStyle cmdSavePic
Timer1.Enabled = False
End Sub

Private Sub cmdSaveText_Click()

    CDText.ShowSave

    If Trim(CDText.FileName) <> "" Then

        txtMain.SaveFile CDText.FileName

    End If

End Sub

Private Sub cmdSaveText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If mHover Then SetInitialBTStyle cmdSaveText
Timer1.Enabled = False
End Sub

Private Sub cmdWrite_Click()
    Dim i As Long, j As Long, tx As String, ch As String, NrPix As Long
    Dim pix(0 To 2) As Long, wid As Long, hig As Long
    Dim r As Long, g As Long, b As Long, comp(1 To 8) As Long
    Dim aa(0 To 2) As Long, bb(0 To 2) As Long
    tx = "start message" & txtMain.Text & "end message"
    wid = picMain.ScaleWidth
    hig = picMain.ScaleHeight

    If Len(tx) * 3 > hig * wid Then

        tx = MsgBox("Text is " & Len(tx) * 3 - wid * hig & " Characters Longer than this Picture's capacity", vbCritical)
        Exit Sub

    End If

    txtPassword.Text = Trim$(txtPassword.Text)
    Shuffle (txtPassword.Text)

    For i = 1 To Len(tx)

        ch = ByteToBin(Asc(putc(Mid$(tx, i, 1))))
        NrPix = (CLng(i) - 1) * 3
        aa(0) = (NrPix Mod hig)
        bb(0) = (NrPix \ hig)
        pix(0) = picMain.Point(bb(0), aa(0)) 'the first pixel in the group of three.
        r = (pix(0) And RGB(255, 0, 0)) - (pix(0) And RGB(255, 0, 0)) Mod 2: comp(1) = r
        g = ((pix(0) And RGB(0, 255, 0)) \ 256) - ((pix(0) And RGB(0, 255, 0)) \ 256) Mod 2: comp(2) = g
        b = ((pix(0) And RGB(0, 0, 255)) \ 65536) - ((pix(0) And RGB(0, 0, 255)) \ 65536) Mod 2: comp(3) = b
    
        NrPix = NrPix + 1
        aa(1) = (NrPix Mod hig)
        bb(1) = (NrPix \ hig)
        pix(1) = picMain.Point(bb(1), aa(1)) 'the second pixel in the group of three.
        r = (pix(1) And RGB(255, 0, 0)) - (pix(1) And RGB(255, 0, 0)) Mod 2: comp(4) = r
        g = ((pix(1) And RGB(0, 255, 0)) \ 256) - ((pix(1) And RGB(0, 255, 0)) \ 256) Mod 2: comp(5) = g
        b = ((pix(1) And RGB(0, 0, 255)) \ 65536) - ((pix(1) And RGB(0, 0, 255)) \ 65536) Mod 2: comp(6) = b
    
        NrPix = NrPix + 1
        aa(2) = (NrPix Mod hig)
        bb(2) = (NrPix \ hig)
        pix(2) = picMain.Point(bb(2), aa(2)) 'the third pixel in the group of three.
        r = (pix(2) And RGB(255, 0, 0)) - (pix(2) And RGB(255, 0, 0)) Mod 2: comp(7) = r
        g = ((pix(2) And RGB(0, 255, 0)) \ 256) - ((pix(2) And RGB(0, 255, 0)) \ 256) Mod 2: comp(8) = g
        b = ((pix(2) And RGB(0, 0, 255)) \ 65536) 'last component remains unchanged
    
        For j = 1 To 8

            comp(j) = comp(j) + CInt(Mid$(ch, j, 1)) * 1

        Next

        picMain.PSet (bb(0), aa(0)), RGB(comp(1), comp(2), comp(3))
        picMain.PSet (bb(1), aa(1)), RGB(comp(4), comp(5), comp(6))
        picMain.PSet (bb(2), aa(2)), RGB(comp(7), comp(8), b)
    
    Next
End Sub

Private Sub cmdRead_Click()
    Dim i As Long, j As Long, k As Long, n As Long, pix(0 To 2) As Long
    Dim tx As String, nmd As Long, start As Integer
    Dim endmess As String, comp(1 To 8) As Long, ch As Long

    txtPassword.Text = Trim$(txtPassword.Text)
    Shuffle (txtPassword.Text)

    For i = 0 To picMain.ScaleWidth - 1

        For j = 0 To picMain.ScaleHeight - 1

            nmd = n Mod 3

            If nmd = 0 Then

                If start < 14 Then

                    start = start + 1

                    If start = 14 And tx <> "start message" Then

                        txtMain.Text = "There is no Text in this Picture or your Password is Wrong!!"
                        
                        Exit Sub

                    ElseIf start = 14 Then

                        tx = ""

                    End If

                End If

                ch = 0
                pix(nmd) = picMain.Point(i, j)
                comp(8) = ((pix(nmd) And RGB(255, 0, 0)) Mod 2)
                comp(7) = (((pix(nmd) And RGB(0, 255, 0)) \ 256) Mod 2)
                comp(6) = (((pix(nmd) And RGB(0, 0, 255)) \ 65536) Mod 2)

                For k = 8 To 6 Step -1

                    ch = ch + (2 ^ (k - 1)) * comp(k)

                Next

            End If

            If nmd = 1 Then

                pix(nmd) = picMain.Point(i, j)
                comp(5) = ((pix(nmd) And RGB(255, 0, 0)) Mod 2)
                comp(4) = (((pix(nmd) And RGB(0, 255, 0)) \ 256) Mod 2)
                comp(3) = (((pix(nmd) And RGB(0, 0, 255)) \ 65536) Mod 2)

                For k = 5 To 3 Step -1

                    ch = ch + (2 ^ (k - 1)) * comp(k)

                Next

            End If

            If nmd = 2 Then

                pix(nmd) = picMain.Point(i, j)
                comp(2) = ((pix(nmd) And RGB(255, 0, 0)) Mod 2)
                comp(1) = (((pix(nmd) And RGB(0, 255, 0)) \ 256) Mod 2)

                For k = 2 To 1 Step -1

                    ch = ch + (2 ^ (k - 1)) * comp(k)

                Next

            End If
        
            n = n + 1

            If n = 3 Then

                n = 0
                tx = tx & getc(Chr$(ch))

            End If

            endmess = Right$(tx, 11)

            If endmess = "end message" Then

                txtMain.Text = Left$(tx, Len(tx) - 11)
                
                Exit Sub

            End If

        Next

    Next
End Sub

Private Sub cmdClear_Click()

    txtMain.Text = ""

End Sub

Private Sub cmdLoadPic_Click()

    Dialog1.ShowOpen

    If Dialog1.FileName <> "" Then

        picMain.Picture = LoadPicture(Dialog1.FileName)
        txtMain.Text = ""
        

    End If

End Sub

Private Sub cmdSavePic_Click()

    Dialog2.ShowSave

    If Dialog2.FileName <> "" Then SavePicture picMain.Image, Dialog2.FileName

End Sub

Private Sub cmdWrite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If mHover Then SetInitialBTStyle cmdWrite
Timer1.Enabled = False
End Sub

Private Sub Form_Load()

    picMain.Picture = LoadPicture(App.Path & "/BlueSoft.bmp")

    GetInitialBTStyle cmdClear
    GetInitialBTStyle cmdLoadPic
    GetInitialBTStyle cmdLoadText
    GetInitialBTStyle cmdRead
    GetInitialBTStyle cmdSavePic
    GetInitialBTStyle cmdSaveText
    GetInitialBTStyle cmdWrite

    BTFlat cmdClear
    BTFlat cmdLoadPic
    BTFlat cmdLoadText
    BTFlat cmdRead
    BTFlat cmdSavePic
    BTFlat cmdSaveText
    BTFlat cmdWrite
    
    mHover = True
    i = 0
End Sub
    
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If mHover Then

        BTFlat cmdClear
        BTFlat cmdLoadPic
        BTFlat cmdLoadText
        BTFlat cmdRead
        BTFlat cmdSavePic
        BTFlat cmdSaveText
        BTFlat cmdWrite

    End If
Timer1.Enabled = True
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&

End Sub


Private Function ByteToBin(n As Integer) As String

    Dim j As String

    Do While n >= 1

        j = n Mod 2 & j
        n = n \ 2

    Loop

    If Len(j) < 8 Then j = String$(8 - Len(j), "0") & j

    ByteToBin = j

End Function

Private Function putc(c As String) As String

    Dim ps As String
    ps = frmMain.txtPassword.Text

    If ps <> "" Then

        Randomize Asc(Mid$(ps, 1 + Int(Len(ps) * Rnd), 1)) * (1 + Int(Len(ps) * Rnd)) * 13
        putc = Chr$(arrayA(Asc(c), 1 + Int(Len(ps) * Rnd)))

    Else

        putc = c

    End If

End Function

Private Function getc(c As String) As String

    Dim ps As String
    ps = frmMain.txtPassword.Text

    If ps <> "" Then

        Randomize Asc(Mid$(ps, 1 + Int(Len(ps) * Rnd), 1)) * (1 + Int(Len(ps) * Rnd)) * 13
        getc = Chr$(arrayB(Asc(c), 1 + Int(Len(ps) * Rnd)))

    Else

        getc = c

    End If

End Function

Private Sub Shuffle(pas As String)

    Dim i As Integer, j As Integer, k As Double, X As Integer, Y As Integer, t As Integer
    ln = Len(pas)
    Dim f As Long
    If ln > 0 Then

        k = 1

        For j = 1 To ln

            k = k + Asc(Mid$(pas, j, 1)) * j

        Next

        k = Sqr(k)
        ReDim arrayA(0 To 255, 1 To ln) As Integer
        ReDim arrayB(0 To 255, 1 To ln) As Integer

        For i = 1 To Len(pas)

            For j = 0 To 255

                arrayA(j, i) = j

            Next

        Next

        For j = 1 To ln

            f = Rnd(-1)
            Randomize Asc(Mid$(pas, j, 1)) * CDbl(j) * k

            For i = 1 To 10000

                Y = Int(255 * Rnd)
                t = 255 - Int(255 * Rnd)
                X = arrayA(Y, j)
                arrayA(Y, j) = arrayA(t, j)
                arrayA(t, j) = X

            Next

        Next

        For i = 1 To ln

            For j = 0 To 255

                arrayB(arrayA(j, i), i) = j

            Next

        Next

    End If

End Sub

Private Sub Timer1_Timer()
i = i + 1
Load Shape1(i)
Load Shape2(i)
Shape1(i).Left = Shape1(i - 1).Left + 4
Shape2(i).Left = Shape2(i - 1).Left - 4
Shape1(i).Visible = True
Shape2(i).Visible = True
If i = 60 Then
'unload all shapes
For intcntr = 1 To i
Unload Shape1(intcntr)
Unload Shape2(intcntr)
Next
'restore counter
i = 0
End If
'Debug.Print i
End Sub
