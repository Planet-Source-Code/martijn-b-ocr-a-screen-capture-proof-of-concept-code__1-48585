VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "ScreenOCR"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   145
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   612
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picText 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   599
      TabIndex        =   3
      Top             =   720
      Width           =   9015
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   7320
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.bmp"
      DialogTitle     =   "Select a Picture"
      Filter          =   "Pictures (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg"
   End
   Begin VB.CommandButton cmdChangeFont 
      Caption         =   "Change &Font"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.PictureBox picCharScan 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   5400
      ScaleHeight     =   151
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Capture"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private lCharWidth As Long
Private lCharHeight As Long
Private arrCharVal(0& To 1&, 0& To 255&) As Variant


Private Sub cmdChangeFont_Click()

  With dlgMain

    .Flags = cdlCFBoth
    subSetFont picCharScan, dlgMain
    .ShowFont

    If .FontName <> vbNullString Then

      If .FontSize > 18 Then

        MsgBox "The code does not support font sizes above 18 point."
        Exit Sub

      End If

  '***  set the font
      subSetFont dlgMain, picCharScan
      subSetFont dlgMain, picText
  '***  get the values for this font
      subGetCharValues picCharScan

    End If

  End With

End Sub


Private Sub subGetCharValues(ctl As PictureBox)

  '***  stores the values for all characters in arrCharVal(0&,)
  Dim lCount As Long
  Dim szChars As SizeAPI

  With ctl

    .BorderStyle = 0&
    .Appearance = 0&
    .AutoRedraw = True
    .BackColor = vbWhite
    .ForeColor = vbBlack
  '***  this will set the range of characters used
  '***  for standard ascii change:
  '***  For lCount = to 31& to 127&
  '***  reset the maximum values
    lCharWidth = 0
    lCharHeight = 0

    For lCount = 0& To 30&

      arrCharVal(0&, lCount) = Array(-1&)
      arrCharVal(1&, lCount) = lCount

    Next

    For lCount = 31& To 255&

  '***  get the font dimensions. this used to be
  '***  .Width = .TextWidth("A")
  '***  .Height = .TextHeight("A")
  '***  now I use this for better results with the Terminal Font.
      szChars = GetTextSize(picCharScan.Font, Chr$(lCount))
      .Width = szChars.Width
      .Height = szChars.Height
  '***  store the maximum values

      If .Width > lCharWidth Then lCharWidth = .Width

      If .Height > lCharHeight Then lCharHeight = .Height

      .Cls
      ctl.Print Chr$(lCount);
      arrCharVal(0&, lCount) = fcnGetCharArray(ctl)
      arrCharVal(1&, lCount) = lCount

    Next

  End With

  lCharWidth = lCharWidth - 1&
  lCharHeight = lCharHeight - 1&
  'subSortCharVal

End Sub


Private Function fcnGetCharArray(ctl As PictureBox) As Long()

  Dim lX As Long, lY As Long
  Dim lValue() As Long
  Dim lValue2() As Long
  Dim lPower() As Long
  Dim lStart As Long
  Dim lEnd As Long
  Dim lCount As Long
  Dim lPos As Long
  lPower() = fcnGetPower()
  ReDim lValue(0 To ctl.Width - 1) As Long

  With ctl

    .ScaleMode = vbPixels
  '***  get the values

    For lX = 0& To .ScaleWidth - 1

      For lY = 0& To .ScaleHeight - 1

        If GetPixel(.hDC, lX, lY) = vbBlack Then

          lValue(lX) = lValue(lX) + lPower(lY)

        End If

      Next

    Next

  End With

  lStart = LBound(lValue)
  lEnd = UBound(lValue)
  lValue2 = lValue
  lPos = 0

  '***  remove leading zero values
  For lCount = lStart To lEnd

    If lValue(lCount) = 0 And lPos = 0& Then

    Else

      lValue2(lPos) = lValue(lCount)
      lPos = lPos + 1&

    End If

  Next

  If lPos = 0& Then

  '***  space character..
    lValue2 = lValue

  Else

    ReDim Preserve lValue2(0& To lPos - 1&)
    '***  remove trailing zero values
    Do Until lValue2(UBound(lValue2)) <> 0

      ReDim Preserve lValue2(0& To UBound(lValue2) - 1&)

    Loop

  End If

  '***  the binary character
  fcnGetCharArray = lValue2()

End Function


Private Function fcnGetPower() As Long()

  '***  create an array with 2^x values
  Dim lX As Long
  Dim lPower() As Long
  ReDim lPower(0& To 30&)

  For lX = 0& To 30&

    lPower(lX) = 2 ^ lX

  Next

  fcnGetPower = lPower

End Function


Private Sub cmdScan_Click()

  Dim ctl As PictureBox
  Dim lX As Long, lY As Long
  Dim lLine() As Long
  Dim lPower() As Long
  Dim lPicWidth As Long
  Dim lPicHeight As Long
  Dim lvLine As Long
  Dim sReturn As String
  Dim sOutput As String
  Dim lskip As Long
  lPower() = fcnGetPower()
  Set ctl = picText

  With ctl

    .Cls
  '***  create the picture with the text
    ctl.Print "...Hello! I just like to know how things work..."
    ctl.Print "abcdefghijklmnopqrstuvwxyz"
    ctl.Print UCase$("abcdefghijklmnopqrstuvwxyz")
    ctl.Print "!@#$%^&*()"
    ctl.Print "1234567890"
    ctl.Refresh
    DoEvents
    .ScaleMode = vbPixels
  '***  get the values
    lPicWidth = .ScaleWidth - 1
    lPicHeight = .ScaleHeight - 1
    ReDim lLine(0 To lPicWidth) As Long
    Debug.Assert lPicHeight > lCharHeight
  '***  now get the first line

    For lX = 0& To lPicWidth

      For lY = 0& To lCharHeight

        If GetPixel(.hDC, lX, lY) = vbBlack Then

          lLine(lX) = lLine(lX) + lPower(lY)

        End If

      Next

    Next
    
    sReturn = fcnScanLine(lLine)

    If sReturn <> vbNullString Then

      lskip = lCharHeight
      sOutput = sOutput & sReturn & vbNewLine

    End If

    For lvLine = lCharHeight + 1& To .ScaleHeight
      lblProgress.Caption = Round((lvLine / .ScaleHeight) * 100) & "%"
      lblProgress.Refresh
      For lX = 0& To lPicWidth
        
        '***  shift the bits 'up'
        lLine(lX) = lLine(lX) \ 2&

        If GetPixel(.hDC, lX, lvLine) = vbBlack Then
          
          '***  add a black pixel value
          lLine(lX) = lLine(lX) + lPower(lCharHeight)

        End If

      Next

      If lskip = 0 Then

        sReturn = fcnScanLine(lLine)

        If sReturn <> vbNullString Then

          lskip = lCharHeight
          sOutput = sOutput & sReturn & vbNewLine

        End If

      Else
        
        '***  This line was already scanned.
        lskip = lskip - 1

      End If

    Next

  End With

  MsgBox sOutput

End Sub


Private Sub subSetFont(oFrom As Object, oTo As Object)

  With oTo

    .FontName = oFrom.FontName
    .FontSize = oFrom.FontSize
    .FontBold = oFrom.FontBold
    .FontItalic = oFrom.FontItalic
    .FontUnderline = oFrom.FontUnderline

  End With

End Sub


Private Sub Form_Load()

  '***  set the default font
  picCharScan.FontName = "Tahoma"
  picCharScan.FontSize = 10
  subSetFont picCharScan, picText
  subSetFont picText, dlgMain
  '***  get the values for this font
  subGetCharValues picCharScan

End Sub


Private Function fcnScanLine(lLine() As Long) As String

  Dim lStart As Long
  Dim lEnd As Long
  Dim lCStart As Long
  Dim lCEnd As Long
  Dim lCount As Long
  Dim lPos As Long
  Dim sOutput As String
  Dim lChar As Long
  Dim lCharCount As Long
  Dim lBestSize As Long
  lStart = LBound(lLine)
  lEnd = UBound(lLine)
  lPos = lStart

  Do

    lChar = -1
    lBestSize = -1
    
    '***  adjust the scanning range here.
    For lCount = 31& To 128&

      If arrCharVal(0&, lCount)(0) = lLine(lPos) Then

  '*** check this character
        lCStart = LBound(arrCharVal(0&, lCount)) + 1&
        lCEnd = UBound(arrCharVal(0&, lCount))

        For lCharCount = lCStart To lCEnd

          If lPos + lCharCount <= lEnd Then

            If arrCharVal(0&, lCount)(lCharCount) <> lLine(lPos + lCharCount) Then

              Exit For

            End If

          End If

        Next

        If lCharCount > lCEnd Then
          
          '***  found a match
          '***  check if it is 'bigger' than the last match (if any)
          '***  an 'O' could be recognized as a smaller 'C'
          If lBestSize < lCharCount Then

            lChar = arrCharVal(1&, lCount)
            lBestSize = lCharCount

          End If

        End If

      End If

    Next

  '***  did we find a match?

    If lChar <> -1 Then

      sOutput = sOutput & Chr$(lChar)
      '***  move right a few pixels
      lPos = lPos + lBestSize

    Else

      lPos = lPos + 1&

    End If

  Loop Until lPos >= lEnd

  fcnScanLine = Trim$(sOutput)

End Function

