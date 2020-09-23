Attribute VB_Name = "ModCmnDlg"

Option Explicit
'Commondialog API - more efficient than using MS Common Dialog Control (comdlg32.ocx)
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXPLORER = &H80000
'UDT that makes calling the commondialog easier
Public Type CMDialog
    Ownerform As Long
    Filter As String
    Filetitle As String
    FilterIndex As Long
    Filename As String
    DefaultExtension As String
    OverwritePrompt As Boolean
    AllowMultiSelect As Boolean
    Initdir As String
    DialogTitle As String
    Flags As Long
End Type

' Private class type definitions
Private Type RECT
    Left     As Long
    Top      As Long
    Right    As Long
    Bottom   As Long
End Type

' Private class constants
Private Const WM_DESTROY = &H2
Private Const WM_SETFOCUS = &H7
Private Const WM_INITDIALOG = &H110
Private Const WM_COMMAND = &H111
Private Const WM_USER = &H400
Private Const WM_CHOOSEFONT_GETLOGFONT = (WM_USER + 1)

Private Const BN_CLICKED = 0&
Private Const BM_CLICK = &HF5

Private Const CB_GETCURSEL = &H147
Private Const CB_GETITEMDATA = &H150

Private Const CLR_INVALID = -1

 'Private class API function declarations
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const DWLP_DLGPROC = 4
Private Const GWLP_WNDPROC = (-4)

' SUBCLASSING  *********************
Private lpOrigDialogProc    As Long
Private mhwndApply          As Long
Private mhwndcboColor       As Long
Private mRegMsg             As Long   ' registered message ,value
' Focus Fix
Private mbSubClassBtns      As Boolean
Private lpOrigBtnOKProc     As Long
Private lpOrigBtnCancelProc As Long
Private mhwndOK             As Long
Private mhwndCancel         As Long

' Private variables used for communication
' and the callback routine implemented in this module
Private msDialogTitle       As String
Private mlLeft              As Long
Private mlTop               As Long
Private mlLastColor         As Long

Public cmndlg As CMDialog
'****************COMMONDIALOG CODE*********************
Public Sub ShowOpen()
    Dim OFName As OPENFILENAME
    Dim temp As String
    
    With cmndlg
        .Filter = "All Files (.*)|*.*"
        OFName.lStructSize = Len(OFName)
        OFName.hwndOwner = .Ownerform
        OFName.hInstance = App.hInstance
        OFName.lpstrFilter = Replace(.Filter, "|", Chr(0))
        OFName.lpstrFile = Space$(254)
        OFName.nMaxFile = 255
        OFName.lpstrFileTitle = Space$(254)
        OFName.nMaxFileTitle = 255
        OFName.lpstrInitialDir = .Initdir
        OFName.lpstrTitle = .DialogTitle
        OFName.nFilterIndex = .FilterIndex
        OFName.Flags = .Flags Or OFN_EXPLORER Or IIf(.AllowMultiSelect, OFN_ALLOWMULTISELECT, 0)
        If GetOpenFileName(OFName) Then
            .FilterIndex = OFName.nFilterIndex
            If .AllowMultiSelect Then
                temp = Replace(Trim$(OFName.lpstrFile), Chr(0), ";")
                If Right(temp, 2) = ";;" Then temp = Left(temp, Len(temp) - 2)
                .Filename = temp
            Else
                .Filename = StripTerminator(Trim$(OFName.lpstrFile))
                .Filetitle = StripTerminator(Trim$(OFName.lpstrFileTitle))
            End If
        Else
            .Filename = ""
        End If
    End With

End Sub
Public Sub ShowSave()
    Dim OFName As OPENFILENAME
    
    With cmndlg
        '.Filter = "Bitmap (.bmp)|*.bmp|Jpeg (.jpg)|*.jpg"
        OFName.lStructSize = Len(OFName)
        OFName.hwndOwner = .Ownerform
        OFName.hInstance = App.hInstance
        OFName.lpstrFilter = Replace(.Filter, "|", Chr(0))
        OFName.nMaxFile = 255
        OFName.lpstrFileTitle = Space$(254)
        OFName.nMaxFileTitle = 255
        OFName.lpstrInitialDir = .Initdir
        OFName.lpstrTitle = .DialogTitle
        OFName.nFilterIndex = .FilterIndex
        OFName.lpstrDefExt = .DefaultExtension
        OFName.lpstrFile = .Filename & Space$(254 - Len(.Filename))
        OFName.Flags = .Flags Or IIf(.OverwritePrompt, OFN_OVERWRITEPROMPT, 0)
       
        If GetSaveFileName(OFName) Then
            .Filename = StripTerminator(Trim$(OFName.lpstrFile))
            .Filetitle = StripTerminator(Trim$(OFName.lpstrFileTitle))
            .FilterIndex = OFName.nFilterIndex
        Else
            .Filename = ""
        End If
    End With
End Sub

'****************STRING FUNCTIONS*********************
Public Function StripTerminator(ByVal strString As String) As String
    'Removes chr(0)'s from the end of a string
    'API tends to do this
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

' Common Dialog callback function
Public Function CFDialogProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Const sREGMSG = "WOWLFCHANGE"   ' Undocumented registered message ,string
  Dim rcHeight     As Long
  Dim rcWidth      As Long
  Dim rc           As RECT
  Dim rcDesk       As RECT
  Dim idxColor     As Integer
  Dim lCurColor    As Long
  Dim CurrentFont  As StdFont
  'Dim fntCurrent   As LOGFONT
'  Static fntLast   As LOGFONT
  
  CFDialogProc = 0
  
  Select Case uMsg
  
      Case WM_INITDIALOG
          
          ' set the new title
          If Len(msDialogTitle) Then
              SetWindowText hwnd, msDialogTitle
          End If
      
          GetWindowRect GetDesktopWindow, rcDesk
          rc.Left = Abs(((rcDesk.Right - rcDesk.Left) - rcWidth) / 2)
          rc.Top = Abs(((rcDesk.Bottom - rcDesk.Top) - rcHeight) / 2)
                  
          MoveWindow hwnd, rc.Left, rc.Top, rcWidth, rcHeight, 1
              
          ' get hwnd of Apply Button(ControlID = &H402) once for use with registered message
          mhwndApply = GetDlgItem(hwnd, &H402)
          
          ' get hwnd of Color combo (ControlID = &H473) once for use with WM_COMMAND
          mhwndcboColor = GetDlgItem(hwnd, &H473)
          
          ' get value for registered message "WOWLFCHANGE"
          mRegMsg = RegisterWindowMessage(sREGMSG)
          
          ' Focus Fix
          If mbSubClassBtns Then
              ' get hwnd of OK & Cancel buttons
              mhwndOK = GetDlgItem(hwnd, 1)
              mhwndCancel = GetDlgItem(hwnd, 2)
              
              ' start subclassing buttons
              lpOrigBtnOKProc = SetWindowLong(mhwndOK, GWLP_WNDPROC, AddressOf BtnOKProc)
              lpOrigBtnCancelProc = SetWindowLong(mhwndCancel, GWLP_WNDPROC, AddressOf BtnCancelProc)
          End If
          
      Case mRegMsg
          ' Undocumented registered message "WOWLFCHANGE", return 0.
          ' User clicked a control: static control holds current font.
          ' Simulate user clicked Apply button, catch BN_CLICKED.
          SendMessage mhwndApply, BM_CLICK, 0, ByVal 0
          
      Case WM_COMMAND
          ' look for BN_CLICKED from Apply Button
          ' wparam: low word holds control ID, high word holds notification
          ' lparam: hwnd control
          
      Case WM_DESTROY
          
          ' Focus Fix: stop subclassing buttons
          If mbSubClassBtns Then
              If lpOrigBtnOKProc Then
                  SetWindowLong mhwndOK, GWLP_WNDPROC, lpOrigBtnOKProc
                  lpOrigBtnOKProc = 0
              End If
              If lpOrigBtnCancelProc Then
                  SetWindowLong mhwndCancel, GWLP_WNDPROC, lpOrigBtnCancelProc
                  lpOrigBtnCancelProc = 0
              End If
          End If
          
          ' stop subclassing Common Dialog
          SetWindowLong hwnd, DWLP_DLGPROC, lpOrigDialogProc
  End Select
End Function

' OK Button callback function
Public Function BtnOKProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  BtnOKProc = CallWindowProc(lpOrigBtnOKProc, hwnd, uMsg, wParam, lParam)
  
  If uMsg = WM_SETFOCUS Then
      SendMessage mhwndOK, BM_CLICK, 0, ByVal 0
  End If
  
End Function

' Cancel Button callback function
Public Function BtnCancelProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  BtnCancelProc = CallWindowProc(lpOrigBtnCancelProc, hwnd, uMsg, wParam, lParam)

  If uMsg = WM_SETFOCUS Then
      SendMessage mhwndCancel, BM_CLICK, 0, ByVal 0
  End If
  
End Function

Private Function Get_HiWord(ByRef lThis As Long) As Long
  If (lThis And &H80000000) = &H80000000 Then
     Get_HiWord = ((lThis And &H7FFF0000) \ &H10000) Or &H8000&
  Else
     Get_HiWord = (lThis And &HFFFF0000) \ &H10000
  End If
End Function

Private Function Get_LoWord(ByRef lThis As Long) As Long
  Get_LoWord = (lThis And &HFFFF&)
End Function


Public Function CheckExt(Filename As String, ext As String)
    Dim stg1 As String
    Dim stg2 As String
    
    stg1 = Right$(Filename, 4)                                      ' get extension
    stg2 = Left$(Filename, Len(Filename) - 4)                       ' get filename without extension
    
    If InStr(Filename, ".") = False Then                            ' if no extension present
       CheckExt = Filename & ext                                    ' add extension to filename
    Else
       CheckExt = stg2 & ext                                        ' Makes sure we have correct extension
    End If
      
End Function

Public Function SaveJPEG(ByVal Filename As String, pic As PictureBox, PForm As Form, Optional ByVal Overwrite As Boolean = True, Optional ByVal Quality As Byte = 90) As Boolean
    Dim JPEGclass As cJpeg
    Dim m_Picture As IPictureDisp
    Dim m_DC As Long
    Dim m_Millimeter As Single
    m_Millimeter = PForm.ScaleX(100, vbPixels, vbMillimeters)
    Set m_Picture = pic
    m_DC = pic.hDC
    'this is not my code....from PSC
    'initialize class
    Set JPEGclass = New cJpeg
    'check there is image to save and the filename string is not empty
    If m_DC <> 0 And LenB(Filename) > 0 Then
        'check for valid quality
        If Quality < 1 Then Quality = 1
        If Quality > 100 Then Quality = 100
        'set quality
        JPEGclass.Quality = Quality
        'save in full color
        JPEGclass.SetSamplingFrequencies 1, 1, 1, 1, 1, 1
        'copy image from hDC
        If JPEGclass.SampleHDC(m_DC, CLng(m_Picture.Width / m_Millimeter), CLng(m_Picture.Height / m_Millimeter)) = 0 Then
            'if overwrite is set and file exists, delete the file
            If Overwrite And LenB(Dir$(Filename)) > 0 Then Kill Filename
            'save file and return True if success
            SaveJPEG = JPEGclass.SaveFile(Filename) = 0
        End If
    End If
    'clear memory
    Set JPEGclass = Nothing
End Function

Public Sub BitmapSave(pic As PictureBox)
   cmndlg.DialogTitle = "Save As Bitmap"
   cmndlg.Flags = cmndlg.OverwritePrompt
   cmndlg.Filter = "Bitmap (*.bmp)|*.bmp"                     ' sets the file type
   ShowSave                                                                   'show save dialog
   If cmndlg.Filename = "" Then Exit Sub
   cmndlg.Filename = CheckExt(cmndlg.Filename, ".bmp")
   pic.Picture = pic.Image                                 ' make sure the picture is there to save
   SavePicture pic.Picture, cmndlg.Filename            'save as bitmap
   MsgBox "Bitmap saved in folder " & cmndlg.Filename
End Sub

Public Sub JpegSave(pic As PictureBox)
Dim iresponse As String
Dim Fname As String
   On Error GoTo JpgErr
   pic.Picture = pic.Image
   cmndlg.DialogTitle = "Save As Jpeg"
   cmndlg.Flags = cmndlg.OverwritePrompt
   cmndlg.Filter = "Jpeg (*.jpg)|*.jpg"                                ' sets the file type
   '===================
   ShowSave
   '===================
   If cmndlg.Filename = "" Then Exit Sub
   cmndlg.Filename = CheckExt(cmndlg.Filename, ".jpg")

   If SaveJPEG(cmndlg.Filename, pic, Form1, True, 90) = True Then  ' save pic as Jpeg
      MsgBox "Jpeg saved in folder " & cmndlg.Filename
   End If
JpgErr:
   Exit Sub
End Sub
