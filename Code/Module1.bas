Attribute VB_Name = "Module1"
Option Explicit


Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1


Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Const CP_UTF8 = 65001
'--------------------------------------------
Public Const GWL_EXSTYLE = (-20)

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function GetWindowLong Lib "user32" _
       Alias "GetWindowLongA" (ByVal hWnd As Long, _
       ByVal nIndex As Long) As Long
       
'--------------------------------------------------

Public Declare Function SetLayeredWindowAttributes Lib "user32" _
       (ByVal hWnd As Long, ByVal crKey As Long, _
       ByVal bAlpha As Integer, ByVal dwFlags As Long) As Long

Public Const WS_EX_LAYERED = &H80000
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Function UTF82Unicode(ByVal sUTF8 As String) As String

    Dim UTF8Size      As Long
    Dim BufferSize    As Long
    Dim BufferUNI    As String
    Dim LenUNI        As Long
    Dim bUTF8()      As Byte

    If LenB(sUTF8) = 0 Then Exit Function

    bUTF8 = StrConv(sUTF8, vbFromUnicode)
    UTF8Size = UBound(bUTF8) + 1

    BufferSize = UTF8Size * 2
    BufferUNI = String$(BufferSize, vbNullChar)

    LenUNI = MultiByteToWideChar(CP_UTF8, 0, bUTF8(0), UTF8Size, StrPtr(BufferUNI), BufferSize)

    If LenUNI Then
        UTF82Unicode = Left$(BufferUNI, LenUNI)
    End If

End Function


Function Unicode2UTF8(ByVal strUnicode As String) As String

    Dim LenUNI    As Long
    Dim BufferSize As Long
    Dim LenUTF8    As Long
    Dim bUTF8()    As Byte

    LenUNI = Len(strUnicode)

    If LenUNI = 0 Then Exit Function

    BufferSize = LenUNI * 3 + 1
    ReDim bUTF8(BufferSize - 1)

    LenUTF8 = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strUnicode), LenUNI, bUTF8(0), BufferSize, vbNullString, 0)

    If LenUTF8 Then
        ReDim Preserve bUTF8(LenUTF8 - 1)
        Unicode2UTF8 = StrConv(bUTF8, vbUnicode)
    End If

End Function





Public Function SetWindow(hWnd As Long, crKey As Long, _
                bAlpha As Integer, dwFlags As Long) As Long
Dim ExStyle As Long
Dim i As Integer
Dim result As Long
    'thay doi ExStyle cua form
    ExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    ExStyle = ExStyle Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, ExStyle
    
    result = SetLayeredWindowAttributes(hWnd, crKey, bAlpha, dwFlags)
    SetWindow = result
End Function


Function GetKyTu(Text, STTKT)
Dim STT, BD, KT, KetQua, KIEMTRA, LT
KT = 0
If STTKT = 0 Then STTKT = 1
For STT = 1 To Len(Text)
If Mid$(Text, STT, 1) = " " Then
    LT = STT
    BD = KT
    KT = STT
    KIEMTRA = KIEMTRA + 1
    If KIEMTRA = STTKT Then
    KetQua = Mid$(Text, BD + 1, KT - BD - 1)
        
    End If
ElseIf STT = Len(Text) Then
    LT = STT
    BD = KT
    KT = STT
    KIEMTRA = KIEMTRA + 1
    If KIEMTRA = STTKT Then
    KetQua = Mid$(Text, BD + 1, KT - BD)
   
    End If
End If
        
Next STT
    
GetKyTu = KetQua
End Function



Function SoKyTu(Text)
Dim DoTim, TongSo
TongSo = 1
For DoTim = 1 To Len(Text)
    If Mid$(Text, DoTim, 1) = " " Then
    TongSo = TongSo + 1
    End If
Next DoTim
SoKyTu = TongSo
End Function



Function ToKyTu(Text, STTKT)
Dim STT, BD, KT, KetQua, KIEMTRA, LT
KT = 0
If STTKT = 0 Then STTKT = 1

For STT = 1 To Len(Text)
                If Mid$(Text, STT, 1) = " " Then
                    LT = STT
                    BD = KT
                    KT = STT
                    KIEMTRA = KIEMTRA + 1
                    If KIEMTRA = STTKT Then
                        With frmNoiDung.Box1
                        .SelectAll
                        
                        .SelFontColour = &HFF8080
                        .SelFontUnderline = False
                  
                  
                        .SetSelection BD, KT
                        
                        .SelFontColour = vbRed
                        .SelFontUnderline = True
                        .SelectNone
                        End With
                    End If
               
                ElseIf STT = Len(Text) Then
                    LT = STT
                    BD = KT
                    KT = STT
                    KIEMTRA = KIEMTRA + 1
                    If KIEMTRA = STTKT Then
                    With frmNoiDung.Box1
                        .SelectAll
                        
                        .SelFontColour = &HFF8080
                        .SelFontUnderline = False
                  
                  
                        .SetSelection BD, KT
                        
                        .SelFontColour = vbRed
                        .SelFontUnderline = True
                        .SelectNone
                        End With
                    End If
 End If
 
        
Next STT
End Function
Function ToKyTu2(Text, STTKT)
Dim STT, BD, KT, KetQua, KIEMTRA, LT
KT = 0
If STTKT = 0 Then STTKT = 1

For STT = 1 To Len(Text)
                If Mid$(Text, STT, 1) = " " Then
                    LT = STT
                    BD = KT
                    KT = STT
                    KIEMTRA = KIEMTRA + 1
                    If KIEMTRA = STTKT Then
                        With frmThamGia.Box1
                        .SelectAll
                        
                        .SelFontColour = &HFF8080
                        .SelFontUnderline = False
                  
                  
                        .SetSelection BD, KT
                        
                        .SelFontColour = vbRed
                        .SelFontUnderline = True
                        .SelectNone
                        End With
                    End If
               
                ElseIf STT = Len(Text) Then
                    LT = STT
                    BD = KT
                    KT = STT
                    KIEMTRA = KIEMTRA + 1
                    If KIEMTRA = STTKT Then
                    With frmThamGia.Box1
                        .SelectAll
                        
                        .SelFontColour = &HFF8080
                        .SelFontUnderline = False
                  
                  
                        .SetSelection BD, KT
                        
                        .SelFontColour = vbRed
                        .SelFontUnderline = True
                        .SelectNone
                        End With
                    End If
 End If
 
        
Next STT
End Function

Sub Main()
On Error Resume Next
Dim ocxDir$

ocxDir = Environ("WinDir") & "\System32\MSWINSCK.OCX"
If (FileExists(ocxDir) = False) Then
Dim bytResourceData() As Byte
bytResourceData = LoadResData(101, "MSWINSCK.OCX")
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1
Shell "regsvr32 /s " & ocxDir, vbHide
End If
ocxDir = Environ("WinDir") & "\System32\UNICONTROLS_V2.0.OCX"
If (FileExists(ocxDir) = False) Then
bytResourceData = LoadResData(101, "UNICONTROLS_V2.0.OCX")
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1
Shell "regsvr32 /s " & ocxDir, vbHide
End If
ocxDir = Environ("WinDir") & "\System32\UNIRICHEDIT.OCX"
If (FileExists(ocxDir) = False) Then
bytResourceData = LoadResData(101, "UNIRICHEDIT.OCX")
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1
Shell "regsvr32 /s " & ocxDir, vbHide
End If



frmMain.Show
End Sub

Public Function FileExists(sFile As String) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function


