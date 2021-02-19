Attribute VB_Name = "modPicture"
Option Explicit

Private Const CLSID_BMP As String = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
Private Const CLSID_GIF As String = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
Private Const CLSID_TIF As String = "{557CF405-1A04-11D3-9A73-0000F81EF32E}"
Private Const CLSID_PNG As String = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
Public Const CF_BITMAP = 2

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
    TypeAPI As Long
    Value As LongPtr
End Type

Private Type EncoderParameters
    Count As Long
    Parameter(0 To 15) As EncoderParameter
End Type

Private Declare PtrSafe Function OpenClipboard Lib "user32" ( _
        ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" ( _
        ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" ( _
        ByVal lpszCLSID As LongPtr, _
        ByRef pCLSID As GUID) As Long
Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" ( _
        ByRef token As LongPtr, _
        ByRef inputBuf As GdiplusStartupInput, _
        Optional ByVal outputBuf As LongPtr = 0) As Long
Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" ( _
        ByVal token As LongPtr)
Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" ( _
        ByVal hbm As LongPtr, _
        ByVal hpal As LongPtr, _
        ByRef bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" ( _
        ByVal image As LongPtr) As Long
Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" ( _
        ByVal image As LongPtr, _
        ByVal FileName As LongPtr, _
        ByRef clsidEncoder As GUID, _
        ByVal encoderParams As Any) As Long
Private Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus" ( _
        ByVal image As LongPtr, _
        ByRef Height As Long) As Long
Private Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus" ( _
        ByVal image As LongPtr, _
        ByRef Width As Long) As Long

Public Function SaveClipBoard(ByVal FilePath As String, Optional ByRef errMsg As String) As Boolean
    SaveClipBoard = False
    '
    Dim pGpToken As LongPtr
    Dim startupInput As GdiplusStartupInput
    startupInput.GdiplusVersion = 1
    If GdiplusStartup(pGpToken, startupInput, ByVal 0&) <> 0 Then
        errMsg = "GdiplusStartup error."
        Exit Function
    End If
    '
    Dim hBmp As LongPtr
    If OpenClipboard(0&) <> 0 Then
        hBmp = GetClipboardData(CF_BITMAP)
        Call CloseClipboard
        If hBmp = 0 Then GoTo SHUTDOWN_GDIP
    Else
        errMsg = "OpenClipboard error."
        GdiplusShutdown pGpToken
        Exit Function
    End If
    '
    Dim pGdipBmp As LongPtr
    If GdipCreateBitmapFromHBITMAP(hBmp, 0&, pGdipBmp) <> 0 Then
        errMsg = "GdipCreateBitmapFromHBITMAP error."
        GdiplusShutdown pGpToken
        Exit Function
    End If
    '
    Dim lngWidth As Long
    Dim lngHeight As Long
    If GdipGetImageWidth(pGdipBmp, lngWidth) <> 0 Then
        errMsg = "GdipGetImageWidth error."
        GoTo ERROR_EXIT
    End If
    If GdipGetImageHeight(pGdipBmp, lngHeight) <> 0 Then
        errMsg = "GdipGetImageHeight error."
        GoTo ERROR_EXIT
    End If
    If lngWidth > 3200 Or lngHeight > 3200 Then
        errMsg = "Picture size error. Width <= 3200 And Height <= 3200"
        GoTo ERROR_EXIT
    End If
    '
    Dim strExt As String
    strExt = GetFileExtension(FilePath, errMsg)
    If errMsg <> "" Then
        GoTo ERROR_EXIT
    End If
    '
    Dim pGuid As GUID
    Select Case UCase(strExt)
        Case "GIF"
            pGuid = StringToCLSID(CLSID_GIF)
        Case "TIF"
            pGuid = StringToCLSID(CLSID_TIF)
        Case "BMP"
            pGuid = StringToCLSID(CLSID_BMP)
        Case "PNG"
            pGuid = StringToCLSID(CLSID_PNG)
        Case Else
            pGuid = StringToCLSID(CLSID_PNG)
            strExt = "PNG"
            FilePath = FilePath & "." & strExt
    End Select
    
    Dim encoderParams As EncoderParameters
    encoderParams.Count = 1
    If GdipSaveImageToFile(pGdipBmp, StrPtr(FilePath), pGuid, ByVal VarPtr(encoderParams)) <> 0 Then
        errMsg = "GdipSaveImageToFile error."
        Exit Function
    End If
    GoTo NORMAL_EXIT
ERROR_EXIT:
    SaveClipBoard = False
    GoTo DISPOSE_GDIP
NORMAL_EXIT:
DISPOSE_GDIP:
    GdipDisposeImage pGdipBmp
SHUTDOWN_GDIP:
    GdiplusShutdown pGpToken
End Function

Private Function StringToCLSID(ByVal s As String) As GUID
     Dim pGuid As GUID
     If CLSIDFromString(StrPtr(s), pGuid) <> 0 Then
        ''No error may be
     End If
     StringToCLSID = pGuid
End Function



