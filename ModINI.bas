Attribute VB_Name = "ModINI"
Option Explicit

'sDefInitFileName is setup as (AppPath\AppEXEName.Ini)
'and is used as the Default Initialization Filename
Private sDefInitFileName As String

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Function GetInitEntry(ByVal sSection As String, ByVal sKeyName As String, Optional ByVal sDefault As String = "", Optional ByVal sInitFileName As String = "") As String
    
    'This Function Reads In a String From The Init File.
    'Returns Value From Init File or sDefault If No Value Exists.
    'sDefault Defaults to an Empty String ("").
    'Creates and Uses sDefInitFileName (AppPath\AppEXEName.Ini)
    'if sInitFileName Parameter Is Not Passed In.
    
    Dim sBuffer As String
    Dim sInitFile As String
    
    'If Init Filename NOT Passed In
    If Len(sInitFileName) = 0 Then
        'If Static Init FileName NOT Already Created
        If Len(sDefInitFileName) = 0 Then
            'Create Static Init FileName
            sDefInitFileName = App.Path
            If Right$(sDefInitFileName, 1) <> "\" Then
                sDefInitFileName = sDefInitFileName & "\"
            End If
            sDefInitFileName = sDefInitFileName & App.EXEName & ".ini"
        End If
        sInitFile = sDefInitFileName
    Else 'If Init Filename Passed In
        sInitFile = sInitFileName
    End If
    
    sBuffer = String$(999999, " ")
    GetInitEntry = Left$(sBuffer, GetPrivateProfileString(sSection, ByVal sKeyName, sDefault, sBuffer, Len(sBuffer), sInitFile))
    
End Function
Public Function SetInitEntry(ByVal sSection As String, Optional ByVal sKeyName As String, Optional ByVal sValue As String, Optional ByVal sInitFileName As String = "") As Long
    
    'This Function Writes a String To The Init File.
    'Returns WritePrivateProfileString Success or Error.
    'Creates and Uses sDefInitFileName (AppPath\AppEXEName.Ini)
    'if sInitFileName Parameter Is Not Passed In.
    
    '***** CAUTION *****
    'If sValue is Null then sKeyName is deleted from the Init File.
    'If sKeyName is Null then sSection is deleted from the Init File.
    
    Dim sInitFile As String
    
    'If Init Filename NOT Passed In
    If Len(sInitFileName) = 0 Then
        'If Static Init FileName NOT Already Created
        If Len(sDefInitFileName) = 0 Then
            'Create Static Init FileName
            sDefInitFileName = App.Path
            If Right$(sDefInitFileName, 1) <> "\" Then
                sDefInitFileName = sDefInitFileName & "\"
            End If
            sDefInitFileName = sDefInitFileName & App.EXEName & ".ini"
        End If
        sInitFile = sDefInitFileName
    Else 'If Init Filename Passed In
        sInitFile = sInitFileName
    End If
    
    If Len(sKeyName) > 0 And Len(sValue) > 0 Then
        SetInitEntry = WritePrivateProfileString(sSection, ByVal sKeyName, ByVal sValue, sInitFile)
    ElseIf Len(sKeyName) > 0 Then
        SetInitEntry = WritePrivateProfileString(sSection, ByVal sKeyName, vbNullString, sInitFile)
    Else
        SetInitEntry = WritePrivateProfileString(sSection, vbNullString, vbNullString, sInitFile)
    End If
    
End Function

