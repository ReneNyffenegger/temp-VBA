'
'  https://stackoverflow.com/a/27508116/180275
'
Option Explicit

Private Const PAGE_EXECUTE_READWRITE = &H40

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Long, Source As Long, ByVal Length As Long)

Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Long, _
        ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
        ByVal lpProcName As String) As Long

Private Declare Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As Long, _
        ByVal pTemplateName As Long, ByVal hWndParent As Long, _
        ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer

Dim HookBytes(0 To 5) As Byte
Dim OriginBytes(0 To 5) As Byte
Dim pFunc As Long
Dim Flag As Boolean

Private Function GetPtr(ByVal Value As Long) As Long
    GetPtr = Value
End Function

Public Sub RecoverBytes()
    If Flag Then MoveMemory ByVal pFunc, ByVal VarPtr(OriginBytes(0)), 6
End Sub

Public Function Hook() As Boolean
    Dim TmpBytes(0 To 5) As Byte
    Dim p As Long
    Dim OriginProtect As Long

    Hook = False

    pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")


    If VirtualProtect(ByVal pFunc, 6, PAGE_EXECUTE_READWRITE, OriginProtect) <> 0 Then

        MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal pFunc, 6
        If TmpBytes(0) <> &H68 Then

            MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal pFunc, 6

            p = GetPtr(AddressOf MyDialogBoxParam)

            HookBytes(0) = &H68
            MoveMemory ByVal VarPtr(HookBytes(1)), ByVal VarPtr(p), 4
            HookBytes(5) = &HC3

            MoveMemory ByVal pFunc, ByVal VarPtr(HookBytes(0)), 6
            Flag = True
            Hook = True
        End If
    End If
End Function

Private Function MyDialogBoxParam(ByVal hInstance As Long, _
        ByVal pTemplateName As Long, ByVal hWndParent As Long, _
        ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer
    If pTemplateName = 4070 Then
        MyDialogBoxParam = 1
    Else
        RecoverBytes
        MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, _
                           hWndParent, lpDialogFunc, dwInitParam)
        Hook
    End If
End Function

#if false then

another answer in the same SO question
Option Explicit

Private Const PAGE_EXECUTE_READWRITE = &H40

Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As LongPtr, Source As LongPtr, ByVal Length As LongPtr)

Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (lpAddress As LongPtr, _
ByVal dwSize As LongPtr, ByVal flNewProtect As LongPtr, lpflOldProtect As LongPtr) As LongPtr

Private Declare PtrSafe Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As LongPtr

Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, _
ByVal lpProcName As String) As LongPtr

Private Declare PtrSafe Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As LongPtr, _
ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, _
ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer

Dim HookBytes(0 To 11) As Byte
Dim OriginBytes(0 To 11) As Byte
Dim pFunc As LongPtr
Dim Flag As Boolean

Private Function GetPtr(ByVal Value As LongPtr) As LongPtr
    GetPtr = Value
End Function

Public Sub RecoverBytes()
    If Flag Then MoveMemory ByVal pFunc, ByVal VarPtr(OriginBytes(0)), 12
End Sub

Public Function Hook() As Boolean
    Dim TmpBytes(0 To 11) As Byte
    Dim p As LongPtr, osi As Byte
    Dim OriginProtect As LongPtr

    Hook = False

    #If Win64 Then
        osi = 1
    #Else
        osi = 0
    #End If

    pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")

    If VirtualProtect(ByVal pFunc, 12, PAGE_EXECUTE_READWRITE, OriginProtect) <> 0 Then

        MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal pFunc, osi+1
        If TmpBytes(osi) <> &HB8 Then

            MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal pFunc, 12

            p = GetPtr(AddressOf MyDialogBoxParam)

            If osi Then HookBytes(0) = &H48
            HookBytes(osi) = &HB8
            osi = osi + 1
            MoveMemory ByVal VarPtr(HookBytes(osi)), ByVal VarPtr(p), 4 * osi
            HookBytes(osi + 4 * osi) = &HFF
            HookBytes(osi + 4 * osi + 1) = &HE0

            MoveMemory ByVal pFunc, ByVal VarPtr(HookBytes(0)), 12
            Flag = True
            Hook = True
        End If
    End If
End Function

Private Function MyDialogBoxParam(ByVal hInstance As LongPtr, _
ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, _
ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer

    If pTemplateName = 4070 Then
        MyDialogBoxParam = 1
    Else
        RecoverBytes
        MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, _
                   hWndParent, lpDialogFunc, dwInitParam)
        Hook
    End If
End Function

#end if
