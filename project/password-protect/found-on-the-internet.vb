'
'   https://www.mrexcel.com/board/threads/lock-unlock-vbaprojects-programmatically-without-sendkeys.1136415/
'
Option Explicit


#If VBA7 Then
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hhk As LongPtr) As Long
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
    Private Declare PtrSafe Function GetDlgItem Lib "user32" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long) As LongPtr
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare PtrSafe Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As LongPtr, ByVal wFlag As Long) As LongPtr
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hwnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As String) As LongPtr
    Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function SetActiveWindow Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    
    Private lHook As LongPtr
#Else
    
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hhk As Long) As Long
    Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
    Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hwnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
    Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
    
    Private lHook As Long
#End If

Private sWinClassName As String, sWorkbookName As String, sPassword As String



Public Property Let LockVBProject(ByVal WorkbookName As String, ByVal Password As String, ByVal bLock As Boolean)

    #If VBA7 Then
        Dim hwnd As LongPtr
    #Else
        Dim hwnd As Long
    #End If
    
    Const WH_CBT = 5

    On Error GoTo errHandler

    hwnd = GetActiveWindow
    
    With Application.VBE
        Set .ActiveVBProject = Application.Workbooks(WorkbookName).VBProject
        If bLock Then
            If .ActiveVBProject.Protection = 0 Then
                sWinClassName = "VBAProject - Project Properties"
                sWorkbookName = WorkbookName
            Else
                MsgBox "VBProect already locked": Exit Property
            End If
        Else
            If .ActiveVBProject.Protection Then
                sWinClassName = "VBAProject Password"
            Else
                MsgBox "VBProect already unlocked": Exit Property
            End If
        End If
    End With
    
    sPassword = Password
    lHook = SetWindowsHookEx(WH_CBT, AddressOf Catch_DlgBox_Activation, 0, GetCurrentThreadId)
    Application.VBE.CommandBars(1).FindControl(ID:=2578, recursive:=True).Execute
    
    If hwnd = Application.hwnd Then
        SetActiveWindow Application.hwnd
    End If
    
Exit Property

errHandler:
    Call UnHook
    MsgBox "Runtime Error : " & Err.Number & vbCr & vbCr & Err.Description, vbExclamation
    
End Property



#If VBA7 Then
    Private Function Catch_DlgBox_Activation(ByVal idHook As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
        Dim hwnd As LongPtr
#Else
    Private Function Catch_DlgBox_Activation(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Dim hwnd As Long
#End If

    Const HCBT_ACTIVATE = 5
    Const SWP_HIDEWINDOW = &H80

    Dim sBuff As String * 256, lRet As Long
    
    If idHook = HCBT_ACTIVATE Then
    lRet = GetClassName(wParam, sBuff, 256)
    If Left(sBuff, lRet) = "#32770" Then
        sBuff = ""
        lRet = GetWindowText(wParam, sBuff, 256)
        If Left(sBuff, lRet) = sWinClassName Then
            Call UnHook
            SetWindowPos wParam, 0, 0, 0, 0, 0, SWP_HIDEWINDOW
            Call SetTimer(Application.hwnd, wParam, 0, AddressOf Protect_UnProtect_Routine)
        End If
    End If
    End If
    
    Catch_DlgBox_Activation = CallNextHookEx(lHook, idHook, ByVal wParam, ByVal lParam)
 
End Function


Private Sub UnHook()
    UnhookWindowsHookEx lHook
End Sub


#If VBA7 Then
Private Sub Protect_UnProtect_Routine(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal nIDEvent As LongPtr, ByVal dwTimer As Long)
    Dim hCurrentDlg As LongPtr, hwndSysTab As LongPtr
#Else
Private Sub Protect_UnProtect_Routine(ByVal hwnd As Long, ByVal uMsg As Long, ByVal nIDEvent As Long, ByVal dwTimer As Long)
    Dim hCurrentDlg As Long, hwndSysTab As Long
#End If

    Const TCM_FIRST = &H1300
    Const TCM_SETCURSEL = (TCM_FIRST + 12)
    Const TCM_SETCURFOCUS = (TCM_FIRST + 48)
    Const EM_SETMODIFY = &HB9
    Const BM_SETCHECK = &HF1
    Const BST_CHECKED = &H1
    Const BM_GETCHECK = &HF0
    Const BM_CLICK = &HF5
    Const WM_SETTEXT = &HC
    Const WH_CBT = 5
    Const GW_CHILD = 5
    
    On Error GoTo errHandler
    
    Call KillTimer(Application.hwnd, nIDEvent)
    
    hCurrentDlg = nIDEvent
    
    If sWinClassName = "VBAProject - Project Properties" Then
    
        hwndSysTab = FindWindowEx(hCurrentDlg, 0, "SysTabControl32", vbNullString)
        Call SendMessage(hwndSysTab, TCM_SETCURFOCUS, 1, 0)
        Call SendMessage(hwndSysTab, TCM_SETCURSEL, 1, 0)
        
        If SendMessage(GetDlgItem(GetNextWindow(hCurrentDlg, GW_CHILD), &H1557), BM_GETCHECK, 0, 0) = 0 Then
            Call SendMessage(GetDlgItem(GetNextWindow(hCurrentDlg, GW_CHILD), &H1557), BM_SETCHECK, BST_CHECKED, 0)
            Call SendMessage(GetDlgItem(GetNextWindow(hCurrentDlg, GW_CHILD), &H1555), WM_SETTEXT, 0, sPassword)
            Call SendMessage(GetDlgItem(GetNextWindow(hCurrentDlg, GW_CHILD), &H1555), EM_SETMODIFY, True, 0)
            Call SendMessage(GetDlgItem(GetNextWindow(hCurrentDlg, GW_CHILD), &H1556), WM_SETTEXT, 0, sPassword)
            Call SendMessage(GetDlgItem(GetNextWindow(hCurrentDlg, GW_CHILD), &H1556), EM_SETMODIFY, True, 0)
        End If
        
        Call SendMessage(GetDlgItem(hCurrentDlg, &H1), BM_CLICK, 0, 0)
        Call Application.OnTime(Now, "SaveVBProjectChanges")
        
    ElseIf sWinClassName = "VBAProject Password" Then
    
        Call SendMessage(GetDlgItem(hCurrentDlg, &H155E), WM_SETTEXT, 0, sPassword)
        Call SendMessage(GetDlgItem(hCurrentDlg, &H155E), EM_SETMODIFY, True, 0)
        lHook = SetWindowsHookEx(WH_CBT, AddressOf Catch_DlgBox_Creation, 0, GetCurrentThreadId)
        Call SendMessage(GetDlgItem(hCurrentDlg, &H1), BM_CLICK, 0, 0)
        Call Application.OnTime(Now, "UnHook")
        
    End If
    
    Exit Sub
    
errHandler:
    Call UnHook
    MsgBox "Runtime Error : " & Err.Number & vbCr & vbCr & Err.Description, vbExclamation


End Sub


#If VBA7 Then
    Private Function Catch_DlgBox_Creation(ByVal idHook As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
    Private Function Catch_DlgBox_Creation(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

    Const HCBT_CREATEWND = 3
    Dim sBuff As String * 256, lRet As Long
    
    If idHook = HCBT_CREATEWND Then
        lRet = GetClassName(wParam, sBuff, 256)
        If Left(sBuff, lRet) = "#32770" Then
            Catch_DlgBox_Creation = -1
            Exit Function
        End If
    End If
    
    Catch_DlgBox_Creation = CallNextHookEx(lHook, idHook, ByVal wParam, ByVal lParam)

End Function

Private Sub SaveVBProjectChanges()
    On Error Resume Next
    Application.EnableEvents = False
        Workbooks(sWorkbookName).Save
    Application.EnableEvents = True
End Sub
