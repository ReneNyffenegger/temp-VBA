(ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
ByVal lpsz2 As String) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
(ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
 
Private Declare Function GetWindowTextLength Lib "user32" Alias _
"GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 
Dim Ret As Long, ChildRet As Long, OpenRet As Long
Dim strBuff As String, ButCap As String
Dim MyPassword As String

Const WM_SETTEXT = &HC
Const BM_CLICK = &HF5

Sub UnlockVBA()
    Dim xlAp As Object, oWb As Object
    
    Set xlAp = CreateObject("Excel.Application")
    
    xlAp.Visible = True
    
    '~~> Open the workbook in a separate instance
    Set oWb = xlAp.Workbooks.Open("C:\Sample.xlsm")
    
    '~~> Launch the VBA Project Password window
    '~~> I am assuming that it is protected. If not then
    '~~> put a check here.
    xlAp.VBE.CommandBars(1).FindControl(ID:=2578, recursive:=True).Execute

    '~~> Your passwword to open then VBA Project
    MyPassword = "Blah Blah"
    
    '~~> Get the handle of the "VBAProject Password" Window
    Ret = FindWindow(vbNullString, "VBAProject Password")
    
    If Ret <> 0 Then
        'MsgBox "VBAProject Password Window Found"
        
        '~~> Get the handle of the TextBox Window where we need to type the password
        ChildRet = FindWindowEx(Ret, ByVal 0&, "Edit", vbNullString)
        
        If ChildRet <> 0 Then
            'MsgBox "TextBox's Window Found"
            '~~> This is where we send the password to the Text Window
            SendMess MyPassword, ChildRet
        
            DoEvents
        
            '~~> Get the handle of the Button's "Window"
            ChildRet = FindWindowEx(Ret, ByVal 0&, "Button", vbNullString)
            
            '~~> Check if we found it or not
            If ChildRet <> 0 Then
                'MsgBox "Button's Window Found"
    
                '~~> Get the caption of the child window
                strBuff = String(GetWindowTextLength(ChildRet) + 1, Chr$(0))
                GetWindowText ChildRet, strBuff, Len(strBuff)
                ButCap = strBuff
    
                '~~> Loop through all child windows
                Do While ChildRet <> 0
                    '~~> Check if the caption has the word "OK"
                    If InStr(1, ButCap, "OK") Then
                        '~~> If this is the button we are looking for then exit
                        OpenRet = ChildRet
                        Exit Do
                    End If
    
                    '~~> Get the handle of the next child window
                    ChildRet = FindWindowEx(Ret, ChildRet, "Button", vbNullString)
                    '~~> Get the caption of the child window
                    strBuff = String(GetWindowTextLength(ChildRet) + 1, Chr$(0))
                    GetWindowText ChildRet, strBuff, Len(strBuff)
                    ButCap = strBuff
                Loop
    
                '~~> Check if we found it or not
                If OpenRet <> 0 Then
                    '~~> Click the OK Button
                    SendMessage ChildRet, BM_CLICK, 0, vbNullString
                Else
                    MsgBox "The Handle of OK Button was not found"
                End If
            Else
                 MsgBox "Button's Window Not Found"
            End If
        Else
            MsgBox "The Edit Box was not found"
        End If
    Else
        MsgBox "VBAProject Password Window was not Found"
    End If
End Sub

Sub SendMess(Message As String, hwnd As Long)
    Call SendMessage(hwnd, WM_SETTEXT, False, ByVal Message)
End Sub
