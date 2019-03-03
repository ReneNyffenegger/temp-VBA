MessageBoxA
--------------
7759FD1E  mov         edi,edi  
7759FD20  push        ebp  
7759FD21  mov         ebp,esp  
7759FD23  push        0  
7759FD25  push        dword ptr [ebp+14h]  
7759FD28  push        dword ptr [ebp+10h]  
7759FD2B  push        dword ptr [ebp+0Ch]  
7759FD2E  push        dword ptr [ebp+8]  
7759FD31  call        7759FCD6  
7759FD36  pop         ebp  
7759FD37  ret         10h  


After hooking
==============


orig_MessageBoxA:                                                                             <<< 3) --- hook calls orig // Should be called trampoline
------------------
7752002C 8B FF                mov         edi,edi  
7752002E 55                   push        ebp  
7752002F 8B EC                mov         ebp,esp  
77520031 E9 ED FC 07 00       jmp         7759FD23  


7759FD1E E9 4D E1 DC 88       jmp         hook_MessageBoxA (036DE70h)                           <<< 1) ---- Client tries to call Message Box, ends up here.
7759FD23 6A 00                push        0  
7759FD25 FF 75 14             push        dword ptr [ebp+14h]  
7759FD28 FF 75 10             push        dword ptr [ebp+10h]  
7759FD2B FF 75 0C             push        dword ptr [ebp+0Ch]  
7759FD2E FF 75 08             push        dword ptr [ebp+8]  
7759FD31 E8 A0 FF FF FF       call        7759FCD6  
7759FD36 5D                   pop         ebp  
7759FD37 C2 10 00             ret         10h  




int WINAPI hook_MessageBoxA(HWND hWnd, LPCSTR lpText, LPCSTR lpCaption, UINT uType) {          <<< 2) ---- Jump instruction comes here
0036DE70 55                   push        ebp  
0036DE71 8B EC                mov         ebp,esp  
0036DE73 81 EC C0 00 00 00    sub         esp,0C0h  
0036DE79 53                   push        ebx  
0036DE7A 56                   push        esi  
0036DE7B 57                   push        edi  
0036DE7C 8D BD 40 FF FF FF    lea         edi,[ebp-0C0h]  
0036DE82 B9 30 00 00 00       mov         ecx,30h  
0036DE87 B8 CC CC CC CC       mov         eax,0CCCCCCCCh  
0036DE8C F3 AB                rep stos    dword ptr es:[edi]  
0036DE8E B9 5C 90 3A 00       mov         ecx,offset _FE3CD439_test@c (03A905Ch)  
0036DE93 E8 01 E4 FC FF       call        @__CheckForDebuggerJustMyCode@4 (033C299h)  

    printf("Hello world\n");
0036DE98 68 DC 87 37 00       push        offset string "Hello world\n" (03787DCh)  
0036DE9D E8 B3 E1 FC FF       call        _printf (033C055h)  
0036DEA2 83 C4 04             add         esp,4  

    return orig_MessageBoxA(hWnd, lpText, lpCaption, uType);
0036DEA5 8B F4                mov         esi,esp  

    return orig_MessageBoxA(hWnd, lpText, lpCaption, uType);
0036DEA7 8B 45 14             mov         eax,dword ptr [uType]  
0036DEAA 50                   push        eax  
0036DEAB 8B 4D 10             mov         ecx,dword ptr [lpCaption]  
0036DEAE 51                   push        ecx  
0036DEAF 8B 55 0C             mov         edx,dword ptr [lpText]  
0036DEB2 52                   push        edx  
0036DEB3 8B 45 08             mov         eax,dword ptr [hWnd]  
0036DEB6 50                   push        eax  
0036DEB7 FF 15 EC 71 3A 00    call        dword ptr [_orig_MessageBoxA (03A71ECh)]  
0036DEBD 3B F4                cmp         esi,esp  
0036DEBF E8 EE E3 FC FF       call        __RTC_CheckEsp (033C2B2h)  
0036DED4 8B E5                mov         esp,ebp  
0036DED6 5D                   pop         ebp  
0036DED7 C2 10 00             ret         10h  
}
