//
//  cl /LD VBA-Internals.c    user32.lib OleAut32.lib /FeVBA-Internals.dll /link /def:VBA-Internals.def
//
//  gcc -c VBA-Internals.c
//  gcc -shared VBA-Internals.o -lOleAut32 -o VBA-Internals.dll -Wl,--add-stdcall-alias
//

#include <windows.h>
#include <stdlib.h>

#define NOF_BREAKPOINTS 2


LONG WINAPI VectoredHandler(PEXCEPTION_POINTERS exPtr);

#define msg(txt)
#define msg_2(txt) MessageBox(0, txt, 0, 0);

// #include "c:\about\libc\search\t\search.c"
#include <search.h>


// --------------------------------------------------------------------

typedef char   instruction;
typedef LPVOID address;



instruction replaceInstruction(address addr, instruction instr) {

    char txt[400];

    instruction oldInstr;
    DWORD       oldProtection;

//  wsprintf(txt, "replace instruction at %d with %x", addr, instr);
//  msg_2(txt);


    VirtualProtect(addr, 1, PAGE_EXECUTE_READWRITE, &oldProtection);

    oldInstr =  *((instruction*) addr);
 *((instruction*) addr) = instr;


//  wsprintf(txt, "replace instruction %x with %x at %d", oldInstr, instr, addr);
//  msg_2("replaced");

    return oldInstr;
}

typedef struct {
    address       addr;
    char         *name;
    instruction   origInstruction;
}
breakpoint;


// int compareBreakpoints(const breakpoint *const f1, const breakpoint *const f2) 
   int compareBreakpoints(const void       *const f1, const void       *const f2)
{
  char txt[400];
//wsprintf(txt, "Comparing f1 (%d, addr: %d, %s) with f2 (%d, addr: %d, %s)", f1, ((breakpoint*)f1)->addr, ((breakpoint*)f1)->name, f2, ((breakpoint*)f2)->addr, ((breakpoint*)f2)->name);
// MM  wsprintf(txt, "Comparing f1 (%d, addr: %d    ) with f2 (%d, addr: %d    )", f1, ((breakpoint*)f1)->addr,                          f2, ((breakpoint*)f2)->addr                         );
// MM  msg(txt);
  return ((long)( ((breakpoint*)f1)->addr)) - ((long)(((breakpoint*)f2)->addr));
}

// breakpoint *breakpointTreeRoot = 0;
   void       *breakpointTreeRoot = 0;

__declspec(dllexport) void __stdcall addBreakpoint(address addr, char* name) {

    
    char txt[400]; 
    breakpoint *f;

// MM    wsprintf(txt, "addBreakpoint %s at %d", name, addr);
// MM    msg(txt);


    f = malloc(sizeof(breakpoint));
// MM    wsprintf(txt, "malloc, f= %d, sizeof(breakpoint) = %d",f, sizeof(breakpoint));
// MM    msg(txt);

    f->addr            = addr;
    f->name            = strdup(name);
//  msg("calling replaceInstruction");

// MM    wsprintf(txt, "addBreakpoint, after f->name = strdup(name): %s, addr(name) = %d, f = %d", f->name, f->name, f);
// MM    msg(txt);

    f->origInstruction = replaceInstruction(addr, 0xcc); // 0xcc = INT3
//  msg("returned from replaceInstruction");

// MM    wsprintf(txt, "addBreakpoint f->name = %s, addr f->name= %d, addr = %d, f = %d, f->origInstruction: %x", f->name, f->name, f->addr, f, f->origInstruction);
// MM    msg_2(txt);
    tsearch(f, &breakpointTreeRoot, compareBreakpoints);
//  tsearch( (void*) addr, &breakpointTreeRoot, compareBreakpoints);
//  msg("leaving addBreakpoint");
}

__declspec(dllexport) void __stdcall addDllFunctionBreakpoint(char* module, char* funcName) {

     address addr;
     HMODULE mod;
     char    breakpointName[200];
     char    txt           [400];

// MM     wsprintf(txt, "addDllFunctionBreakpoint %s / %s", module, funcName);
// MM     msg(txt);

     mod = GetModuleHandle(module);
     if (! mod) {
       MessageBox(0, module, "GetModuleHandle", 0);
       return;
     }

     addr = GetProcAddress(mod, funcName);
     if (!addr) {
          MessageBox(0, funcName, "GetProcAddress", 0);
          return;
     }

// MM     wsprintf(breakpointName, "%s.%s", module, funcName);
// MM     msg(breakpointName);

     addBreakpoint(addr, breakpointName);
//   msg("leaving addDllFunctionBreakpoint");

}


// address      func_addrs[NOF_BREAKPOINTS];
// instruction  old_instr [NOF_BREAKPOINTS];

// QQ int   nrLastFuncBreakpointHit;
address      addrLastBreakpointHit;

typedef void (WINAPI *addrCallBack_t)(BSTR);
addrCallBack_t addrCallBack;




void callCallback(char* txt) {
    int wslen;
    BSTR bstr;

    wslen = MultiByteToWideChar(CP_ACP, 0, txt, strlen(txt),    0,     0); bstr  = SysAllocStringLen(0, wslen);
            MultiByteToWideChar(CP_ACP, 0, txt, strlen(txt), bstr, wslen);

   (*addrCallBack)(bstr);

    SysFreeString(bstr);

}

LONG WINAPI VectoredHandler(PEXCEPTION_POINTERS exPtr) {


    int i;
    char txt[400];

    if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_BREAKPOINT) {

//    wsprintf(txt,"EXCEPTION_BREAKPOINT @ address %x", exPtr->ContextRecord->Eip);
//    callCallback(txt);
//
//

       breakpoint **breakpointFound;
       breakpoint  *breakpointHit;
       address addressHitByBreakpoint = (address) exPtr->ContextRecord->Eip;
//     addressHitByBreakpoint = (address) (( (DWORD) addressHitByBreakpoint) - 1);

//MM     wsprintf(txt, "addressHitByBreakpoint = %d", addressHitByBreakpoint);
//MM     msg(txt);


       breakpoint findMe;
       findMe.addr = addressHitByBreakpoint;
//     breakpoint* p;
//     p = &findMe;


//     breakpointHit = tfind(&findMe, &breakpointTreeRoot, compareBreakpoints);
       breakpointFound = tfind(&findMe, &breakpointTreeRoot, compareBreakpoints);


       if (!breakpointFound) {
         MessageBox(0, "Did not tfind breakpointHit...", 0, 0);
       }
       else {

          breakpointHit = *breakpointFound;

//        wsprintf(txt, "&findMe = %d, breakpointHit = %d", &findMe, breakpointHit);
//        msg(txt);
//        wsprintf(txt, "breakpointHit = %d", breakpointHit);

//        wsprintf(txt, "Hit Breakpoint, breakpointHit = %d, addr = %d (name = %s, addr(name) = %d), origInstr: %x", breakpointHit, breakpointHit->addr, breakpointHit->name, breakpointHit->name, breakpointHit->origInstruction);
          wsprintf(txt, "Hit Breakpoint %s", breakpointHit->name);
//        msg_2(txt);
          callCallback(txt);
   
          addrLastBreakpointHit = addressHitByBreakpoint;
   
          replaceInstruction(addressHitByBreakpoint, breakpointHit->origInstruction);

       // Single Step:
          exPtr->ContextRecord->EFlags |= 0x100;

          SetThreadContext(GetCurrentThread(), exPtr->ContextRecord);
       }


// QQ       for (i=0; i<NOF_BREAKPOINTS; i++) {
// QQ   
// QQ         // The EIP register contains the address of the break point
// QQ         // instruction that triggered the exception.
// QQ         // This allows to determine the number of the breakpoint:
// QQ           if (exPtr->ContextRecord->Eip == (int) func_addrs[i]) {
// QQ 
// QQ               wsprintf(txt, "Breakpoint %d was hit", i);
// QQ               callCallback(txt);
// QQ 
// QQ            //
// QQ            // Store the number of the breakpoint so that we can
// QQ            // set the breakpoint again after single stepping
// QQ            // the »current« instruction:
// QQ            //
// QQ               nrLastFuncBreakpointHit = i;
// QQ   
// QQ            //
// QQ            // In order to proceed with the execution of the program, we
// QQ            // restore the old value of the byte that was replaced by
// QQ            // the int-3 instruction:
// QQ            //
// QQ               replaceInstruction(func_addrs[i], old_instr[i]);
// QQ   
// QQ            //
// QQ            // Set TF bit in order to single step the next
// QQ            // instruction. This allows to set the break point
// QQ            // again after the single step instruction.
// QQ            //
// QQ   
// QQ               exPtr->ContextRecord->EFlags |= 0x100;
// QQ   
// QQ            //
// QQ            // Resume execution:
// QQ            //
// QQ               SetThreadContext(GetCurrentThread(), exPtr->ContextRecord);
// QQ           }
// QQ        }

       callCallback("EXCEPTION_BREAKPOINT: no matching address found");
       return EXCEPTION_CONTINUE_SEARCH;
    }
    else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_SINGLE_STEP) {
    // callCallback("EXCEPTION_SINGLE_STEP");

    //
    // The processor is one instruction behind the last breakpoint that was
    // hit. We can now set the breakpoint again.


//     TODO: Uncomment me:
//     *func_addrs[nrLastFuncBreakpointHit] = 0xcc;
// QQ  replaceInstruction(func_addrs[nrLastFuncBreakpointHit], 0xcc);

//     msg_2("Single Step: Replacing instruction in single step");
       replaceInstruction(addrLastBreakpointHit, 0xcc); // INT3              
//     msg_2("Single Step: instruction was replaced, going to call SetThreadContext");

    //
    // Apparently, the TF flag is reset after the single execution, it
    // needs not be reset.
    //
    // Thus, we can resume execution again:
    //

    // callCallback("EXCEPTION_SINGLE_STEP -> SetThreadContext");
     //exPtr->ContextRecord->EFlags |= 0x100;
       SetThreadContext(GetCurrentThread(), exPtr->ContextRecord);

    }
    else {

    //
    // Should never be reached.
    //
    
           if      (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_ACCESS_VIOLATION           ) { wsprintf(txt, "Unexpected exception EXCEPTION_ACCESS_VIOLATION         , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_DATATYPE_MISALIGNMENT      ) { wsprintf(txt, "Unexpected exception EXCEPTION_DATATYPE_MISALIGNMENT    , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_BREAKPOINT                 ) { wsprintf(txt, "Unexpected exception EXCEPTION_BREAKPOINT               , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_SINGLE_STEP                ) { wsprintf(txt, "Unexpected exception EXCEPTION_SINGLE_STEP              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_ARRAY_BOUNDS_EXCEEDED      ) { wsprintf(txt, "Unexpected exception EXCEPTION_ARRAY_BOUNDS_EXCEEDED    , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_DENORMAL_OPERAND       ) { wsprintf(txt, "Unexpected exception EXCEPTION_FLT_DENORMAL_OPERAND     , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_DIVIDE_BY_ZERO         ) { wsprintf(txt, "Unexpected exception EXCEPTION_FLT_DIVIDE_BY_ZERO       , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_INEXACT_RESULT         ) { wsprintf(txt, "Unexpected exception EXCEPTION_FLT_INEXACT_RESULT       , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_INVALID_OPERATION      ) { wsprintf(txt, "Unexpected exception EXCEPTION_FLT_INVALID_OPERATION    , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_OVERFLOW               ) { wsprintf(txt, "Unexpected exception EXCEPTION_FLT_OVERFLOW             , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_STACK_CHECK            ) { wsprintf(txt, "Unexpected exception EXCEPTION_FLT_STACK_CHECK          , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_UNDERFLOW              ) { wsprintf(txt, "Unexpected exception EXCEPTION_FLT_UNDERFLOW            , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_INT_DIVIDE_BY_ZERO         ) { wsprintf(txt, "Unexpected exception EXCEPTION_INT_DIVIDE_BY_ZERO       , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_INT_OVERFLOW               ) { wsprintf(txt, "Unexpected exception EXCEPTION_INT_OVERFLOW             , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_PRIV_INSTRUCTION           ) { wsprintf(txt, "Unexpected exception EXCEPTION_PRIV_INSTRUCTION         , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_IN_PAGE_ERROR              ) { wsprintf(txt, "Unexpected exception EXCEPTION_IN_PAGE_ERROR            , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_ILLEGAL_INSTRUCTION        ) { wsprintf(txt, "Unexpected exception EXCEPTION_ILLEGAL_INSTRUCTION      , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_NONCONTINUABLE_EXCEPTION   ) { wsprintf(txt, "Unexpected exception EXCEPTION_NONCONTINUABLE_EXCEPTION , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_STACK_OVERFLOW             ) { wsprintf(txt, "Unexpected exception EXCEPTION_STACK_OVERFLOW           , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_INVALID_DISPOSITION        ) { wsprintf(txt, "Unexpected exception EXCEPTION_INVALID_DISPOSITION      , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_GUARD_PAGE                 ) { wsprintf(txt, "Unexpected exception EXCEPTION_GUARD_PAGE               , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_INVALID_HANDLE             ) { wsprintf(txt, "Unexpected exception EXCEPTION_INVALID_HANDLE           , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
//         else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_POSSIBLE_DEADLOCK          ) MessageBox(0, "Unexpected exception EXCEPTION_POSSIBLE_DEADLOCK        ",  0, 0);
           else {

//             wsprintf(txt, "Unexpected exception with code %8x", exPtr->ExceptionRecord->ExceptionCode);
//             callCallback(txt);

           }

 //        SetThreadContext(GetCurrentThread(), exPtr->ContextRecord);
    
    }


    return EXCEPTION_CONTINUE_SEARCH;
}

__declspec(dllexport) void __stdcall VBAInternalsInit(addrCallBack_t addrCallBack_) {

    int i, res;

    char txt[400];

//  wsprintf(txt, "VBAInternalsInit, sizeof(address) = %d", sizeof(address));
//  msg(txt);


    addrCallBack = addrCallBack_;

    AddVectoredExceptionHandler(1, VectoredHandler);

//  msg("VBAInternalsInit leaving");

 // func_addrs[0] = (char*) GetProcAddress(GetModuleHandle("VBE7.dll"), "rtcMsgBox");
 // func_addrs[1] = (char*) func_1;
 // func_addrs[2] = (char*) func_2;
 // func_addrs[3] = (char*) func_3;

// QQ    for (i=0; i<NOF_BREAKPOINTS; i++) {
// QQ     //
// QQ     // Setting the breakpoints at the functions addresses.
// QQ     //
// QQ     // First, we need to make the code segment writable to be able to
// QQ     // insert the breakpoint instruction. Otherwise, the modification of
// QQ     // the (read-only) code segment would cause an exception.
// QQ     //
// QQ
// QQ        DWORD oldProtection;
// QQ        VirtualProtect(func_addrs[i], 1, PAGE_EXECUTE_READWRITE, &oldProtection);
// QQ
// QQ     //
// QQ     // Inject the int-3 instruction (0xcc):
// QQ     // 
// QQ     // We also want to store the value of the byte before we set the int-3
// QQ     // instruction:
// QQ     //
// QQ        old_instr[i] = replaceInstruction(func_addrs[i], 0xcc);
// QQ
// QQ    }
}

BOOL WINAPI DllMain(
  _In_ HINSTANCE hinstDLL,
  _In_ DWORD     fdwReason,
  _In_ LPVOID    lpvReserved
) {

   int i;

// if      (fdwReason == DLL_PROCESS_ATTACH) {    MessageBox(0, "DllMain DLL_PROCESS_ATTACH", 0, 0)    ;}
// else if (fdwReason == DLL_PROCESS_DETACH) {    MessageBox(0, "DllMain DLL_PROCESS_DETACH", 0, 0)    ;}
// else if (fdwReason == DLL_THREAD_ATTACH ) { /* MessageBox(0, "DllMain DLL_THREAD_ATTACH ", 0, 0) */ ;}
// else if (fdwReason == DLL_THREAD_DETACH ) { /* MessageBox(0, "DllMain DLL_THREAD_DETACH ", 0, 0) */ ;}
// else                                      {    MessageBox(0, "DllMain hmmmm???"          , 0, 0)    ;}

// if (fdwReason == DLL_PROCESS_ATTACH) {
//    func_addrs[0] = (address) GetProcAddress(GetModuleHandle("VBE7.dll"), "rtcMsgBox");
//    func_addrs[1] = (address) GetProcAddress(GetModuleHandle("VBE7.dll"), "rtcRound" );
// }

// if (fdwReason == DLL_PROCESS_DETACH) {
//    for (i=0; i<NOF_BREAKPOINTS; i++) {
//      replaceInstruction(func_addrs[i], old_instr[i]);
//    }
// }

   return TRUE;

}
