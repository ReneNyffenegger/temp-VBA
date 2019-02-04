//
//  cl /LD VBA-Internals.c    user32.lib OleAut32.lib /FeVBA-Internals.dll /link /def:VBA-Internals.def
//

#include <windows.h>
// #include <search.h>
#include <stdlib.h>

#define NOF_BREAKPOINTS 2


LONG WINAPI VectoredHandler(PEXCEPTION_POINTERS exPtr);

// Found at http://web.mit.edu/freebsd/head/crypto/heimdal/lib/roken/tsearch.c
#include "tsearch.c"


// --------------------------------------------------------------------

typedef char   instruction;
typedef LPVOID address;


#define msg(txt)
#define msg_2(txt) MessageBox(0, txt, 0, 0);

instruction replaceInstruction(address addr, instruction instr) {

    char txt[200];

    instruction oldInstr;
    DWORD       oldProtection;

    wsprintf(txt, "replace instruction at %d, going to get oldInstr", addr);
    msg(txt);


    VirtualProtect(addr, 1, PAGE_EXECUTE_READWRITE, &oldProtection);

    oldInstr =  *((instruction*) addr);
 *((instruction*) addr) = instr;

    return oldInstr;
}

typedef struct {
    address       addr;
    char         *name;
    instruction   origInstruction;
}
breakpoint;


int compareBreakpoints(const breakpoint *const f1, const breakpoint *const f2) {
  char txt[200];
  wsprintf(txt, "Comparing f1 (%d, addr: %d, %s) with f2 (%d, addr: %d, %s)", f1, (long) f1->addr, f1->name, f2, (long) f2->addr, f2->name);
  msg_2(txt);
  return ((long)(f1->addr)) - ((long)(f2->addr));
}

// was: breakpoint *breakpointTreeRoot = 0 ...
void *breakpointTreeRoot = 0;

__declspec(dllexport) void __stdcall addBreakpoint(address addr, char* name) {

    
    char txt[200]; 
    breakpoint *f;

    wsprintf(txt, "addBreakpoint %s at %d", name, addr);
    msg(txt);


    f = malloc(sizeof(breakpoint));

    f->addr            = addr;
    f->name            = strdup(name);
    msg("calling replaceInstruction");

    wsprintf(txt, "after f->name = strdup(name): %s", f->name);
    msg(txt);

    f->origInstruction = replaceInstruction(addr, 0xcc); // 0xcc = INT3
    msg("returned from replaceInstruction");

    wsprintf(txt, "addBreakpoint f->name = %s, addr = %d, f = %d", f->name, f->addr, f);
    msg_2(txt);
    tsearch(f, &breakpointTreeRoot, compareBreakpoints);
    msg("leaving addBreakpoint");
}

__declspec(dllexport) void __stdcall addDllFunctionBreakpoint(char* module, char* funcName) {

     address addr;
     HMODULE mod;
     char    breakpointName[200];
     char    txt           [200];

     wsprintf(txt, "addDllFunctionBreakpoint %s / %s", module, funcName);

     msg(txt);

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

     wsprintf(breakpointName, "%s.%s", module, funcName);
     msg(breakpointName);

     addBreakpoint(addr, breakpointName);
     msg("leaving addDllFunctionBreakpoint");

}


address      func_addrs[NOF_BREAKPOINTS];
instruction  old_instr [NOF_BREAKPOINTS];

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
    char txt[200];

    if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_BREAKPOINT) {

//    wsprintf(txt,"EXCEPTION_BREAKPOINT @ address %x", exPtr->ContextRecord->Eip);
//    callCallback(txt);
//
//

       breakpoint *breakpointHit;
       address addressHitByBreakpoint = (address) exPtr->ContextRecord->Eip;
//     addressHitByBreakpoint = (address) (( (DWORD) addressHitByBreakpoint) - 1);

       wsprintf(txt, "addressHitByBreakpoint = %d", addressHitByBreakpoint);
       msg_2(txt);

       breakpointHit = tfind(&addressHitByBreakpoint, &breakpointTreeRoot, compareBreakpoints);

       if (!breakpointHit) {
         MessageBox(0, "Did not tfind breakpointHit...", 0, 0);
       }
       else {

          wsprintf(txt, "breakpointHit = %d", breakpointHit);
          msg_2(txt);

//        wsprintf(txt, "Hit Breakpoint, breakpointHit = %d (name = %s)", breakpointHit->addr, breakpointHit->name);
          callCallback(txt);
   
          addrLastBreakpointHit = addressHitByBreakpoint;
   
          replaceInstruction(addressHitByBreakpoint, breakpointHit->origInstruction);
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

       msg_2("Replacing instruction in single step");
       replaceInstruction(addrLastBreakpointHit, 0xcc); // INT3              
       msg_2("instruction was replaced, going to call SetThreadContext");

    //
    // Apparently, the TF flag is reset after the single execution, it
    // needs not be reset.
    //
    // Thus, we can resume execution again:
    //

    // callCallback("EXCEPTION_SINGLE_STEP -> SetThreadContext");
       SetThreadContext(GetCurrentThread(), exPtr->ContextRecord);

    }
    else {

    //
    // Should never be reached.
    //
    
           if      (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_ACCESS_VIOLATION           ) MessageBox(0, "Unexpected exception EXCEPTION_ACCESS_VIOLATION         ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_DATATYPE_MISALIGNMENT      ) MessageBox(0, "Unexpected exception EXCEPTION_DATATYPE_MISALIGNMENT    ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_BREAKPOINT                 ) MessageBox(0, "Unexpected exception EXCEPTION_BREAKPOINT               ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_SINGLE_STEP                ) MessageBox(0, "Unexpected exception EXCEPTION_SINGLE_STEP              ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_ARRAY_BOUNDS_EXCEEDED      ) MessageBox(0, "Unexpected exception EXCEPTION_ARRAY_BOUNDS_EXCEEDED    ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_DENORMAL_OPERAND       ) MessageBox(0, "Unexpected exception EXCEPTION_FLT_DENORMAL_OPERAND     ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_DIVIDE_BY_ZERO         ) MessageBox(0, "Unexpected exception EXCEPTION_FLT_DIVIDE_BY_ZERO       ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_INEXACT_RESULT         ) MessageBox(0, "Unexpected exception EXCEPTION_FLT_INEXACT_RESULT       ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_INVALID_OPERATION      ) MessageBox(0, "Unexpected exception EXCEPTION_FLT_INVALID_OPERATION    ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_OVERFLOW               ) MessageBox(0, "Unexpected exception EXCEPTION_FLT_OVERFLOW             ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_STACK_CHECK            ) MessageBox(0, "Unexpected exception EXCEPTION_FLT_STACK_CHECK          ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_UNDERFLOW              ) MessageBox(0, "Unexpected exception EXCEPTION_FLT_UNDERFLOW            ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_INT_DIVIDE_BY_ZERO         ) MessageBox(0, "Unexpected exception EXCEPTION_INT_DIVIDE_BY_ZERO       ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_INT_OVERFLOW               ) MessageBox(0, "Unexpected exception EXCEPTION_INT_OVERFLOW             ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_PRIV_INSTRUCTION           ) MessageBox(0, "Unexpected exception EXCEPTION_PRIV_INSTRUCTION         ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_IN_PAGE_ERROR              ) MessageBox(0, "Unexpected exception EXCEPTION_IN_PAGE_ERROR            ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_ILLEGAL_INSTRUCTION        ) MessageBox(0, "Unexpected exception EXCEPTION_ILLEGAL_INSTRUCTION      ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_NONCONTINUABLE_EXCEPTION   ) MessageBox(0, "Unexpected exception EXCEPTION_NONCONTINUABLE_EXCEPTION ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_STACK_OVERFLOW             ) MessageBox(0, "Unexpected exception EXCEPTION_STACK_OVERFLOW           ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_INVALID_DISPOSITION        ) MessageBox(0, "Unexpected exception EXCEPTION_INVALID_DISPOSITION      ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_GUARD_PAGE                 ) MessageBox(0, "Unexpected exception EXCEPTION_GUARD_PAGE               ",  0, 0);
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_INVALID_HANDLE             ) MessageBox(0, "Unexpected exception EXCEPTION_INVALID_HANDLE           ",  0, 0);
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

    addrCallBack = addrCallBack_;

    AddVectoredExceptionHandler(1, VectoredHandler);

    msg("VBAInternalsInit leaving");

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

   if (fdwReason == DLL_PROCESS_ATTACH) {
      func_addrs[0] = (address) GetProcAddress(GetModuleHandle("VBE7.dll"), "rtcMsgBox");
      func_addrs[1] = (address) GetProcAddress(GetModuleHandle("VBE7.dll"), "rtcRound" );
   }

   if (fdwReason == DLL_PROCESS_DETACH) {
      for (i=0; i<NOF_BREAKPOINTS; i++) {
        replaceInstruction(func_addrs[i], old_instr[i]);
      }
   }

   return TRUE;

}
