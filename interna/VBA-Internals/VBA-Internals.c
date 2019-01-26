//
//  cl /LD VBA-Internals.c    user32.lib OleAut32.lib /FeVBA-Internals.dll /link /def:VBA-Internals.def
//
//  gcc -c VBA-Internals.c
//  gcc -shared VBA-Internals.o -lOleAut32 -o VBA-Internals.dll -Wl,--add-stdcall-alias
//

#include <windows.h>
#include <stdlib.h>

#define USE_SEARCH

#ifdef USE_SEARCH
#else
#define NOF_BREAKPOINTS 2
#endif


LONG WINAPI VectoredHandler(PEXCEPTION_POINTERS exPtr);

#define msg(txt)
#define msg_2(txt) MessageBox(0, txt, 0, 0);

// #include "c:\about\libc\search\t\search.c"
#include <search.h>

#define TQ84_DEBUG_ENABLED
#define TQ84_DEBUG_TO_FILE
#include "c:\lib\tq84-c-debug\tq84_debug.c"


// --------------------------------------------------------------------

typedef unsigned char   instruction;
typedef LPVOID address;



instruction replaceInstruction(address addr, instruction instr) {
    TQ84_DEBUG_INDENT_T("replaceInstruction, addr=%d, byte = %u", addr, instr);

    char txt[400];

    instruction oldInstr;
    DWORD       oldProtection;

//  wsprintf(txt, "replace instruction at %d with %x", addr, instr);
//  msg_2(txt);


    TQ84_DEBUG("VirtualProtect");
    VirtualProtect(addr, 1, PAGE_EXECUTE_READWRITE, &oldProtection);

    oldInstr =  *((instruction*) addr);
    TQ84_DEBUG("old byte = %u, now going to changing to %u", oldInstr, instr);
 *((instruction*) addr) = instr;

    TQ84_DEBUG("Returning old byte: %u", oldInstr); 

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

  TQ84_DEBUG_INDENT_T("compareBreakpoints, addr 1 = %d, addr 2 = %d", ((breakpoint*)f1)->addr, ((breakpoint*)f2)->addr);

  char txt[400];
//wsprintf(txt, "Comparing f1 (%d, addr: %d, %s) with f2 (%d, addr: %d, %s)", f1, ((breakpoint*)f1)->addr, ((breakpoint*)f1)->name, f2, ((breakpoint*)f2)->addr, ((breakpoint*)f2)->name);
// MM  wsprintf(txt, "Comparing f1 (%d, addr: %d    ) with f2 (%d, addr: %d    )", f1, ((breakpoint*)f1)->addr,                          f2, ((breakpoint*)f2)->addr                         );
// MM  msg(txt);
  return ((long)( ((breakpoint*)f1)->addr)) - ((long)(((breakpoint*)f2)->addr));
}

// breakpoint *breakpointTreeRoot = 0;
   void       *breakpointTreeRoot = 0;

__declspec(dllexport) void __stdcall addBreakpoint(address addr, char* name) {

    TQ84_DEBUG_INDENT_T("addBreakpoint, addr = %d, name = %s", addr, name);
    
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

    TQ84_DEBUG("going to call replaceInstruction");
    f->origInstruction = replaceInstruction(addr, 0xcc); // 0xcc = INT3
//  msg("returned from replaceInstruction");

// MM    wsprintf(txt, "addBreakpoint f->name = %s, addr f->name= %d, addr = %d, f = %d, f->origInstruction: %x", f->name, f->name, f->addr, f, f->origInstruction);
// MM    msg_2(txt);
    TQ84_DEBUG("going to call tsearch");
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

     wsprintf(breakpointName, "%s.%s", module, funcName);
// MM     msg(breakpointName);

     addBreakpoint(addr, breakpointName);
//   msg("leaving addDllFunctionBreakpoint");

}


#ifdef USE_SEARCH
address      addrLastBreakpointHit;
#else
address      func_addrs[NOF_BREAKPOINTS];
instruction  old_instr [NOF_BREAKPOINTS];
int   nrLastFuncBreakpointHit;
#endif




// typedef void (WINAPI   *addrCallBack_t)(BSTR);
   typedef void (CALLBACK *addrCallBack_t)(BSTR);
addrCallBack_t addrCallBack;




void callCallback(char* txt) {
    TQ84_DEBUG("callCallback, txt = %s", txt);
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
       TQ84_DEBUG("EXCEPTION_BREAKPOINT");
       TQ84_DEBUG("EXCEPTION_BREAKPOINT, addr = %d", exPtr->ContextRecord->Eip);

#ifdef USE_SEARCH


//    callCallback(txt);
//
//

       breakpoint **breakpointFound;
       breakpoint  *breakpointHit;
       address addressHitByBreakpoint = (address) exPtr->ContextRecord->Eip;
       TQ84_DEBUG("addressHitByBreakpoint = %d", exPtr->ContextRecord->Eip);
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
         TQ84_DEBUG("Did not tfind breakpointHit...");
         MessageBox(0, "Did not tfind breakpointHit...", 0, 0);
       }
       else {

          breakpointHit = *breakpointFound;

//        wsprintf(txt, "&findMe = %d, breakpointHit = %d", &findMe, breakpointHit);
//        msg(txt);
//        wsprintf(txt, "breakpointHit = %d", breakpointHit);

          wsprintf(txt, "Hit Breakpoint, breakpointHit = %d, addr = %d (name = %s, addr(name) = %d), origInstr: %x", breakpointHit, breakpointHit->addr, breakpointHit->name, breakpointHit->name, breakpointHit->origInstruction);
          TQ84_DEBUG(txt);
//        wsprintf(txt, "Hit Breakpoint %s", breakpointHit->name);
//        msg_2(txt);
          callCallback(txt);
   
          addrLastBreakpointHit = addressHitByBreakpoint;
   
          replaceInstruction(addressHitByBreakpoint, breakpointHit->origInstruction);

       // Single Step:
//        exPtr->ContextRecord->EFlags |= 0x100;

          TQ84_DEBUG("Going to call SetThreadContext");

          SetThreadContext(GetCurrentThread(), exPtr->ContextRecord);
       }

#else


       for (i=0; i<NOF_BREAKPOINTS; i++) {
   
         // The EIP register contains the address of the break point
         // instruction that triggered the exception.
         // This allows to determine the number of the breakpoint:
           if (exPtr->ContextRecord->Eip == (int) func_addrs[i]) {
 
               wsprintf(txt, "Breakpoint %d was hit", i);
               callCallback(txt);
 
            //
            // Store the number of the breakpoint so that we can
            // set the breakpoint again after single stepping
            // the »current« instruction:
            //
               nrLastFuncBreakpointHit = i;
   
            //
            // In order to proceed with the execution of the program, we
            // restore the old value of the byte that was replaced by
            // the int-3 instruction:
            //
               replaceInstruction(func_addrs[i], old_instr[i]);
   
            //
            // Set TF bit in order to single step the next
            // instruction. This allows to set the break point
            // again after the single step instruction.
            //
   
//             exPtr->ContextRecord->EFlags |= 0x100;
   
            //
            // Resume execution:
            //
               SetThreadContext(GetCurrentThread(), exPtr->ContextRecord);
           }
        }
#endif

       TQ84_DEBUG  ("EXCEPTION_BREAKPOINT: no matching address found");
       callCallback("EXCEPTION_BREAKPOINT: no matching address found");
       return EXCEPTION_CONTINUE_SEARCH;

    }
    else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_SINGLE_STEP) {
       TQ84_DEBUG("EXCEPTION_SINGLE_STEP");
       TQ84_DEBUG("EXCEPTION_SINGLE_STEP: Address: %d", exPtr->ContextRecord->Eip);


#ifdef USE_SEARCH

       TQ84_DEBUG("EXCEPTION_SINGLE_STEP: going to replaceInstruction at lastAddrBreakpoint hit: %d", addrLastBreakpointHit);
       replaceInstruction(addrLastBreakpointHit, 0xcc); // INT3              

#else

// QQ      replaceInstruction(func_addrs[nrLastFuncBreakpointHit], 0xcc);


// OLD     // callCallback("EXCEPTION_SINGLE_STEP");
// OLD 
// OLD     //
// OLD     // The processor is one instruction behind the last breakpoint that was
// OLD     // hit. We can now set the breakpoint again.
// OLD 
// OLD 
// OLD //     TODO: Uncomment me:
// OLD //     *func_addrs[nrLastFuncBreakpointHit] = 0xcc;
// OLD // QQ  replaceInstruction(func_addrs[nrLastFuncBreakpointHit], 0xcc);
// OLD 
// OLD //     msg_2("Single Step: Replacing instruction in single step");
// OLD //     msg_2("Single Step: instruction was replaced, going to call SetThreadContext");
// OLD 
// OLD     //
// OLD     // Apparently, the TF flag is reset after the single execution, it
// OLD     // needs not be reset.
// OLD     //
// OLD     // Thus, we can resume execution again:
// OLD     //
// OLD 
// OLD     // callCallback("EXCEPTION_SINGLE_STEP -> SetThreadContext");

#endif

//     exPtr->ContextRecord->EFlags |= 0x100;

       TQ84_DEBUG("Going to call SetThreadContext");
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

               TQ84_DEBUG("Unexpected exception with code %d at address %d", exPtr->ExceptionRecord->ExceptionCode, exPtr->ContextRecord->Eip);
//             wsprintf(txt, "Unexpected exception with code %d at address %d", exPtr->ExceptionRecord->ExceptionCode, exPtr->ContextRecord->Eip);
//             TQ84_DEBUG(txt);
//             msg_2(txt);
//             callCallback(txt);

           }

     //    SetThreadContext(GetCurrentThread(), exPtr->ContextRecord);
           return EXCEPTION_CONTINUE_SEARCH;
    }


    return EXCEPTION_CONTINUE_SEARCH;
}

__declspec(dllexport) void __stdcall VBAInternalsInit(addrCallBack_t addrCallBack_) {

    TQ84_DEBUG_INDENT_T("VBAInternalsInit");


    int i, res;

    char txt[400];

//  wsprintf(txt, "VBAInternalsInit, sizeof(address) = %d", sizeof(address));
//  msg(txt);


    addrCallBack = addrCallBack_;
    callCallback("addrCallBack assigned");

    AddVectoredExceptionHandler(1, VectoredHandler);

//  msg("VBAInternalsInit leaving");

//  func_addrs[0] = (char*) GetProcAddress(GetModuleHandle("VBE7.dll"), "rtcMsgBox");
//  func_addrs[1] = (char*) GetProcAddress(GetModuleHandle("VBE7.dll"), "rtcRound" );

 // func_addrs[2] = (char*) func_2;
 // func_addrs[3] = (char*) func_3;

#ifdef USE_SEARCH
#else
    for (i=0; i<NOF_BREAKPOINTS; i++) {
       callCallback("Breakpoint");
       TQ84_DEBUG_INDENT_T("i = %d", i);
     //
     // Setting the breakpoints at the functions addresses.
     //
     // First, we need to make the code segment writable to be able to
     // insert the breakpoint instruction. Otherwise, the modification of
     // the (read-only) code segment would cause an exception.
     //

        DWORD oldProtection;
        VirtualProtect(func_addrs[i], 1, PAGE_EXECUTE_READWRITE, &oldProtection);

     //
     // Inject the int-3 instruction (0xcc):
     // 
     // We also want to store the value of the byte before we set the int-3
     // instruction:
     //
        old_instr[i] = replaceInstruction(func_addrs[i], 0xcc);

    }
#endif

}

BOOL WINAPI DllMain(
  _In_ HINSTANCE hinstDLL,
  _In_ DWORD     fdwReason,
  _In_ LPVOID    lpvReserved
) {

   int i;
   static int debug_is_open = 0;

   if (!debug_is_open) {
      tq84_debug_open("c:\\temp\\dbg.vba", "w");
      debug_is_open ++;
   }

   TQ84_DEBUG_INDENT_T("DllMain");

// if      (fdwReason == DLL_PROCESS_ATTACH) {    MessageBox(0, "DllMain DLL_PROCESS_ATTACH", 0, 0)    ;}
// else if (fdwReason == DLL_PROCESS_DETACH) {    MessageBox(0, "DllMain DLL_PROCESS_DETACH", 0, 0)    ;}
// else if (fdwReason == DLL_THREAD_ATTACH ) { /* MessageBox(0, "DllMain DLL_THREAD_ATTACH ", 0, 0) */ ;}
// else if (fdwReason == DLL_THREAD_DETACH ) { /* MessageBox(0, "DllMain DLL_THREAD_DETACH ", 0, 0) */ ;}
// else                                      {    MessageBox(0, "DllMain hmmmm???"          , 0, 0)    ;}

   if (fdwReason == DLL_PROCESS_ATTACH) {
      TQ84_DEBUG("fdwReason == DLL_PROCESS_ATTACH");
#ifdef USE_SEARCH
#else
      func_addrs[0] = (address) GetProcAddress(GetModuleHandle("VBE7.dll"), "rtcMsgBox");
      func_addrs[1] = (address) GetProcAddress(GetModuleHandle("VBE7.dll"), "rtcRound" );
#endif

   }

   if (fdwReason == DLL_PROCESS_DETACH) {
      TQ84_DEBUG("DLL_PROCESS_DETACH");
#ifdef USE_SEARCH
#else
      for (i=0; i<NOF_BREAKPOINTS; i++) {
        TQ84_DEBUG("going to call replaceInstruction");
        replaceInstruction(func_addrs[i], old_instr[i]);
      }
#endif
   }

   return TRUE;

}
