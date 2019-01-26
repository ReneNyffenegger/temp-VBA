//
//  cl /LD VBA-Internals.c    user32.lib OleAut32.lib /FeVBA-Internals.dll /link /def:VBA-Internals.def
//
//  gcc -c VBA-Internals.c
//  gcc -shared VBA-Internals.o -lOleAut32 -o VBA-Internals.dll -Wl,--add-stdcall-alias
//
//  TODO
//    Detouring: https://reverseengineering.stackexchange.com/a/13694

#include <windows.h>
#include <stdlib.h>

// #include "vbObject.h"

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


HANDLE curProcess;

instruction replaceInstruction(address addr, instruction instr) { // {
    TQ84_DEBUG_INDENT_T("replaceInstruction, addr=%d, byte = %u", addr, instr);

    char txt[400];
    SIZE_T nofBytes;

    instruction oldInstr;
    DWORD       oldProtection;

//  wsprintf(txt, "replace instruction at %d with %x", addr, instr);
//  msg_2(txt);


//  TQ84_DEBUG("VirtualProtect");
    VirtualProtect(addr, 1, PAGE_EXECUTE_READWRITE, &oldProtection);

//  oldInstr =  *((instruction*) addr);
    if (! ReadProcessMemory(curProcess, addr, &oldInstr, 1, &nofBytes)) {
      MessageBox(0, "Could not ReadProcessMemory", 0, 0);
    }
//  TQ84_DEBUG("nofBytes read: %d, byte = %u, now going to change to %u", nofBytes, oldInstr, instr);

// *((instruction*) addr) = instr;
    if (! WriteProcessMemory(curProcess, addr, &instr, 1, &nofBytes)) {
      MessageBox(0, "Could not ReadProcessMemory", 0, 0);
    }
//  TQ84_DEBUG("nofBytes written: %d", nofBytes);

//  VirtualProtect(addr, 1, oldProtection, &oldProtection);

//  if (! FlushInstructionCache(curProcess, addr, 10)) {
//    MessageBox(0, "Could not FlushInstructionCache", 0, 0);
//  }

//  TQ84_DEBUG("Returning old byte: %u", oldInstr); 

//  wsprintf(txt, "replace instruction %x with %x at %d", oldInstr, instr, addr);
//  msg_2("replaced");

    return oldInstr;
} // }

typedef struct { // {
    address       addr;
    char         *name;
    instruction   origInstruction;
    char         *format;
}
breakpoint; // }

int compareBreakpoints(const void *const f1, const void *const f2) { // {


//  TQ84_DEBUG_INDENT_T("compareBreakpoints, addr 1 = %d, addr 2 = %d", ((breakpoint*)f1)->addr, ((breakpoint*)f2)->addr);

  char txt[400];
//wsprintf(txt, "Comparing f1 (%d, addr: %d, %s) with f2 (%d, addr: %d, %s)", f1, ((breakpoint*)f1)->addr, ((breakpoint*)f1)->name, f2, ((breakpoint*)f2)->addr, ((breakpoint*)f2)->name);
// MM  wsprintf(txt, "Comparing f1 (%d, addr: %d    ) with f2 (%d, addr: %d    )", f1, ((breakpoint*)f1)->addr,                          f2, ((breakpoint*)f2)->addr                         );
// MM  msg(txt);
  return ((long)( ((breakpoint*)f1)->addr)) - ((long)(((breakpoint*)f2)->addr));
} // }

// breakpoint *breakpointTreeRoot = 0;
   void       *breakpointTreeRoot = 0;

__declspec(dllexport) void __stdcall addBreakpoint(address addr, char* name, char* format) { // {

    TQ84_DEBUG_INDENT_T("addBreakpoint, addr = %d, name = %s", addr, name);
    
    char txt[400]; 
    breakpoint *f;



// MM    wsprintf(txt, "addBreakpoint %s at %d", name, addr);
// MM    msg(txt);


    f = malloc(sizeof(breakpoint));
// MM    wsprintf(txt, "malloc, f= %d, sizeof(breakpoint) = %d",f, sizeof(breakpoint));
// MM    msg(txt);

    f->addr            = addr;
    f->name            = strdup(name  );
    f->format          = strdup(format);
//  msg("calling replaceInstruction");

// MM    wsprintf(txt, "addBreakpoint, after f->name = strdup(name): %s, addr(name) = %d, f = %d", f->name, f->name, f);
// MM    msg(txt);

//  TQ84_DEBUG("going to call replaceInstruction");
    f->origInstruction = replaceInstruction(addr, 0xcc); // 0xcc = INT3
//  msg("returned from replaceInstruction");

// MM    wsprintf(txt, "addBreakpoint f->name = %s, addr f->name= %d, addr = %d, f = %d, f->origInstruction: %x", f->name, f->name, f->addr, f, f->origInstruction);
// MM    msg_2(txt);
//  TQ84_DEBUG("going to call tsearch");
    tsearch(f, &breakpointTreeRoot, compareBreakpoints);
//  tsearch( (void*) addr, &breakpointTreeRoot, compareBreakpoints);
//  msg("leaving addBreakpoint");
} // }

__declspec(dllexport) void __stdcall addDllFunctionBreakpoint(char* module, char* funcName, char* format) { // {

     address addr;
     HMODULE mod;
     char    breakpointName[200];
     char    txt           [400];

     TQ84_DEBUG_INDENT_T("addDllFunctionBreakpoint, module = %s, funcName = s, format = %s", module, funcName, format);

// MM     wsprintf(txt, "addDllFunctionBreakpoint %s / %s", module, funcName);
// MM     msg(txt);

     mod = GetModuleHandle(module);
     if (! mod) {
       TQ84_DEBUG("! mod");
       MessageBox(0, module, "GetModuleHandle", 0);
       return;
     }

     addr = GetProcAddress(mod, funcName);
     if (!addr) {
       TQ84_DEBUG("! addr");
          MessageBox(0, funcName, "GetProcAddress", 0);
          return;
     }

     wsprintf(breakpointName, "%s.%s", module, funcName);
// MM     msg(breakpointName);

     addBreakpoint(addr, breakpointName, format);
//   msg("leaving addDllFunctionBreakpoint");

} // }


#ifdef USE_SEARCH // {
address      addrLastBreakpointHit;
 // }
#else // {
address      func_addrs[NOF_BREAKPOINTS];
instruction  old_instr [NOF_BREAKPOINTS];
int   nrLastFuncBreakpointHit;
#endif // }



// typedef void (WINAPI   *addrCallBack_t)(BSTR);
   typedef void (CALLBACK *addrCallBack_t)(BSTR);
addrCallBack_t addrCallBack;


void callCallback(char* txt) { // {
     TQ84_DEBUG_INDENT_T("callCallback, txt = %s", txt);
     int wslen;
     BSTR bstr;
 
     wslen = MultiByteToWideChar(CP_ACP, 0, txt, strlen(txt),    0,     0); bstr  = SysAllocStringLen(0, wslen);
             MultiByteToWideChar(CP_ACP, 0, txt, strlen(txt), bstr, wslen);
 
    (*addrCallBack)(bstr);
 
     SysFreeString(bstr);
} // }

LONG WINAPI VectoredHandler(PEXCEPTION_POINTERS exPtr) { // {

    int i;
    char txt [400];
    char txt_[400];

    if      (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_BREAKPOINT ) { // {

#ifdef USE_SEARCH // {

       TQ84_DEBUG_INDENT_T("EXCEPTION_BREAKPOINT, addr = %d", exPtr->ContextRecord->Eip);


       breakpoint **breakpointFound;
       breakpoint  *breakpointHit;
       address addressHitByBreakpoint = (address) exPtr->ContextRecord->Eip;

//     TQ84_DEBUG("addressHitByBreakpoint = %d", exPtr->ContextRecord->Eip);
//     addressHitByBreakpoint = (address) (( (DWORD) addressHitByBreakpoint) - 1);

//MM     wsprintf(txt, "addressHitByBreakpoint = %d", addressHitByBreakpoint);
//MM     msg(txt);


       breakpoint findMe;
       findMe.addr = addressHitByBreakpoint;

       breakpointFound = tfind(&findMe, &breakpointTreeRoot, compareBreakpoints);


       if (!breakpointFound) { // {
         TQ84_DEBUG("Did not tfind breakpointHit...");
         MessageBox(0, "Did not tfind breakpointHit...", 0, 0);
       } // }
       else { // {

          breakpointHit = *breakpointFound;

//        wsprintf(txt, "&findMe = %d, breakpointHit = %d", &findMe, breakpointHit);
//        msg(txt);
//        wsprintf(txt, "breakpointHit = %d", breakpointHit);

//        wsprintf(txt ,"Hit Breakpoint, breakpointHit = %d, addr = %d (name = %s, addr(name) = %d), orig instr1stByte: %d", breakpointHit, breakpointHit->addr, breakpointHit->name, breakpointHit->name, breakpointHit->origInstruction);
          wsprintf(txt_,">>%s<<: >>%s<<", breakpointHit->name, breakpointHit->format);
//        TQ84_DEBUG("txt_: %s", txt_);

          unsigned long* esp = (unsigned long*) exPtr->ContextRecord->Esp;
          wsprintf(txt , txt_, *(esp+1), *(esp+2), *(esp+3), *(esp+4),*(esp+5), *(esp+6), *(esp+7), *(esp+8));
          TQ84_DEBUG("txt = %s", txt);

//        wsprintf(txt, "Hit Breakpoint %s", breakpointHit->name);
//        msg_2(txt);
          callCallback(txt);
   
          addrLastBreakpointHit = addressHitByBreakpoint;
   
          replaceInstruction(addressHitByBreakpoint, breakpointHit->origInstruction);

       // Single Step:
          exPtr->ContextRecord->EFlags |= 0x100;

       // TQ84_DEBUG("Going to call SetThreadContext");

          tq84_debug_dedent();
          SetThreadContext(GetCurrentThread(), exPtr->ContextRecord);
       } // }
 // }
#else // {


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
   
               exPtr->ContextRecord->EFlags |= 0x100;
   
            //
            // Resume execution:
            //
               SetThreadContext(GetCurrentThread(), exPtr->ContextRecord);
           }
        }
#endif // }

       TQ84_DEBUG  ("EXCEPTION_BREAKPOINT: no matching address found");
       callCallback("EXCEPTION_BREAKPOINT: no matching address found");
       return EXCEPTION_CONTINUE_SEARCH;

    } // }
    else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_SINGLE_STEP) { // {
       TQ84_DEBUG("EXCEPTION_SINGLE_STEP: Address: %d", exPtr->ContextRecord->Eip);


#ifdef USE_SEARCH  // {

//     TQ84_DEBUG("EXCEPTION_SINGLE_STEP: going to replaceInstruction at lastAddrBreakpoint hit: %d", addrLastBreakpointHit);
       replaceInstruction(addrLastBreakpointHit, 0xcc); // INT3              

 // }
#else // {

// QQ      replaceInstruction(func_addrs[nrLastFuncBreakpointHit], 0xcc);


     // callCallback("EXCEPTION_SINGLE_STEP");
 
     //
     // The processor is one instruction behind the last breakpoint that was
     // hit. We can now set the breakpoint again.
 
 
 //     TODO: Uncomment me:
 //     *func_addrs[nrLastFuncBreakpointHit] = 0xcc;
 // QQ  replaceInstruction(func_addrs[nrLastFuncBreakpointHit], 0xcc);
 
 //     msg_2("Single Step: Replacing instruction in single step");
 //     msg_2("Single Step: instruction was replaced, going to call SetThreadContext");
 
     //
     // Apparently, the TF flag is reset after the single execution, it
     // needs not be reset.
     //
     // Thus, we can resume execution again:
     //
 
     // callCallback("EXCEPTION_SINGLE_STEP -> SetThreadContext");

#endif // }

//     TQ84_DEBUG("Removing 0x100");
//     exPtr->ContextRecord->EFlags &= ~(0x100);
//     TQ84_DEBUG("Going to call SetThreadContext");
//     SetThreadContext(GetCurrentThread(), exPtr->ContextRecord);
       return EXCEPTION_CONTINUE_EXECUTION;



    } // }
    else { // {

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


           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_WAIT_0                        ) { wsprintf(txt, "Unexpected exception STATUS_WAIT_0                      , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ABANDONED_WAIT_0              ) { wsprintf(txt, "Unexpected exception STATUS_ABANDONED_WAIT_0            , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                              
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_USER_APC                      ) { wsprintf(txt, "Unexpected exception STATUS_USER_APC                    , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                      
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_TIMEOUT                       ) { wsprintf(txt, "Unexpected exception STATUS_TIMEOUT                     , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                     
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_PENDING                       ) { wsprintf(txt, "Unexpected exception STATUS_PENDING                     , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                     
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_EXCEPTION_HANDLED                ) { wsprintf(txt, "Unexpected exception DBG_EXCEPTION_HANDLED              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_CONTINUE                         ) { wsprintf(txt, "Unexpected exception DBG_CONTINUE                       , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                   
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_SEGMENT_NOTIFICATION          ) { wsprintf(txt, "Unexpected exception STATUS_SEGMENT_NOTIFICATION        , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FATAL_APP_EXIT                ) { wsprintf(txt, "Unexpected exception STATUS_FATAL_APP_EXIT              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_TERMINATE_THREAD                 ) { wsprintf(txt, "Unexpected exception DBG_TERMINATE_THREAD               , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                           
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_TERMINATE_PROCESS                ) { wsprintf(txt, "Unexpected exception DBG_TERMINATE_PROCESS              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_CONTROL_C                        ) { wsprintf(txt, "Unexpected exception DBG_CONTROL_C                      , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_PRINTEXCEPTION_C                 ) { wsprintf(txt, "Unexpected exception DBG_PRINTEXCEPTION_C               , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                           
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_RIPEXCEPTION                     ) { wsprintf(txt, "Unexpected exception DBG_RIPEXCEPTION                   , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                       
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_CONTROL_BREAK                    ) { wsprintf(txt, "Unexpected exception DBG_CONTROL_BREAK                  , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                        
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_COMMAND_EXCEPTION                ) { wsprintf(txt, "Unexpected exception DBG_COMMAND_EXCEPTION              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_GUARD_PAGE_VIOLATION          ) { wsprintf(txt, "Unexpected exception STATUS_GUARD_PAGE_VIOLATION        , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_DATATYPE_MISALIGNMENT         ) { wsprintf(txt, "Unexpected exception STATUS_DATATYPE_MISALIGNMENT       , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                   
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_BREAKPOINT                    ) { wsprintf(txt, "Unexpected exception STATUS_BREAKPOINT                  , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                        
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_SINGLE_STEP                   ) { wsprintf(txt, "Unexpected exception STATUS_SINGLE_STEP                 , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                         
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_LONGJUMP                      ) { wsprintf(txt, "Unexpected exception STATUS_LONGJUMP                    , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                      
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_UNWIND_CONSOLIDATE            ) { wsprintf(txt, "Unexpected exception STATUS_UNWIND_CONSOLIDATE          , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_EXCEPTION_NOT_HANDLED            ) { wsprintf(txt, "Unexpected exception DBG_EXCEPTION_NOT_HANDLED          , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ACCESS_VIOLATION              ) { wsprintf(txt, "Unexpected exception STATUS_ACCESS_VIOLATION            , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                              
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_IN_PAGE_ERROR                 ) { wsprintf(txt, "Unexpected exception STATUS_IN_PAGE_ERROR               , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                           
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INVALID_HANDLE                ) { wsprintf(txt, "Unexpected exception STATUS_INVALID_HANDLE              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INVALID_PARAMETER             ) { wsprintf(txt, "Unexpected exception STATUS_INVALID_PARAMETER           , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                               
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_NO_MEMORY                     ) { wsprintf(txt, "Unexpected exception STATUS_NO_MEMORY                   , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                       
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ILLEGAL_INSTRUCTION           ) { wsprintf(txt, "Unexpected exception STATUS_ILLEGAL_INSTRUCTION         , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                 
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_NONCONTINUABLE_EXCEPTION      ) { wsprintf(txt, "Unexpected exception STATUS_NONCONTINUABLE_EXCEPTION    , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                      
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INVALID_DISPOSITION           ) { wsprintf(txt, "Unexpected exception STATUS_INVALID_DISPOSITION         , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                 
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ARRAY_BOUNDS_EXCEEDED         ) { wsprintf(txt, "Unexpected exception STATUS_ARRAY_BOUNDS_EXCEEDED       , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                   
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_DENORMAL_OPERAND        ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_DENORMAL_OPERAND      , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_DIVIDE_BY_ZERO          ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_DIVIDE_BY_ZERO        , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_INEXACT_RESULT          ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_INEXACT_RESULT        , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_INVALID_OPERATION       ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_INVALID_OPERATION     , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                     
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_OVERFLOW                ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_OVERFLOW              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_STACK_CHECK             ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_STACK_CHECK           , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                               
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_UNDERFLOW               ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_UNDERFLOW             , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                             
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INTEGER_DIVIDE_BY_ZERO        ) { wsprintf(txt, "Unexpected exception STATUS_INTEGER_DIVIDE_BY_ZERO      , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INTEGER_OVERFLOW              ) { wsprintf(txt, "Unexpected exception STATUS_INTEGER_OVERFLOW            , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                              
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_PRIVILEGED_INSTRUCTION        ) { wsprintf(txt, "Unexpected exception STATUS_PRIVILEGED_INSTRUCTION      , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_STACK_OVERFLOW                ) { wsprintf(txt, "Unexpected exception STATUS_STACK_OVERFLOW              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_DLL_NOT_FOUND                 ) { wsprintf(txt, "Unexpected exception STATUS_DLL_NOT_FOUND               , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                           
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ORDINAL_NOT_FOUND             ) { wsprintf(txt, "Unexpected exception STATUS_ORDINAL_NOT_FOUND           , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                               
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ENTRYPOINT_NOT_FOUND          ) { wsprintf(txt, "Unexpected exception STATUS_ENTRYPOINT_NOT_FOUND        , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_CONTROL_C_EXIT                ) { wsprintf(txt, "Unexpected exception STATUS_CONTROL_C_EXIT              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_DLL_INIT_FAILED               ) { wsprintf(txt, "Unexpected exception STATUS_DLL_INIT_FAILED             , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                             
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_MULTIPLE_FAULTS         ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_MULTIPLE_FAULTS       , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                   
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_MULTIPLE_TRAPS          ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_MULTIPLE_TRAPS        , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_REG_NAT_CONSUMPTION           ) { wsprintf(txt, "Unexpected exception STATUS_REG_NAT_CONSUMPTION         , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                 
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_STACK_BUFFER_OVERRUN          ) { wsprintf(txt, "Unexpected exception STATUS_STACK_BUFFER_OVERRUN        , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INVALID_CRUNTIME_PARAMETER    ) { wsprintf(txt, "Unexpected exception STATUS_INVALID_CRUNTIME_PARAMETER  , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                        
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ASSERTION_FAILURE             ) { wsprintf(txt, "Unexpected exception STATUS_ASSERTION_FAILURE           , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                               
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_SXS_EARLY_DEACTIVATION        ) { wsprintf(txt, "Unexpected exception STATUS_SXS_EARLY_DEACTIVATION      , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_SXS_INVALID_DEACTIVATION      ) { wsprintf(txt, "Unexpected exception STATUS_SXS_INVALID_DEACTIVATION    , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                      

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
    } // }


    return EXCEPTION_CONTINUE_SEARCH;
} // }

__declspec(dllexport) void __stdcall VBAInternalsInit(addrCallBack_t addrCallBack_) { // {

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

} // }

BOOL WINAPI DllMain( // {
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

   TQ84_DEBUG("EXCEPTION_BREAKPOINT = %x, EXCEPTION_SINGLE_STEP = %x", EXCEPTION_BREAKPOINT, EXCEPTION_SINGLE_STEP);

// if      (fdwReason == DLL_PROCESS_ATTACH) {    MessageBox(0, "DllMain DLL_PROCESS_ATTACH", 0, 0)    ;}
// else if (fdwReason == DLL_PROCESS_DETACH) {    MessageBox(0, "DllMain DLL_PROCESS_DETACH", 0, 0)    ;}
// else if (fdwReason == DLL_THREAD_ATTACH ) { /* MessageBox(0, "DllMain DLL_THREAD_ATTACH ", 0, 0) */ ;}
// else if (fdwReason == DLL_THREAD_DETACH ) { /* MessageBox(0, "DllMain DLL_THREAD_DETACH ", 0, 0) */ ;}
// else                                      {    MessageBox(0, "DllMain hmmmm???"          , 0, 0)    ;}

   if (fdwReason == DLL_PROCESS_ATTACH) { // {
      TQ84_DEBUG("fdwReason == DLL_PROCESS_ATTACH");
      curProcess = GetCurrentProcess();
#ifdef USE_SEARCH
#else
      func_addrs[0] = (address) GetProcAddress(GetModuleHandle("VBE7.dll"), "rtcMsgBox");
      func_addrs[1] = (address) GetProcAddress(GetModuleHandle("VBE7.dll"), "rtcRound" );
#endif

   } // }

   if (fdwReason == DLL_PROCESS_DETACH) { // {
      TQ84_DEBUG("DLL_PROCESS_DETACH");
#ifdef USE_SEARCH
#else
      for (i=0; i<NOF_BREAKPOINTS; i++) {
//      TQ84_DEBUG("going to call replaceInstruction");
        replaceInstruction(func_addrs[i], old_instr[i]);
      }
#endif
   } // }

   if (fdwReason == DLL_THREAD_DETACH) { // {
//    TQ84_DEBUG("DLL_THREAD_DETACH");
   } // }

   if (fdwReason == DLL_THREAD_ATTACH) { // {
//    TQ84_DEBUG("DLL_THREAD_ATTACH");
   } // }

   return TRUE;

} // }
