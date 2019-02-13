//
//  TODO: IsBadReadPtr(…)
//        DllVbeSetMsiApis
//
#include "msg.h"
#include "txt.h"
#include "funcsInDll.h"
#include "int3.h"

HANDLE msgFile;

void callbackFuncInDll(char *funcName, DWORD address) {
    writeToFile(txt("callbackFuncInDll received %s @ %d\n", funcName, address));
 // writeToFile(txt("callbackFuncInDll received %s @                         ));


    HANDLE   h = GetModuleHandle("VBE7.dll");
    FARPROC  a = GetProcAddress(h, funcName);
    writeToFile(txt("proc addr: %d\n", a) );

    add_breakpoint(a, funcName);
    writeToFile("breakpoint added\n");

//writeToFile("\n");

}

LONG WINAPI VectoredHandler(PEXCEPTION_POINTERS exPtr) { // {


    if      (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_BREAKPOINT ) { // {

       writeToFile(txt("EXCEPTION_BREAKPOINT at %d\n", exPtr->ContextRecord->Eip));

       breakpoint *bp = addr_to_breakpoint((int3_addr) exPtr->ContextRecord->Eip);

       writeToFile(txt("  name = %s\n", bp->name));

       int3_to_orig(bp);

       SetThreadContext(GetCurrentThread(), exPtr->ContextRecord);


    } //
    else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_SINGLE_STEP) { // {

       writeToFile(txt("EXCEPTION_SINGLE_STEP at %d\n", exPtr->ContextRecord->Eip));

//     replaceInstruction(addrLastBreakpointHit, 0xcc); // INT3              
       
       breakpoint *bp = addr_to_breakpoint((int3_addr) exPtr->ContextRecord->Eip);
       writeToFile(txt("  name = %s\n", bp->name));

       set_int3(bp);

       return EXCEPTION_CONTINUE_EXECUTION;


    } // }
    else {

           if      (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_ACCESS_VIOLATION           ) { writeToFile(txt("Unexpected exception EXCEPTION_ACCESS_VIOLATION         , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_DATATYPE_MISALIGNMENT      ) { writeToFile(txt("Unexpected exception EXCEPTION_DATATYPE_MISALIGNMENT    , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_BREAKPOINT                 ) { writeToFile(txt("Unexpected exception EXCEPTION_BREAKPOINT               , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_SINGLE_STEP                ) { writeToFile(txt("Unexpected exception EXCEPTION_SINGLE_STEP              , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_ARRAY_BOUNDS_EXCEEDED      ) { writeToFile(txt("Unexpected exception EXCEPTION_ARRAY_BOUNDS_EXCEEDED    , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_DENORMAL_OPERAND       ) { writeToFile(txt("Unexpected exception EXCEPTION_FLT_DENORMAL_OPERAND     , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_DIVIDE_BY_ZERO         ) { writeToFile(txt("Unexpected exception EXCEPTION_FLT_DIVIDE_BY_ZERO       , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_INEXACT_RESULT         ) { writeToFile(txt("Unexpected exception EXCEPTION_FLT_INEXACT_RESULT       , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_INVALID_OPERATION      ) { writeToFile(txt("Unexpected exception EXCEPTION_FLT_INVALID_OPERATION    , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_OVERFLOW               ) { writeToFile(txt("Unexpected exception EXCEPTION_FLT_OVERFLOW             , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_STACK_CHECK            ) { writeToFile(txt("Unexpected exception EXCEPTION_FLT_STACK_CHECK          , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_UNDERFLOW              ) { writeToFile(txt("Unexpected exception EXCEPTION_FLT_UNDERFLOW            , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_INT_DIVIDE_BY_ZERO         ) { writeToFile(txt("Unexpected exception EXCEPTION_INT_DIVIDE_BY_ZERO       , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_INT_OVERFLOW               ) { writeToFile(txt("Unexpected exception EXCEPTION_INT_OVERFLOW             , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_PRIV_INSTRUCTION           ) { writeToFile(txt("Unexpected exception EXCEPTION_PRIV_INSTRUCTION         , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_IN_PAGE_ERROR              ) { writeToFile(txt("Unexpected exception EXCEPTION_IN_PAGE_ERROR            , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_ILLEGAL_INSTRUCTION        ) { writeToFile(txt("Unexpected exception EXCEPTION_ILLEGAL_INSTRUCTION      , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_NONCONTINUABLE_EXCEPTION   ) { writeToFile(txt("Unexpected exception EXCEPTION_NONCONTINUABLE_EXCEPTION , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_STACK_OVERFLOW             ) { writeToFile(txt("Unexpected exception EXCEPTION_STACK_OVERFLOW           , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_INVALID_DISPOSITION        ) { writeToFile(txt("Unexpected exception EXCEPTION_INVALID_DISPOSITION      , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_GUARD_PAGE                 ) { writeToFile(txt("Unexpected exception EXCEPTION_GUARD_PAGE               , addr = %d\n", exPtr->ContextRecord->Eip)); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_INVALID_HANDLE             ) { writeToFile(txt("Unexpected exception EXCEPTION_INVALID_HANDLE           , addr = %d\n", exPtr->ContextRecord->Eip)); }


           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_WAIT_0                        ) { writeToFile(txt("Unexpected exception STATUS_WAIT_0                      , addr = %d\n", exPtr->ContextRecord->Eip)); }                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ABANDONED_WAIT_0              ) { writeToFile(txt("Unexpected exception STATUS_ABANDONED_WAIT_0            , addr = %d\n", exPtr->ContextRecord->Eip)); }                                              
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_USER_APC                      ) { writeToFile(txt("Unexpected exception STATUS_USER_APC                    , addr = %d\n", exPtr->ContextRecord->Eip)); }                                      
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_TIMEOUT                       ) { writeToFile(txt("Unexpected exception STATUS_TIMEOUT                     , addr = %d\n", exPtr->ContextRecord->Eip)); }                                     
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_PENDING                       ) { writeToFile(txt("Unexpected exception STATUS_PENDING                     , addr = %d\n", exPtr->ContextRecord->Eip)); }                                     
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_EXCEPTION_HANDLED                ) { writeToFile(txt("Unexpected exception DBG_EXCEPTION_HANDLED              , addr = %d\n", exPtr->ContextRecord->Eip)); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_CONTINUE                         ) { writeToFile(txt("Unexpected exception DBG_CONTINUE                       , addr = %d\n", exPtr->ContextRecord->Eip)); }                                   
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_SEGMENT_NOTIFICATION          ) { writeToFile(txt("Unexpected exception STATUS_SEGMENT_NOTIFICATION        , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FATAL_APP_EXIT                ) { writeToFile(txt("Unexpected exception STATUS_FATAL_APP_EXIT              , addr = %d\n", exPtr->ContextRecord->Eip)); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_TERMINATE_THREAD                 ) { writeToFile(txt("Unexpected exception DBG_TERMINATE_THREAD               , addr = %d\n", exPtr->ContextRecord->Eip)); }                                           
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_TERMINATE_PROCESS                ) { writeToFile(txt("Unexpected exception DBG_TERMINATE_PROCESS              , addr = %d\n", exPtr->ContextRecord->Eip)); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_CONTROL_C                        ) { writeToFile(txt("Unexpected exception DBG_CONTROL_C                      , addr = %d\n", exPtr->ContextRecord->Eip)); }                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_PRINTEXCEPTION_C                 ) { writeToFile(txt("Unexpected exception DBG_PRINTEXCEPTION_C               , addr = %d\n", exPtr->ContextRecord->Eip)); }                                           
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_RIPEXCEPTION                     ) { writeToFile(txt("Unexpected exception DBG_RIPEXCEPTION                   , addr = %d\n", exPtr->ContextRecord->Eip)); }                                       
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_CONTROL_BREAK                    ) { writeToFile(txt("Unexpected exception DBG_CONTROL_BREAK                  , addr = %d\n", exPtr->ContextRecord->Eip)); }                                        
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_COMMAND_EXCEPTION                ) { writeToFile(txt("Unexpected exception DBG_COMMAND_EXCEPTION              , addr = %d\n", exPtr->ContextRecord->Eip)); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_GUARD_PAGE_VIOLATION          ) { writeToFile(txt("Unexpected exception STATUS_GUARD_PAGE_VIOLATION        , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_DATATYPE_MISALIGNMENT         ) { writeToFile(txt("Unexpected exception STATUS_DATATYPE_MISALIGNMENT       , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                   
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_BREAKPOINT                    ) { writeToFile(txt("Unexpected exception STATUS_BREAKPOINT                  , addr = %d\n", exPtr->ContextRecord->Eip)); }                                        
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_SINGLE_STEP                   ) { writeToFile(txt("Unexpected exception STATUS_SINGLE_STEP                 , addr = %d\n", exPtr->ContextRecord->Eip)); }                                         
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_LONGJUMP                      ) { writeToFile(txt("Unexpected exception STATUS_LONGJUMP                    , addr = %d\n", exPtr->ContextRecord->Eip)); }                                      
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_UNWIND_CONSOLIDATE            ) { writeToFile(txt("Unexpected exception STATUS_UNWIND_CONSOLIDATE          , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_EXCEPTION_NOT_HANDLED            ) { writeToFile(txt("Unexpected exception DBG_EXCEPTION_NOT_HANDLED          , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ACCESS_VIOLATION              ) { writeToFile(txt("Unexpected exception STATUS_ACCESS_VIOLATION            , addr = %d\n", exPtr->ContextRecord->Eip)); }                                              
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_IN_PAGE_ERROR                 ) { writeToFile(txt("Unexpected exception STATUS_IN_PAGE_ERROR               , addr = %d\n", exPtr->ContextRecord->Eip)); }                                           
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INVALID_HANDLE                ) { writeToFile(txt("Unexpected exception STATUS_INVALID_HANDLE              , addr = %d\n", exPtr->ContextRecord->Eip)); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INVALID_PARAMETER             ) { writeToFile(txt("Unexpected exception STATUS_INVALID_PARAMETER           , addr = %d\n", exPtr->ContextRecord->Eip)); }                                               
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_NO_MEMORY                     ) { writeToFile(txt("Unexpected exception STATUS_NO_MEMORY                   , addr = %d\n", exPtr->ContextRecord->Eip)); }                                       
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ILLEGAL_INSTRUCTION           ) { writeToFile(txt("Unexpected exception STATUS_ILLEGAL_INSTRUCTION         , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                 
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_NONCONTINUABLE_EXCEPTION      ) { writeToFile(txt("Unexpected exception STATUS_NONCONTINUABLE_EXCEPTION    , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                      
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INVALID_DISPOSITION           ) { writeToFile(txt("Unexpected exception STATUS_INVALID_DISPOSITION         , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                 
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ARRAY_BOUNDS_EXCEEDED         ) { writeToFile(txt("Unexpected exception STATUS_ARRAY_BOUNDS_EXCEEDED       , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                   
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_DENORMAL_OPERAND        ) { writeToFile(txt("Unexpected exception STATUS_FLOAT_DENORMAL_OPERAND      , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_DIVIDE_BY_ZERO          ) { writeToFile(txt("Unexpected exception STATUS_FLOAT_DIVIDE_BY_ZERO        , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_INEXACT_RESULT          ) { writeToFile(txt("Unexpected exception STATUS_FLOAT_INEXACT_RESULT        , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_INVALID_OPERATION       ) { writeToFile(txt("Unexpected exception STATUS_FLOAT_INVALID_OPERATION     , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                     
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_OVERFLOW                ) { writeToFile(txt("Unexpected exception STATUS_FLOAT_OVERFLOW              , addr = %d\n", exPtr->ContextRecord->Eip)); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_STACK_CHECK             ) { writeToFile(txt("Unexpected exception STATUS_FLOAT_STACK_CHECK           , addr = %d\n", exPtr->ContextRecord->Eip)); }                                               
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_UNDERFLOW               ) { writeToFile(txt("Unexpected exception STATUS_FLOAT_UNDERFLOW             , addr = %d\n", exPtr->ContextRecord->Eip)); }                                             
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INTEGER_DIVIDE_BY_ZERO        ) { writeToFile(txt("Unexpected exception STATUS_INTEGER_DIVIDE_BY_ZERO      , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INTEGER_OVERFLOW              ) { writeToFile(txt("Unexpected exception STATUS_INTEGER_OVERFLOW            , addr = %d\n", exPtr->ContextRecord->Eip)); }                                              
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_PRIVILEGED_INSTRUCTION        ) { writeToFile(txt("Unexpected exception STATUS_PRIVILEGED_INSTRUCTION      , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_STACK_OVERFLOW                ) { writeToFile(txt("Unexpected exception STATUS_STACK_OVERFLOW              , addr = %d\n", exPtr->ContextRecord->Eip)); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_DLL_NOT_FOUND                 ) { writeToFile(txt("Unexpected exception STATUS_DLL_NOT_FOUND               , addr = %d\n", exPtr->ContextRecord->Eip)); }                                           
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ORDINAL_NOT_FOUND             ) { writeToFile(txt("Unexpected exception STATUS_ORDINAL_NOT_FOUND           , addr = %d\n", exPtr->ContextRecord->Eip)); }                                               
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ENTRYPOINT_NOT_FOUND          ) { writeToFile(txt("Unexpected exception STATUS_ENTRYPOINT_NOT_FOUND        , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_CONTROL_C_EXIT                ) { writeToFile(txt("Unexpected exception STATUS_CONTROL_C_EXIT              , addr = %d\n", exPtr->ContextRecord->Eip)); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_DLL_INIT_FAILED               ) { writeToFile(txt("Unexpected exception STATUS_DLL_INIT_FAILED             , addr = %d\n", exPtr->ContextRecord->Eip)); }                                             
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_MULTIPLE_FAULTS         ) { writeToFile(txt("Unexpected exception STATUS_FLOAT_MULTIPLE_FAULTS       , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                   
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_MULTIPLE_TRAPS          ) { writeToFile(txt("Unexpected exception STATUS_FLOAT_MULTIPLE_TRAPS        , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_REG_NAT_CONSUMPTION           ) { writeToFile(txt("Unexpected exception STATUS_REG_NAT_CONSUMPTION         , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                 
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_STACK_BUFFER_OVERRUN          ) { writeToFile(txt("Unexpected exception STATUS_STACK_BUFFER_OVERRUN        , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INVALID_CRUNTIME_PARAMETER    ) { writeToFile(txt("Unexpected exception STATUS_INVALID_CRUNTIME_PARAMETER  , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                        
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ASSERTION_FAILURE             ) { writeToFile(txt("Unexpected exception STATUS_ASSERTION_FAILURE           , addr = %d\n", exPtr->ContextRecord->Eip)); }                                               
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_SXS_EARLY_DEACTIVATION        ) { writeToFile(txt("Unexpected exception STATUS_SXS_EARLY_DEACTIVATION      , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_SXS_INVALID_DEACTIVATION      ) { writeToFile(txt("Unexpected exception STATUS_SXS_INVALID_DEACTIVATION    , addr = %d\n", exPtr->ContextRecord->Eip)); }                                                      
           else                                                                                     { writeToFile(txt("Unexpected exception ?                                  , addr = %d, code = %d\n", exPtr->ContextRecord->Eip, exPtr->ExceptionRecord->ExceptionCode)); }

           MessageBoxA(0, "Should not be reached", "VectoredHandler", 0);

           return EXCEPTION_CONTINUE_SEARCH;

    }


    return EXCEPTION_CONTINUE_SEARCH;

} // }

__declspec(dllexport) void __stdcall Init () { // {

    createNewOutputFile(TEXT("c:\\temp\\trap-rtc.msg"));

    writeToFile("Initialized\n");

    AddVectoredExceptionHandler(1, VectoredHandler);
    writeToFile("VectoredHandler added\n");

    iterateOverFuncsInDll("VBE7.dll", "C:\\Program Files\\Common Files\\microsoft shared\\VBA\\VBA7", callbackFuncInDll);

} // }

BOOL DllMain(HINSTANCE hInst, DWORD reason, LPVOID reserved) { // {
    
    if (reason == DLL_PROCESS_DETACH) {

        writeToFile("DllMain, DLL_PROCESS_DETACH -> rm_all_int3\n");
        rm_all_int3();

    }
    return TRUE;
} // }
