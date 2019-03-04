#include "C:\about\COM\c\structs\IDispatch_vTable.h"

#include "vbObject.h"



typedef struct VCOMInitializerStruct VCOMInitializerStruct;
struct VCOMInitializerStruct {

        IDispatch_vTable      iDispatch;                       // 4) = loaderMem
//      vtbl_QueryInterface   ; // as longPtr   
//      vtbl_AddRef           ; // as longPtr
//      vtbl_Release          ; // as longPtr
//      vtbl_GetTypeInfoCount ; // as longPtr
//      vtbl_GetTypeInfo      ; // as longPtr
//      vtbl_GetIDsOfNames    ; // as longPtr
//      vtbl_Invoke           ; // as longPtr
//      -------------------------------------
        void *RootObjectMem                          ; //   2) VirtualAlloc(RootObjectSize)

        void *helperObject                           ; //                     7) new ErrEx_Helper -- Obj with NATIVECODE\d\d\d\d_EVERYTHINGACCESS 

        void *SysFreeString         ; // as longPtr
        fn_WideCharToMultiByte  WideCharToMultiByte  ; //                 6)
        fn_GetProcAddress       GetProcAddress       ; //                 6)

        char                  *nativeCode            ; // 1) = "Cryptic String"
        void                  *loaderMem             ; //      3) = VirtualAlloc(…) 8) Copy from native code to loaderMem
        int                    ignoreFlag            ; //
        VCOMInitializerStruct *vTablePtr             ; //            5) setting back to start of this strcut

        HANDLE                 kernel32Handle        ; //                 6)

    //
    //  rootObject is an interface whose IUnknown or IDispatch points
    //  to the iDispatch table
    //
        VBAObject_withDisp     *rootObject            ; //                              9) copy from loaderMem to rootObject
        VBAObject_withDisp     *classFactory          ; // as object     // A copy (?) of rootObject, thus a  call classFactory.init ...

};
