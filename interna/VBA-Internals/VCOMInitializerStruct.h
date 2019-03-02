#include "C:\about\COM\c\structs\IDispatch_vTable.h"

#include "vbObject.h"



typedef struct VCOMInitializerStruct VCOMInitializerStruct;
struct VCOMInitializerStruct {

        IDispatch_vTable      iDispatch;                 // = loaderMem, apparently works because first method invoked?
//      vtbl_QueryInterface   ; // as longPtr            // = loaderMem, apparently works because first method invoked?
//      vtbl_AddRef           ; // as longPtr
//      vtbl_Release          ; // as longPtr
//      vtbl_GetTypeInfoCount ; // as longPtr
//      vtbl_GetTypeInfo      ; // as longPtr
//      vtbl_GetIDsOfNames    ; // as longPtr
//      vtbl_Invoke           ; // as longPtr
//      -------------------------------------
        void *RootObjectMem         ; // as longPtr

        void *helperObject          ; // new ErrEx_Helper -- Obj with NATIVECODE\d\d\d\d_EVERYTHINGACCESS 

        void *SysFreeString         ; // as longPtr
        fn_WideCharToMultiByte  WideCharToMultiByte;
        fn_GetProcAddress       GetProcAddress;

        char                  *nativeCode            ; // as string
        void                  *loaderMem             ; // as longPtr    // VirtualAlloc(len(.nativeCode)) --> then copied from nativeCode
        int                    ignoreFlag            ; // as boolean    // TypeOf .RootObject Is VBA.Collection
        VCOMInitializerStruct *vTablePtr; // as longPtr    // Points to the start of this struct 

        HANDLE                 kernel32Handle       ; // as longPtr

    //
    //  rootObject is an interface whose IUnknown or IDispatch points
    //  to the iDispatch table
    //
//      IDispatch_vTable       *rootObject            ; // as object     // VirtualAlloc(size = 325020); // 325020 =  ROOTOBJECT_SIZE  --> then: varPtr(…)
        minimalVBAObject       *rootObject            ; // as object     // VirtualAlloc(size = 325020); // 325020 =  ROOTOBJECT_SIZE  --> then: varPtr(…)
        void                   *classFactory          ; // as object     // A copy (?) of rootObject, thus a  call classFactory.init ...

};
