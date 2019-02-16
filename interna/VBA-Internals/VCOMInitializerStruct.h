#include "C:\about\COM\c\structs\IDispatch_vTable.h"

typedef struct {

        IDispatch_vTable      iDispatch;
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

        void *nativeCode            ; // as string
        void *loaderMem             ; // as longPtr    // VirtualAlloc(len(.nativeCode)) --> then copied from nativeCode
        int   IgnoreFlag            ; // as boolean    // TypeOf .RootObject Is VBA.Collection
        void *vTablePtr             ; // as longPtr    // Does it point to the the start of THIS struct ?

        HANDLE kernel32Handle       ; // as longPtr

        void *rootObject            ; // as object     // VirtualAlloc(size = 325020); // 325020 =  ROOTOBJECT_SIZE  --> then: varPtr(
        void *classFactory          ; // as object     // A copy (?) of rootObject, thus a  call classFactory.init ...

} VCOMInitializerStruct;
