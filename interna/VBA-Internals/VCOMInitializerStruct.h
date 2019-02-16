#include "C:\about\COM\c\structs\IDispatch_vTable.h"

typedef struct {

        IDispatch_vTable      iDispatch;
//      vtbl_QueryInterface   ; // as longPtr
//      vtbl_AddRef           ; // as longPtr
//      vtbl_Release          ; // as longPtr
//      vtbl_GetTypeInfoCount ; // as longPtr
//      vtbl_GetTypeInfo      ; // as longPtr
//      vtbl_GetIDsOfNames    ; // as longPtr
//      vtbl_Invoke           ; // as longPtr
//      -------------------------------------
        void *RootObjectMem         ; // as longPtr

        void *HelperObject          ; // as object
        void *SysFreeString         ; // as longPtr
        void *WideCharToMultiByte   ; // as longPtr
        void *GetProcAddress        ; // as longPtr
        void *NativeCode            ; // as string
        void *LoaderMem             ; // as longPtr
        void *IgnoreFlag            ; // as boolean
        void *VTablePtr             ; // as longPtr
        void *Kernel32Handle        ; // as longPtr
        void *RootObject            ; // as object
        void *ClassFactory          ; // as object

} VCOMInitializerStruct;
