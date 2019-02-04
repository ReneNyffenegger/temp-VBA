#include <windows.h>

// typedef void (*fnptr)(void (*)());
long eax_c;

#define PUSH_REGISTERS  \
  asm ("pushl %eax\n"   \
       "pushl %ebx\n"   \
       "pushl %ecx\n"   \
       "pushl %edx\n");

#define POP_REGISTERS   \
  asm ("popl %edx\n"    \
       "popl %ecx\n"    \
       "popl %ebx\n"    \
       "popl %eax\n");

#define JMP_TO_FUNCTION(function)                                     \
  asm (                                                               \
      "movl %%ebp, %%esp\n"    /* Undo initialization of   */         \
      "popl %%ebp\n"           /* stackframe.              */         \
   /* ----------------------------------------------       */         \
/*    "pushl (_AddRefOrig)\n"  -- Jump to original address */         \
/*    "ret\n"                  --                          */         \
      "jmpl *_" #function "\n" /* Jump to original address */         \
      :                                                               \
      : /* "r" (AddRefOrig)   */                                      \
  );

#define PUSH_RET_FUNCTION(function)   \
    asm(" pushl _" #function "\n"     \
        " ret\n"                      \
    );

//
// https://gist.github.com/tfzxyinhao/2818b31a7ce94154a133
//
char* unicode2ascii(const wchar_t* wszString) {

	int outlen = WideCharToMultiByte(CP_ACP, 0, wszString, wcslen(wszString), NULL, 0, NULL, NULL);
	char* utf8 = malloc(outlen + 1);
	WideCharToMultiByte(CP_ACP, 0, wszString, wcslen(wszString), utf8, outlen, NULL, NULL);
	utf8[outlen] = '\0';
	return utf8;
}

int vTableIsChanged = 0;

void writeToFile(char* fmt, ...) {

  va_list args;

  char buf[1024];

  va_start(args, fmt);
 //
 // wvsprintf is the WinAPI equivalent for vsprintf
 //
  wvsprintf(buf, fmt, args);
  va_end(args);

  HANDLE hFile = CreateFile(
       "c:\\temp\\vTableTest.txt", // name of the write
        FILE_APPEND_DATA,          // open for writing
        FILE_SHARE_WRITE,          // 
        NULL,                      // default security
        OPEN_ALWAYS,  // CREATE_NEW,             // create new file only
        FILE_ATTRIBUTE_NORMAL,     // normal file
        NULL);                     // no attr. template

  DWORD x;

  WriteFile(hFile, buf   , strlen(buf ), &x, 0);
//WriteFile(hFile, text  , strlen(text), 0 , 0);
  WriteFile(hFile, "\x0a",           1 , &x, 0);

  CloseHandle(hFile);

}

__declspec(dllexport) __stdcall void dbg(char* txt) {
    writeToFile(txt);
}

// int main() {
// 
//   writeToFile("The number is %d, the text is >%s<", 42, "Hello world");
//   writeToFile("second line");
// 
// }


// long *vTable;

//  IUnknown_vTable 
//  IDispatch_vTable:
//    --> GONE TO about-COM/c/structs!

struct vb_object {

 // struct IUnknown_vTable   *vTbl;
    struct IDispatch_vTable *vTbl;   // Probably 3rd element
 
    long  b[6];
 
    struct IUnknown_vTable * iunknown; // Probably first element
    ULONG  refCnt;

};


// Is this a better layout?
//
// typedef struct {
// 
//   IUnknown_vTable  *vTable_IUnknown;
//   unsigned long     refCount;
//   IDispatch_vTable *vTable_IDispatch;
//   
// } vbObject;


//__declspec(dllexport) __stdcall void ChangeVTable(struct IUnknown_vTable** addrObj)
__declspec(dllexport) __stdcall void ChangeVTable(struct vb_object* addrObj)
{

  if (vTableIsChanged) {
    writeToFile("ChangeVTable, vTable is already changed, returning");
    vTableIsChanged = 1;
    return;
  }

  writeToFile("ChangeVTable, addrObj = %p", addrObj);
//vTable =  (struct IUnknown_vTable*) *addrObj;
//vTable =                           *addrObj;
  vTable =                            addrObj->vTbl;
  writeToFile("ChangeVTable, vTable = %p", vTable);

//AddRefOrig = (funcPtr_void_void)      vTable[1];
//AddRefOrig =                          vTable[1];
  QueryInterfaceOrig   =                vTable->vtbl_IUnknown.QueryInterface;
  AddRefOrig           =                vTable->vtbl_IUnknown.AddRef;
  ReleaseOrig          =                vTable->vtbl_IUnknown.Release;

  GetTypeInfoCountOrig =                vTable->GetTypeInfoCount;
  GetTypeInfoOrig      =                vTable->GetTypeInfo;
  GetIDsOfNamesOrig    =                vTable->GetIDsOfNames;
  InvokeOrig           =                vTable->Invoke;

  vTable->vtbl_IUnknown.QueryInterface = &QueryInterface;
  vTable->vtbl_IUnknown.AddRef         = &AddRef;
  vTable->vtbl_IUnknown.Release        = &Release;

  vTable->GetTypeInfoCount             = &GetTypeInfoCount;
  vTable->GetTypeInfo                  = &GetTypeInfo;
  vTable->GetIDsOfNames                = &GetIDsOfNames;
  vTable->Invoke                       = &Invoke;

  vTableIsChanged = 1;
}

__stdcall HRESULT QueryInterface(void* this, const REFIID iid, void** ppv) {

 // this =  8(%ebp)
 // ppv  = 16(%ebp)

//  PUSH_REGISTERS
  
    wchar_t* str;
    StringFromIID(iid, &str);
  
    char* strIID = unicode2ascii(str);
    writeToFile("QueryInterface, this = %p, iid = %s", this, strIID);
    free(strIID);
    CoTaskMemFree(str);
  
    HRESULT r = QueryInterfaceOrig(this, iid, ppv);
    writeToFile("r = %x, ppv = %p", r, ppv);

//  POP_REGISTERS

    return r;

//    asm ("pushl 16(%ebp) \n"
//       "  pushl 12(%ebp) \n"
//       "  pushl  8(%ebp) \n"
//
////     "  pushl  4(%ebp) \n"
////     "  movl  _QueryInterfaceOrig, %eax\n"
//       
//       "  call _QueryInterfaceOrig\n"
//       "  int3\n"
//       );



//
////JMP_TO_FUNCTION(QueryInterfaceOrig)
//  PUSH_RET_FUNCTION(QueryInterfaceOrig)
//
//  PUSH_REGISTERS
  
//  asm("movl %eax, _eax_c\n");
//  writeToFile("eax = %d", eax_c);
//
//  POP_REGISTERS

}

__stdcall HRESULT AddRef(void* this) {

  struct vb_object *vb_obj = (struct vb_object*) this;

//PUSH_REGISTERS

//writeToFile("AddRef, this = %p", this);

  HRESULT r = AddRefOrig(this);

//POP_REGISTERS

  writeToFile("AddRef, r = %d, refCnt = %d", r, vb_obj->refCnt);

  return r;

//JMP_TO_FUNCTION(AddRefOrig)

//   asm (
//       "movl %%ebp, %%esp\n"    // Undo initialization of
//       "popl %%ebp\n"           // stackframe.
//    // ----------------------------------------------
// //    "pushl (_AddRefOrig)\n"  // Jump to original address
// //    "ret\n"                  //
//       "jmpl *_AddRefOrig\n"    // Jump to original address
//       :
// //    : "r" (AddRefOrig)
//   );

}

   __stdcall HRESULT Release(void* this)
// __stdcall HRESULT Release(struct vb_object* this)
{

  struct vb_object* vb_obj = (struct vb_object*) this;

//PUSH_REGISTERS

  writeToFile("Release, vb_object = %p, refCnt: %d", vb_obj, vb_obj->refCnt); //, (int) *(this+4));

  HRESULT r = ReleaseOrig(this);
  writeToFile("Release, r = %d / refCnt: %d", r, vb_obj->refCnt); //, (int) *(this+4));
  return r;

//POP_REGISTERS

//JMP_TO_FUNCTION(ReleaseOrig)

//  asm (
//      "movl %%ebp, %%esp\n"    // Undo initialization of
//      "popl %%ebp\n"           // stackframe.
//   // ----------------------------------------------
////    "pushl (_AddRefOrig)\n"  // Jump to original address
////    "ret\n"                  //
//      "jmpl *_ReleaseOrig\n"   // Jump to original address
//      :
////    : "r" (AddRefOrig)
//  );

}

__stdcall HRESULT GetTypeInfoCount(void* this, UINT *pctInfo) {

  struct vb_object* vb_obj = (struct vb_object*) this;
  writeToFile("GetTypeInfoCount, vb_object = %p", vb_obj);

  HRESULT r = GetTypeInfoCountOrig(this, pctInfo);
  writeToFile("         r = %d / pctInfo: %d", r, pctInfo);
  return r;

}

__stdcall HRESULT GetTypeInfo     (void* this, UINT iTInfo, LCID lcid, ITypeInfo **ppTI) {
  writeToFile("GetTypeInfo");
  return GetTypeInfoOrig(this, iTInfo, lcid, ppTI);

}
__stdcall HRESULT GetIDsOfNames   (void* this, REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) {
  writeToFile("GetIDsOfNames");
  return GetIDsOfNamesOrig(this, riid, rgszNames, cNames, lcid, rgDispId);

}
__stdcall HRESULT Invoke          (void* this, DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pvarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) {
  writeToFile("Invoke");
  return InvokeOrig               (      this,        dispidMember,        riid,      lcid,      wFlags,             pDispParams, pvarResult, pExcepInfo, puArgErr);
}
