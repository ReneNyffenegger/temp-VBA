#include <windows.h>

// typedef void (*fnptr)(void (*)());

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
        FILE_APPEND_DATA,       // open for writing
        FILE_SHARE_WRITE,       // 
        NULL,                   // default security
        OPEN_ALWAYS,  // CREATE_NEW,             // create new file only
        FILE_ATTRIBUTE_NORMAL,  // normal file
        NULL);                  // no attr. template

  DWORD x;

  WriteFile(hFile, buf   , strlen(buf ), &x, 0);
//WriteFile(hFile, text  , strlen(text), 0 , 0);
  WriteFile(hFile, "\x0a",           1 , &x, 0);

  CloseHandle(hFile);



}

// int main() {
// 
//   writeToFile("The number is %d, the text is >%s<", 42, "Hello world");
//   writeToFile("second line");
// 
// }


// long *vTable;

typedef __stdcall HRESULT (*funcPtr_IUnknown_QueryInterface)(void*, const IID*, void**);
typedef __stdcall HRESULT (*funcPtr_IUnknown_AddRef        )(void*);
typedef __stdcall HRESULT (*funcPtr_IUnknown_Release       )(void*);

struct IUnkown_vTable {
  funcPtr_IUnknown_QueryInterface  QueryInterface;
  funcPtr_IUnknown_AddRef          AddRef;
  funcPtr_IUnknown_Release         Release;
};

struct IUnkown_vTable *vTable;


//typedef void (*funcPtr_void_void)(void);
// funcPtr_void_void  AddRefOrig;
funcPtr_IUnknown_QueryInterface QueryInterfaceOrig;;
funcPtr_IUnknown_AddRef         AddRefOrig;
funcPtr_IUnknown_Release        ReleaseOrig;

__stdcall HRESULT QueryInterface(void* this, const IID* iid, void** ppv);
__stdcall HRESULT AddRef        (void* this);
__stdcall HRESULT Release       (void* this);

__declspec(dllexport) __stdcall void ChangeVTable(long** addrObj) {

  if (vTableIsChanged) {
    writeToFile("setFuncPointersInVTable, vTable is already changed, returning");
    vTableIsChanged = 1;
    return;
  }

  writeToFile("setFuncPointersInVTable, addrObj = %p", addrObj);
  vTable =  (struct IUnkown_vTable*) *addrObj;
  writeToFile("setFuncPointersInVTable, vTable = %p", vTable);

//AddRefOrig = (funcPtr_void_void)        vTable[1];
//AddRefOrig =                            vTable[1];
  QueryInterfaceOrig =                    vTable->QueryInterface;
  AddRefOrig         =                    vTable->AddRef;
  ReleaseOrig        =                    vTable->Release;


//vTable[1]  = (long ) &AddRef;
  vTable->QueryInterface = &QueryInterface;
  vTable->AddRef         = &AddRef;
  vTable->Release        = &Release;

  vTableIsChanged = 1;
}

__stdcall HRESULT QueryInterface(void* this, const IID* iid, void** ppv) {

  PUSH_REGISTERS

  writeToFile("QueryInterface, this = %p");

  POP_REGISTERS

  JMP_TO_FUNCTION(QueryInterfaceOrig)

//   asm (
//       "movl %%ebp, %%esp\n"    // Undo initialization of
//       "popl %%ebp\n"           // stackframe.
//    // ----------------------------------------------
// //    "pushl (_AddRefOrig)\n"  // Jump to original address
// //    "ret\n"                  //
//       "jmpl *_QueryInterfaceOrig\n"    // Jump to original address
//       :
// //    : "r" (AddRefOrig)
//   );

}

__stdcall HRESULT AddRef(void* this) {

  PUSH_REGISTERS

  writeToFile("AddRef, this = %p");

  POP_REGISTERS

  JMP_TO_FUNCTION(AddRefOrig)

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

__stdcall HRESULT Release(void* this) {

  PUSH_REGISTERS

  writeToFile("Release, this = %p");

  POP_REGISTERS

  JMP_TO_FUNCTION(ReleaseOrig)

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
