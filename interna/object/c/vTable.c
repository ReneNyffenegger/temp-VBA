#include <windows.h>

// typedef void (*fnptr)(void (*)());

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

typedef __stdcall unsigned long (*funcPtr_IUnknown_AddRef)(void*);

struct IUnkown_vTable {
  void* todo              ;
  funcPtr_IUnknown_AddRef AddRef;
};

struct IUnkown_vTable *vTable;


typedef void (*funcPtr_void_void)(void);
// funcPtr_void_void  AddRefOrig;
funcPtr_IUnknown_AddRef AddRefOrig;

// static __attribute__((stdcall)) void AddRef(void);
__stdcall unsigned long AddRef(void* this);

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
  AddRefOrig =                            vTable->AddRef;


//vTable[1]  = (long ) &AddRef;
  vTable->AddRef = &AddRef;

  vTableIsChanged = 1;
}


// static __attribute__((stdcall)) void AddRef(void) {
__stdcall unsigned long AddRef(void* this) {

  asm ("pushl %eax\n"
       "pushl %ebx\n"
       "pushl %ecx\n"
       "pushl %edx\n");

   writeToFile("AddRef, this = %p");

  asm ("popl %edx\n"
       "popl %ecx\n"
       "popl %ebx\n"
       "popl %eax\n");

//(*AddRefOrig)();

//writeToFile("AddRef, this = %p", this);
  asm (
      "movl %%ebp, %%esp\n"    // Undo initialization of
      "popl %%ebp\n"           // stackframe.
   // ----------------------------------------------
//    "pushl (_AddRefOrig)\n"  // Jump to original address
//    "ret\n"                  //
      "jmpl *_AddRefOrig\n"    // Jump to original address
      :
//    : "r" (AddRefOrig)
  );

}
