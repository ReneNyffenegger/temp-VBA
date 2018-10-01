#include <windows.h>

// typedef void (*fnptr)(void (*)());


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

int main() {

  writeToFile("The number is %d, the text is >%s<", 42, "Hello world");
  writeToFile("second line");

}


static __attribute__((stdcall)) AddRef(void* this) {

}
