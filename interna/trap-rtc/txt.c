#include <windows.h>
#include <strsafe.h>
#include "txt.h"
#include "error.h"

#define BUF_SIZE 1024
char buf[BUF_SIZE];

// typedef HRESULT (*fn_StringCchPrintfA)(char*, int, const char*, ...);
// fn_StringCchPrintfA tq84_StringCchPrintfA = 0;

typedef int (*fn_wvsprintfA)(char*, const char*, ...);
fn_wvsprintfA tq84_wvsprintfA = 0;


char *txt(const char *fmt, ...) {
  va_list ap;
  va_start(ap, fmt);


    if (!tq84_wvsprintfA) {

      HANDLE h = GetModuleHandle("User32.dll");
      if (!h) {
        MessageBox(0, Win32_LastError(), "GetModuleHandle User32.dll", 0);
      }
  
      tq84_wvsprintfA = (fn_wvsprintfA) GetProcAddress(h, "wvsprintfA");
      if (!tq84_wvsprintfA) {
        MessageBox(0, Win32_LastError(), "wvsprintfA", 0);
      }
  
    }


    tq84_wvsprintfA(buf, fmt, ap);


//StringCchPrintfA (buf, BUF_SIZE, fmt, ap);
//StringCbVPrint   (buf, BUF_SIZE, fmt, ap);
//wvnsprintfA      (buf, BUF_SIZE, fmt, ap);
//StringCbPrintfA  (buf, BUF_SIZE, fmt, ap);
//wsprintfA        (buf,           fmt, ap);
//wvsprintfA       (buf,           fmt, ap);
//StringCchVPrintfA(buf, BUF_SIZE, fmt, ap);

//wsprintf(buf, fmt, ap);

  va_end(ap);
  return buf;
}
