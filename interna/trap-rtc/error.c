//
// http://www.adp-gmbh.ch/win/programming/common/win32_lasterror.html
//

#include <windows.h>
#include "error.h"

#define SIZE_BUF 512
char buf[SIZE_BUF];

char *Win32_LastError() {
//LPVOID lpMsgBuf;
//std::string ret;
  if (!FormatMessageA( 
//    FORMAT_MESSAGE_ALLOCATE_BUFFER | 
      FORMAT_MESSAGE_FROM_SYSTEM     | 
      FORMAT_MESSAGE_IGNORE_INSERTS,
      NULL,
      GetLastError(),
      MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
//    (LPTSTR) buf, // &lpMsgBuf,
               buf, // &lpMsgBuf,
      SIZE_BUF,
      0 )) { return "error with Win32_LastError"; }
  
//ret = (LPCTSTR) lpMsgBuf;
//::LocalFree(lpMsgBuf);
  return buf;
}
