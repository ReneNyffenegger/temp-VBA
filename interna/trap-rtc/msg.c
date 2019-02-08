#include "msg.h"

HANDLE createNewOutputFile(LPCTSTR fileName) {

  HANDLE ret = CreateFileA(
       fileName                                       , // lpFileName,
       GENERIC_WRITE                                  , // dwDesiredAccess,
       FILE_SHARE_READ                                , // dwShareMode,
       NULL                                           , // lpSecurityAttributes,
       CREATE_ALWAYS                                  , // dwCreationDisposition,
       FILE_ATTRIBUTE_NORMAL | FILE_FLAG_WRITE_THROUGH, // dwFlagsAndAttributes,
       NULL                                             // hTemplateFile
  );

}

void writeToFile(HANDLE f, char *text) {

     DWORD nofBytesWritten;

     WriteFile(f, text, lstrlenA(text), &nofBytesWritten, NULL);

}
