#include "msg.h"

static HANDLE msgFile = 0;

void createNewOutputFile(LPCTSTR fileName) {

  if (msgFile) {
    writeToFile("File already opened");
    return;
  }

  msgFile = CreateFileA(
       fileName                                       , // lpFileName,
       GENERIC_WRITE                                  , // dwDesiredAccess,
       FILE_SHARE_READ                                , // dwShareMode,
       NULL                                           , // lpSecurityAttributes,
       CREATE_ALWAYS                                  , // dwCreationDisposition,
       FILE_ATTRIBUTE_NORMAL | FILE_FLAG_WRITE_THROUGH, // dwFlagsAndAttributes,
       NULL                                             // hTemplateFile
  );

}

void writeToFile(const char *text) {

     DWORD nofBytesWritten;

     WriteFile(msgFile, text, lstrlenA(text), &nofBytesWritten, NULL);

}
