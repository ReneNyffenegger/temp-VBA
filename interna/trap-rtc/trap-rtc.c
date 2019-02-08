#include "msg.h"
#include "funcsInDll.h"

HANDLE msgFile;

void callbackFuncInDll(char *funcName) {
  writeToFile(msgFile, funcName);
  writeToFile(msgFile, "\n");

}



__declspec(dllexport) void __stdcall Init () {

  msgFile = createNewOutputFile(TEXT("c:\\temp\\trap-rtc.msg"));

  writeToFile(msgFile, "Initialized\n");

  iterateOverFuncsInDll("VBE7.dll", "C:\\Program Files\\Common Files\\microsoft shared\\VBA\\VBA7", callbackFuncInDll);

}
