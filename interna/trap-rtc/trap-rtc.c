#include "msg.h"
#include "txt.h"
#include "funcsInDll.h"

HANDLE msgFile;

void callbackFuncInDll(char *funcName, DWORD address) {
  writeToFile(msgFile, txt("callbackFuncInDll received %s @ %d ?", funcName, address));
//writeToFile(msgFile, txt("callbackFuncInDll received %s @                       ));


  HANDLE h = GetModuleHandle("VBE7.dll");
  DWORD  a = GetProcAddress(h, funcName);
  writeToFile(msgFile, txt("proc addr: %d", a) );

  writeToFile(msgFile, "\n");

  

}



__declspec(dllexport) void __stdcall Init () {

  msgFile = createNewOutputFile(TEXT("c:\\temp\\trap-rtc.msg"));

  writeToFile(msgFile, "Initialized\n");

  iterateOverFuncsInDll("VBE7.dll", "C:\\Program Files\\Common Files\\microsoft shared\\VBA\\VBA7", callbackFuncInDll);

}

BOOL DllMain(HINSTANCE hInst, DWORD reason, LPVOID reserved) {
  return TRUE;
}
