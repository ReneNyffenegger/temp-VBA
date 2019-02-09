#include <windows.h>

typedef void (*fnFuncInDll)(char *funcName, DWORD address);

void iterateOverFuncsInDll(char *imageName, char *path, fnFuncInDll callback);
