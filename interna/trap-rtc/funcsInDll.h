#include <windows.h>

typedef void (*fnFuncInDll)(char *funcName);

void iterateOverFuncsInDll(char *imageName, char *path, fnFuncInDll callback);
