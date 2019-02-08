#include <windows.h>
#include <Imagehlp.h>

#include "funcsInDll.h"
#include "error.h"


void iterateOverFuncsInDll(char *imageName, char *path, fnFuncInDll callback) {

  LOADED_IMAGE img;

  if (!MapAndLoad(imageName, path, &img, TRUE, TRUE)) MessageBox(0, Win32_LastError(), "MapAndLoad", 0);

}


