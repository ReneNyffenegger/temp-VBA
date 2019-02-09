//
// TODO: https://stackoverflow.com/a/1128453/180275
//
#include <windows.h>
#include <Imagehlp.h>

#include "funcsInDll.h"
#include "error.h"
#include "txt.h"


void iterateOverFuncsInDll(char *imageName, char *path, fnFuncInDll callback) {

  LOADED_IMAGE             img;
  PIMAGE_NT_HEADERS        peHeader;
//PIMAGE_DATA_DIRECTORY    exportTableOffset_rva;
  PIMAGE_EXPORT_DIRECTORY  exportDirectory;
//PIMAGE_EXPORT_DIRECTORY  ImageExportDirectory;
  char                   **functionNames;
  DWORD                   *functionAddresses;
  int                      funcNo;


  if (!MapAndLoad(imageName, path, &img, TRUE, TRUE)) {
     MessageBoxA(0, Win32_LastError(), "MapAndLoad", 0);
     return;
  }

//callback(txt("MapAndLoad ok, &img = %d, num = %d", &img, 42));

  peHeader = img.FileHeader;

//callback(txt("peHeader = %d", peHeader));
//callback(txt("img.MappedAddress = %d", img.MappedAddress));

//exportTableOffset_rva = & peHeader->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT];


  unsigned long cDirSize;
//PIMAGE_EXPORT_DIRECTORY  ImageExportDirectory = ImageDirectoryEntryToData(img.MappedAddress, FALSE, IMAGE_DIRECTORY_ENTRY_EXPORT, &cDirSize); 
                           exportDirectory      = ImageDirectoryEntryToData(img.MappedAddress, FALSE, IMAGE_DIRECTORY_ENTRY_EXPORT, &cDirSize); 
//callback(txt("ImageExportDirectory = %d", ImageExportDirectory));


//  callback(txt("exportTableOffset_rva = %d", exportTableOffset_rva));
//
//  if (! exportTableOffset_rva) { 
//    MessageBoxA(0, "! exportTableOffset_rva", 0, 0); 
//    UnMapAndLoad(&img);
//  }

//exportDirectory   = ImageRvaToVa(peHeader, img.MappedAddress, (ULONG) exportTableOffset_rva              , 0);
//callback(txt("exportDirectory = %d", exportDirectory));


//callback(txt("exportDirectory->NumberOfNames = %d", exportDirectory->NumberOfNames), 0);

  functionNames     = ImageRvaToVa(peHeader, img.MappedAddress, (ULONG) exportDirectory->AddressOfNames    , 0);
//callback(txt("functionNames = %d", functionNames));
  functionAddresses = ImageRvaToVa(peHeader, img.MappedAddress, (ULONG) exportDirectory->AddressOfFunctions, 0);

//callback(txt("functionAddresses = %d", functionAddresses));
  
  for (funcNo = 0; funcNo < exportDirectory->NumberOfNames; funcNo++) {
//   callback(txt("funcNo = %d", funcNo));
     callback(
           (char*) ImageRvaToVa(peHeader, img.MappedAddress, (ULONG) functionNames    [funcNo], 0), // functionNames[funcNo],
           (DWORD) ImageRvaToVa(peHeader, img.MappedAddress, (ULONG) functionAddresses[funcNo], 0)
     );

  }

  UnMapAndLoad(&img);

}
