#include <windows.h>

HANDLE createNewOutputFile(LPCTSTR fileName);

void writeToFile(HANDLE f, char *text);
