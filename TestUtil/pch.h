// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file

#ifndef PCH_H
#define PCH_H

#define _CRT_SECURE_NO_WARNINGS

// TODO: add headers that you want to pre-compile here
#include <Windows.h>
#include <ObjSafe.h>
#include <iostream>
#include <vector>
#include <string>
#include <filesystem>
#include <tchar.h>

#include "MaxsUtil_h.h"

#ifdef _DEBUG
#define ASSERT(good) if (!(good)) DebugBreak();
#else
#define ASSERT(good)
#endif

extern HMODULE LibInstanceHandle;
extern ULONG LibRefCount;

#define LIB_ADDREF InterlockedIncrement(&LibRefCount)
#define LIB_RELEASE InterlockedDecrement(&LibRefCount)
#define LIB_LOCK InterlockedIncrement(&LibRefCount)
#define LIB_UNLOCK InterlockedDecrement(&LibRefCount)

#endif //PCH_H
