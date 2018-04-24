#pragma once
#include <Windows.h>
#include <CommCtrl.h>
#include <fstream>
#include <sstream>
#include <string>
#include <iostream>
#include <regex>
#include <list>
#include <cstdio>
#include <cerrno>
#include <cctype>
#include "resource.h"
#include "atlbase.h"  
#include "comutil.h"
#import "capicom.dll"

#pragma comment(linker, \
  "\"/manifestdependency:type='Win32' "\
  "name='Microsoft.Windows.Common-Controls' "\
  "version='6.0.0.0' "\
  "processorArchitecture='*' "\
  "publicKeyToken='6595b64144ccf1df' "\
  "language='*'\"")

#pragma comment(lib, "ComCtl32.lib")

#define Z_OK   0

using namespace CAPICOM;
using namespace std;

_bstr_t encrypt(string decrypted, _bstr_t password);
_bstr_t createPassword(wstring filename);
string decrypt(_bstr_t encrypted, _bstr_t password);

typedef int(__stdcall *uncompressPtr)(unsigned char *, unsigned long *, unsigned char *, unsigned long);
typedef int(__stdcall *compressPtr)(unsigned char *, unsigned long *, unsigned char *, unsigned long);

wstring ansitowc(char * str, size_t len);
string wctoansi(wchar_t * wstr, size_t len);
