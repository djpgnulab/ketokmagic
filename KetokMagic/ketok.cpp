#include "ketokmagic.h"

HMODULE myDll = NULL;
HWND hDlg = NULL;
wstring esptPath;

uncompressPtr uncompress = NULL;
compressPtr compress = NULL;

const char * msgCapicomError = "CAPICOM Error !!!!";
const char * msgZlibError = "zlib Error !!!!";
const char * msgUnicodeConversionError = "Unicode Conversion Error !!!!";
const char * titel = "Ketok Magic V 0.0.4";
const char * msgMemoryError = "Alokasi Memori Erro !!!!";
const char * msgBantuan =
"1. Hanya untuk internal DJP\n\n"
"2. Untuk memperbaiki file csv e-SPT PPh Badan (...F1132140111.csv) yang memiliki end tag (;-->) tidak sebaris dengan open tag (<--;)\n\n"
"3. Komputer sudah terinstall capicom dan memiliki zlib.dll (di c:\\windows\\system32 untuk windows 32 bit atau c:\\windows\\syswow64 untuk windows 64 bit)\n\n"
"4. Hasil perbaikan disimpan dengan nama yang sama dengan file yang diperbaiki, sedangkan file yang diperbaiki di-rename menjadi ...F1132140111_backup.csv\n\n"
"Dibuat oleh gatot.sukmono@pajak.go.id";


string msgSistemError = "System Error: ";
char errbuffer[256];

wstring ansitowc(char * str, size_t len) {
	if (len == 0) return wstring();
	int size_needed = MultiByteToWideChar(CP_ACP, 0, str, len, NULL, 0);
	wstring strTo(size_needed, 0);
	MultiByteToWideChar(CP_ACP, 0, str, len, &strTo[0], size_needed);
	return strTo;
}

string wctoansi(wchar_t * wstr, size_t len)
{
	if (len == 0) return string();
	int size_needed = WideCharToMultiByte(CP_ACP, 0, wstr, len, NULL, 0, NULL, NULL);
	string strTo(size_needed, 0);
	WideCharToMultiByte(CP_ACP, 0, wstr, len, &strTo[0], size_needed, NULL, NULL);
	return strTo;
}

void writeFileEspt(wstring path, _bstr_t konten) {
	ofstream output;
	output.open(path, ios::out | ios::binary);
	output.write((char *)konten, konten.length());
	output.close();
}

_bstr_t readFileEspt(wchar_t * const path) 
{
	_bstr_t bstr_konten;
	 wifstream input;
		input.open(path);
		wstring konten;
		if (input.is_open()) {
			input.seekg(0, input.end);
			konten.resize(input.tellg(), sizeof(wchar_t));
			input.seekg(0, input.beg);
			input.read(&konten[0], konten.size());
			 bstr_konten = konten.c_str();
			input.close();
			return(bstr_konten);
		}
	
		strerror_s(errbuffer,255,errno);
		string msg(errbuffer);
		msg.insert(0, msgSistemError.c_str());
		
		throw(runtime_error(msg));
		
}


bool validasiFile(wstring &filename) {
	wstring tmpfilename = wstring(filename.c_str());
	if (tmpfilename.length() != 40) return false;
	tmpfilename.push_back(0);
	_wcslwr_s((wchar_t *)tmpfilename.c_str(), tmpfilename.size());

	wstring ext = tmpfilename.substr(37).c_str();
	if (ext.compare(L"csv") != 0 )  return false;
	
	wstring kodefile = tmpfilename.substr(25, 11);

	if (kodefile.compare(L"f1132140111") != 0 && kodefile.compare(L"f1132150111") != 0)  return false;
	
	return true;
}

void betulkan()  {
	try {
		if (esptPath.empty()) return;

		wstring filename(PathFindFileName(esptPath.c_str()));

		if (validasiFile(filename) == false) {
			string msg = "File ";
			msg += wctoansi(&filename[0], filename.size()) + " tidak dapat diproses";
			throw(runtime_error(msg.c_str()));
		}

		wstring directory = esptPath;
		int pos = directory.find(filename, 0);
		directory.erase(pos);

		_bstr_t password = createPassword(filename);
		_bstr_t encryptedEspt = readFileEspt((wchar_t *)esptPath.c_str());
		string decr = decrypt(encryptedEspt, password);

		istringstream espt(decrypt(encryptedEspt, password));

		regex endtag(".+?-->;\\s*$");
		string newbaris = "";
		list<string> listbaris;

		int bad_baris = 0;
		int good_baris = 0;
		int nbaris = 0;

		for (string baris; getline(espt, baris);) {

			if (*baris.rbegin() == '\r') {
				baris.pop_back();
			}

			newbaris.append(baris);
			if (true == regex_match(newbaris, endtag)) {
				if (nbaris) {
					bad_baris += (nbaris + 1);
					good_baris++;
					nbaris = 0;
				}
				listbaris.push_back(newbaris);
				newbaris = "";
			}
			else {
				nbaris++;
			}
		}

		if (good_baris) {
			newbaris = "";
			for (list<string>::iterator it = listbaris.begin(); it != listbaris.end(); ++it) {
				newbaris += (*it + "\r\n");
			}

			_bstr_t encrypted = encrypt(newbaris, password);

			wstring backupFilename(filename.c_str());
			int dotpos = backupFilename.find(L'.');
			dotpos = dotpos == string::npos ? backupFilename.size() : dotpos--;
			backupFilename.insert(dotpos, L"_backup");

			_wrename((directory + filename).c_str(), (directory + backupFilename).c_str());
			writeFileEspt(directory + filename, encrypted);
			string msg(201, 0);
			sprintf_s(&msg[0], 200, "Memperbaiki %d baris menjadi %d baris\n%s", bad_baris, good_baris, "Semoga Sukses :)");

			MessageBoxA(NULL, msg.c_str() , titel, MB_OK | MB_ICONINFORMATION);
		}
		else {
			MessageBoxA(NULL, "File tidak bermasalah :)", titel, MB_OK | MB_ICONINFORMATION);
		}
	}
	catch (exception& e) {
		MessageBoxA(NULL, e.what(), titel, MB_OK|MB_ICONERROR);
	}
	
}

void openfile() {

	OPENFILENAME ofn = { 0 };
	ofn.lStructSize = sizeof(ofn);
	esptPath.clear();
	esptPath.resize(1025, 0);
	ofn.lpstrFile = &esptPath[0];
	ofn.lpstrFilter = L"File CSV(*.csv)\0*.csv\0";
	ofn.nMaxFile = 1024;
	if (TRUE == GetOpenFileName(&ofn)) {
			SetDlgItemTextW(hDlg, IDC_EDIT2, esptPath.c_str());
	}
	else {
		esptPath.clear();
	}
}

void bantuan() { 
	MessageBoxA(NULL, msgBantuan, titel, MB_OK | MB_ICONINFORMATION);
}


INT_PTR CALLBACK DialogProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDCANCEL:
			SendMessage(hDlg, WM_CLOSE, 0, 0);
			return TRUE;
		case IDBROWSE: 
			openfile();
			return TRUE;
		case IDC_PROSES:
			betulkan();
			return TRUE;
		case IDC_BANTUAN:
			bantuan();
			return TRUE;

		}
		break;

	case WM_CLOSE:
		DestroyWindow(hDlg);
		return TRUE;

	case WM_DESTROY:
		if (myDll != NULL) FreeLibrary(myDll);
		PostQuitMessage(0);
		return TRUE;
	}

	return FALSE;
}

int WINAPI _tWinMain(HINSTANCE hInst, HINSTANCE h0, LPTSTR lpCmdLine, int nCmdShow)
{

	MSG msg;
	BOOL ret;

	HMODULE myDll = LoadLibraryA("zlib.dll");
	if (myDll != NULL) {
		uncompress = (uncompressPtr)GetProcAddress(myDll, "uncompress");
		compress = (compressPtr)GetProcAddress(myDll, "compress");

		if (uncompress == NULL || compress == NULL) {
			MessageBoxA(NULL, "Zlib bermasalah", titel, MB_OK | MB_ICONINFORMATION);
			return -1;
		} 
	}
	else {
		MessageBoxA(NULL, "Tidak ada Zlib", titel, MB_OK | MB_ICONINFORMATION);
		return -1;

	}

	CoInitialize(0);
	InitCommonControls();
	hDlg = CreateDialogParam(hInst, MAKEINTRESOURCE(IDD_DIALOG1), 0, DialogProc, 0);
	ShowWindow(hDlg, nCmdShow);

	while ((ret = GetMessage(&msg, 0, 0, 0)) != 0) {
		if (ret == -1)
			return -1;

		if (!IsDialogMessage(hDlg, &msg)) {
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	CoUninitialize();

	return 0;
}