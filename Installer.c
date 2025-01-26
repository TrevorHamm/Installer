/* EDS Installer: A simple installer that pulls down a zip file from a network 
 * location, extracts it and creates a shortcut to the extracted executable in 
 * the Start menu.
 *
 * Usage: Installer.exe <program_name>
 *
 * The zip file is extracted to the MyApps directory.
 * A shortcut is created in the Programs/MyApps Start menu.
 *
 * Dependencies:
 *      ZLib:    https://github.com/kiyolee/zlib-win-build.git
 *      LibZip:  https://github.com/kiyolee/libzip-win-build.git
 *
 * Application and dependencies are statically linked so that the executable
 * is a single file.
 *
 * Author: Trevor Hamm
 *
 * Actions:
 * - Get program name from commandline arguments
 * - Find newest zip file from network folder by that name
 * - Check / Install / Upgrade local installer   (STEP 1)
 * - Download zip to %localappdata%\MyApps       (STEP 2)
 * - Check/fail if program is currently running
 * - Uninstall current version (if exists)       (STEP 3)
 * - Unzip file                                  (STEP 4)
 * - Create shortcut                             (STEP 5)
 * - Run app on exit                             (STEP 6)
 *
 * TODO:
 * - Add DEBUG flag.
 * - Delete local copy of the application zip after extracting it.
 *      (This isn't working for some reason)
*/

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <direct.h>
#include <locale.h>
#include <objbase.h>
#include <objidl.h>
#include <pathcch.h>
#include <shlobj.h>
#include <shlobj_core.h>
#include <shlwapi.h>
#include <strsafe.h>
#include <sys/stat.h>
#include <tchar.h>
#include <time.h>
#include <tlhelp32.h>  
#include <zip.h>
#include <zipconf.h>
//#include <curl.h>         // Use this if we want to make web service calls

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "kernel32.lib")
#pragma comment(lib, "pathcch.lib")
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "libzip-static.lib")

#define IDC_LISTVIEW 101
#define IDC_EXIT_BUTTON 102
#define IDC_COPY_BUTTON 103
#define IDC_PROGRESS_BAR 104

const wchar_t* PROGRAMDIR = L"c:\\Dev\\Test\\";       // For testing
static BOOL DEBUG = FALSE;
static BOOL GOODTOLAUNCH = FALSE;

// Function declarations
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK ListViewProc(HWND hwnd, UINT uMsg, WPARAM wParam, 
        LPARAM lParam);

static HWND hwndProgressBar;
WNDPROC wpOrigListViewProc;
HWND hListView;
HWND hExitButton;
HWND hCopyButton;
wchar_t exeFileName[MAX_PATH] = { 0 };
int msgIndex = 0;

void AddControls(HWND hwnd);

#pragma pack(push, 1)
typedef struct {
    unsigned int signature;
    unsigned short version;
    unsigned short bit_flag;
    unsigned short compression;
    unsigned short mod_time;
    unsigned short mod_date;
    unsigned int crc32;
    unsigned int compressed_size;
    unsigned int uncompressed_size;
    unsigned short filename_length;
    unsigned short extra_length;
} local_file_header;
#pragma pack(pop)

typedef struct {
    HWND hwnd;
    wchar_t src[MAX_PATH];
    wchar_t dst[MAX_PATH];
} COPYFILEPARAMS;

//============================================================================
// Add an output line to the list view with columns for time, type and message
static void AddMessage(wchar_t* textType, wchar_t* text) {
    LVITEM lvi = { 0 };
    lvi.mask = LVIF_TEXT;
    lvi.iItem = msgIndex;
    msgIndex++;

    // -----------------------------------------------------------------------
    // Get the current time
    time_t t = time(NULL);
    struct tm* tmInfo = localtime(&t);
    wchar_t timeString[9]; // HH:MM:SS
    wcsftime(timeString, 9, L"%H:%M:%S", tmInfo);

    // Set the time in the first column
    lvi.iSubItem = 0; // First column
    lvi.pszText = timeString;
    ListView_InsertItem(hListView, &lvi);

    // -----------------------------------------------------------------------
    // Set the provided textType in the second column
    lvi.iSubItem = 1; // Second column
    lvi.pszText = textType;
    ListView_SetItem(hListView, &lvi);

    // -----------------------------------------------------------------------
    // Set the provided text in the third column
    lvi.iSubItem = 2; // Third column
    lvi.pszText = text;
    ListView_SetItem(hListView, &lvi);

    // Process pending messages to keep the UI responsive
    MSG msg;
    while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
}

//============================================================================
// Program Functions (not gui related)
//============================================================================

// Function to add a space before a capital letter following a lowercase letter
// Used to determine the name for the application shortcut.
static wchar_t* AddSpaces(const wchar_t* appName) {
    size_t len = wcslen(appName);
    size_t newLen = len;

    // First pass: determine the length of the new string
    for (size_t i = 1; i < len - 1; ++i) {
        if (iswlower(appName[i + 1]) && iswupper(appName[i])) {
            ++newLen;
        }
    }

    wchar_t* newName = (wchar_t*)malloc((newLen + 1) * sizeof(wchar_t));
    if (!newName) {
        AddMessage(L"ERROR", L"Memory allocation failed");
        return NULL;
    }

    // Second pass: build the new string
    size_t j = 0;
    for (size_t i = 0; i < len; ++i) {
        // If we are on a capital (and not the first letter) and the 
        // next letter is lowercase, add a space
        if (i > 0 && i < len && iswupper(appName[i]) && 
                iswlower(appName[i + 1])) {
            newName[j++] = L' ';
        }
        newName[j++] = appName[i];
    }
    newName[j] = L'\0';

    return newName;
}

//============================================================================
// Find how many slashes are in a path to count how many levels deep it is.

static int DirDepth(const wchar_t *path)
{
    const wchar_t* ptr = path;
    int count = 0;
    while ((ptr = wcsstr(ptr, L"\\")) != NULL) {
        ptr += 1;
        count++;
    }
    return count;
}

//============================================================================

static BOOL DirectoryExists(LPWSTR szPath) {
    DWORD dwAttrib = GetFileAttributes(szPath);
    return (dwAttrib != INVALID_FILE_ATTRIBUTES &&
        (dwAttrib & FILE_ATTRIBUTE_DIRECTORY));
}

//============================================================================

static BOOL FileExists(LPCTSTR szPath)
{
  DWORD dwAttrib = GetFileAttributes(szPath);

  return (dwAttrib != INVALID_FILE_ATTRIBUTES && 
         !(dwAttrib & FILE_ATTRIBUTE_DIRECTORY));
}

//============================================================================

// For testing progress bar
static void Delay(int seconds)
{
    int milliseconds = 1000 * seconds;
    clock_t startTime = clock();
    while (clock() < startTime + milliseconds)
        ;
}

//============================================================================

static void CreateDirectories(const wchar_t* path) {
    wchar_t* tempPath = _wcsdup(path); // Duplicate path to not modify original
    if (tempPath == NULL) {
        AddMessage(L"ERROR", L"Trying to create a directory with no name");
        return;
    }

    wchar_t* currentPos = tempPath;
    BOOL isDir = FALSE;

    // Remove the filename from the path
    wchar_t* lastSeparator = NULL;
    while (*currentPos) {
        if (*currentPos == L'\\' || *currentPos == L'/') {
            lastSeparator = currentPos;
        }
        currentPos++;
    }

    if (lastSeparator) {
        *lastSeparator = L'\0'; // Null-terminate string at the last separator
    }

    // Reset currentPos to the beginning of the string
    currentPos = tempPath;

    // Traverse the path and create directories
    while (*currentPos) {
        if (*currentPos == L'\\' || *currentPos == L'/') {
            *currentPos = L'\0';
            isDir = DirectoryExists(tempPath);
            if (!isDir) {
                // Create the directory if it doesn't exist
                if (!SUCCEEDED(_wmkdir(tempPath))) {
                    continue;
                }
            }
            *currentPos = L'\\'; // Restore the separator
        }
        currentPos++;
    }
    isDir = DirectoryExists(tempPath);
    if (!isDir) {
        // Create the final directory if it doesn't exist
        if (!SUCCEEDED(_wmkdir(tempPath))) {
            AddMessage(L"ERROR", L"Unable to create directory");
        }
    }
    free(tempPath); // Free the duplicated path
}

//============================================================================

static void DeleteDirectoryContents(const wchar_t* path) {
    WIN32_FIND_DATA findData;
    HANDLE hFind;
    wchar_t searchPath[MAX_PATH];
    wchar_t filePath[MAX_PATH];

    // Create the search path
    swprintf(searchPath, MAX_PATH, L"%s\\*", path);

    hFind = FindFirstFile(searchPath, &findData);
    if (hFind == INVALID_HANDLE_VALUE) {
        wchar_t msg[MAX_PATH + 50] = { 0 };
        wcscpy_s(msg, MAX_PATH + 50, L"Unable to open directory ");
        wcscat_s(msg, MAX_PATH + 50, path);
        AddMessage(L"Error", msg);
        return;
    }
    do {
        if (wcscmp(findData.cFileName, L".") != 0 && 
                wcscmp(findData.cFileName, L"..") != 0) {
            swprintf(filePath, MAX_PATH, L"%s\\%s", path, findData.cFileName);
            if (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
                // Recursively delete subdirectories
                if (wcslen(filePath) > 20 && DirDepth(filePath) > 2) {
                    DeleteDirectoryContents(filePath);
                    RemoveDirectory(filePath);
                } else {
                    wchar_t msg[MAX_PATH + 50] = { 0 };
                    wcscpy_s(msg, MAX_PATH + 50, 
                        L"Please report: Aborting attempt to delete from ");
                    wcscat_s(msg, MAX_PATH + 50, filePath);
                    AddMessage(L"ERROR", msg);
                    FindClose(hFind);
                    return;
                }
            }
            else {
                // Delete files
                if (wcslen(filePath) > 20 && DirDepth(filePath) > 2) {
                    DeleteFile(filePath);
                } else {
                    wchar_t msg[MAX_PATH + 50] = { 0 };
                    wcscpy_s(msg, MAX_PATH + 50, 
                        L"Please report: Aborting attempt to delete file ");
                    wcscat_s(msg, MAX_PATH + 50, filePath);
                    AddMessage(L"ERROR", msg);
                    FindClose(hFind);
                    return;
                }
            }
        }
    } while (FindNextFile(hFind, &findData) != 0);

    FindClose(hFind);
}

static void DeleteDirectory(const wchar_t* path) {
    DeleteDirectoryContents(path);
    RemoveDirectory(path);
}

//============================================================================

static BOOL IsFileNewer(const wchar_t *file1, const wchar_t *file2) {
    WIN32_FILE_ATTRIBUTE_DATA file1Info, file2Info;
    // Get file attributes for the first file
    if (!GetFileAttributesEx(file1, GetFileExInfoStandard, &file1Info)) {
        return FALSE;
    }
    // Get file attributes for the second file
    if (!GetFileAttributesEx(file2, GetFileExInfoStandard, &file2Info)) {
        return FALSE;
    }
    // Compare the file times
    if (CompareFileTime(&file1Info.ftLastWriteTime, 
            &file2Info.ftLastWriteTime) == -1) {
        return TRUE;  // file2 is newer
    } else {
        return FALSE; // file1 is newer or they are the same age
    }
}

//============================================================================

static int IsProcessRunning(const wchar_t* processName) {
    HANDLE hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnapshot == INVALID_HANDLE_VALUE) {
        return 0;
    }

    PROCESSENTRY32W pe;
    pe.dwSize = sizeof(PROCESSENTRY32W);

    if (Process32FirstW(hSnapshot, &pe)) {
        do {
            if (_wcsicmp(pe.szExeFile, processName) == 0) {
                CloseHandle(hSnapshot);
                return 1;
            }
        } while (Process32NextW(hSnapshot, &pe));
    }

    CloseHandle(hSnapshot);
    return 0;
}

//============================================================================

static int CheckIfRunning(wchar_t* zipPath) {
    if (DEBUG == TRUE)
        AddMessage(L"DEBUG", L"417 CheckIfRunning...");
    // Find the executable name in the zip file
    char zipfile_mb[256];
    size_t output_size;
    wcstombs_s(&output_size, zipfile_mb, sizeof(zipfile_mb), zipPath, 
            sizeof(zipfile_mb));
    int err;
    zip_t* zip = zip_open(zipfile_mb, 0, &err);
    if (!zip) {
        return -1;
    }
    zip_int64_t num_entries = zip_get_num_entries(zip, /*flags=*/0);
    const char* exeExt = ".exe";
    for (zip_int64_t i = 0; i < num_entries; ++i) {
        const char* name = zip_get_name(zip, i, /*flags=*/0);
        if (name == NULL) {
            continue;
        }
        // Find exe file
        char* ext;
        ext = strchr(name, '.');
        if (ext == NULL) {
            continue;
        }

        size_t output_size;
        wchar_t processNameW[MAX_PATH];
        if (strcmp(ext, exeExt) == 0) {
            // Check if it is currently running
            mbstowcs_s(&output_size, processNameW, 256, name, 256);
            if (IsProcessRunning(processNameW)) {
                return 1;
            }
            else {
                return 0;
            }
        }
    }
    return -1;
}

//============================================================================

static int ExecuteProgram(wchar_t* exePath)
{
    STARTUPINFO si;
    PROCESS_INFORMATION pi;

    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));

    wchar_t directory[MAX_PATH];
    wcscpy_s(directory, MAX_PATH, exePath);
    PathCchRemoveFileSpec(directory, MAX_PATH);

    // Start the child process. 
    if (!CreateProcess(NULL,   // No module name (use command line)
        exePath,     // Command line
        NULL,           // Process handle not inheritable
        NULL,           // Thread handle not inheritable
        FALSE,          // Set handle inheritance to FALSE
        0,              // No creation flags
        NULL,           // Use parent's environment block
        directory,      // Starting directory 
        &si,            // Pointer to STARTUPINFO structure
        &pi)            // Pointer to PROCESS_INFORMATION structure
        )
    {
        AddMessage(L"ERROR", L"Unable to run program");
        return -1;
    }
    return 0;
}

//============================================================================

// Returns the shortcut and the directory where the shortcut points to
static BOOL FindShortcut(LPCTSTR shortcutName, LPTSTR targetDir, 
        DWORD targetDirSize,
    LPTSTR shortcutPath, DWORD shortcutPathSize) {
    TCHAR startMenuPath[MAX_PATH];
    HRESULT hr = SHGetFolderPath(NULL, CSIDL_PROGRAMS, NULL, 0, startMenuPath);

    if (SUCCEEDED(hr)) {
        TCHAR searchPath[MAX_PATH];
        StringCchPrintf(searchPath, MAX_PATH, _T("%s\\%s.lnk"), startMenuPath,
            shortcutName);

        // Copy the full path of the shortcut to the output parameter
        StringCchCopy(shortcutPath, shortcutPathSize, searchPath);

        // Resolve the shortcut
        IShellLink* psl;
        IPersistFile* ppf;
        WIN32_FIND_DATA wfd;
        TCHAR resolvedPath[MAX_PATH];

        if (!SUCCEEDED(CoInitialize(NULL)))
			return FALSE;
        if (SUCCEEDED(CoCreateInstance(&CLSID_ShellLink, NULL,
                CLSCTX_INPROC_SERVER, &IID_IShellLink, (void**)&psl))) {
            if (SUCCEEDED(psl->lpVtbl->QueryInterface(psl, &IID_IPersistFile,
                    (void**)&ppf))) {
                if (SUCCEEDED(ppf->lpVtbl->Load(ppf, searchPath, STGM_READ))) {
                    if (SUCCEEDED(psl->lpVtbl->GetPath(psl, resolvedPath,
                        MAX_PATH, &wfd, SLGP_UNCPRIORITY))) {
                        PathRemoveFileSpec(resolvedPath);
                        StringCchCopy(targetDir, targetDirSize, resolvedPath);
                        ppf->lpVtbl->Release(ppf);
                        psl->lpVtbl->Release(psl);
                        CoUninitialize();
                        return TRUE;
                    } else {
                        if (DEBUG == TRUE) {
                            AddMessage(L"DEBUG", 
                                L"528 FindShortcut: Did not find shortcut: "); 
                            AddMessage(L"DEBUG", shortcutPath);
                        }
                    }
                    ppf->lpVtbl->Release(ppf);
                }
                psl->lpVtbl->Release(psl);
            } else {
                if (DEBUG == TRUE)
                    AddMessage(L"DEBUG", 
                            L"534 FindShortcut: Unable to query interface");
            }
        } else {
            if (DEBUG == TRUE)
                AddMessage(L"DEBUG", 
                        L"536 FindShortcut: Unable to create instance");
        }
        CoUninitialize();
    }
    return FALSE;
}

//============================================================================

static wchar_t* GetNewestFileInDir(const wchar_t* dirLoc, 
        const wchar_t* fileExtension) {
    WIN32_FIND_DATA findFileData;
    wchar_t searchLoc[MAX_PATH];
    wchar_t* returnPath = (wchar_t*)malloc(MAX_PATH * sizeof(wchar_t));
    wchar_t msg[MAX_PATH + 30] = { 0 };
    if (returnPath == NULL) {
        AddMessage(L"ERROR", L"Memory allocation failed");
        return NULL;
    }

    if (DEBUG == TRUE) {
        wcscpy_s(msg, MAX_PATH + 50, L"558 GetNewestFileInDir: Looking for newest file in ");
        wcscat_s(msg, MAX_PATH + 50, dirLoc);
        AddMessage(L"DEBUG", msg);
    }

    wcscpy_s(searchLoc, MAX_PATH, dirLoc);
    wcscat_s(searchLoc, MAX_PATH, fileExtension);
    HANDLE hFind = FindFirstFile(searchLoc, &findFileData);

    if (hFind == INVALID_HANDLE_VALUE) {
        free(returnPath);
        return NULL;
    }

    FILETIME latestTime = { 0 };
    wchar_t latestFile[MAX_PATH] = { 0 };

    do {
        if (!(findFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {
            if (CompareFileTime(&findFileData.ftLastWriteTime, &latestTime) > 0) {
                latestTime = findFileData.ftLastWriteTime;
                wcscpy_s(latestFile, MAX_PATH, findFileData.cFileName);
            }
        }
    } while (FindNextFile(hFind, &findFileData) != 0);

    FindClose(hFind);

    wcscpy_s(returnPath, MAX_PATH, dirLoc);
    wcscat_s(returnPath, MAX_PATH, L"\\");
    wcscat_s(returnPath, MAX_PATH, latestFile);

    if (DEBUG == TRUE) {
        wcscpy_s(msg, MAX_PATH + 50, L"591 GetNewestFileInDir: Found file: ");
        wcscat_s(msg, MAX_PATH + 50, latestFile);
        AddMessage(L"DEBUG", msg);
    }

    return returnPath;
}

//============================================================================

static DWORD WINAPI CopyFileWithProgress(LPVOID lpParam) {
    COPYFILEPARAMS* params = (COPYFILEPARAMS*)lpParam;
    HWND hwnd = params->hwnd;
    wchar_t* src = params->src;
    wchar_t* dst = params->dst;

    ShowWindow(hwndProgressBar, SW_SHOW);

    FILE* source = _wfopen(src, L"rb");
    if (!source) {
        AddMessage(L"ERROR", L"Cannot open source file");
        free(params);
        return -1;
    }

    FILE* destination = _wfopen(dst, L"wb");
    if (!destination) {
        fclose(source);
        wchar_t msg[MAX_PATH + 30] = { 0 };
        wcscpy_s(msg, MAX_PATH + 30, L"Cannot create destination file ");
        wcscat_s(msg, MAX_PATH + 30, dst);
        AddMessage(L"ERROR", msg);
        free(params);
        return -1;
    }

    fseek(source, 0, SEEK_END);
    long totalSize = ftell(source);
    fseek(source, 0, SEEK_SET);

    if (totalSize <= 0) {
        fclose(source);
        fclose(destination);
        AddMessage(L"ERROR", L"Source file is empty or unreadable");
        free(params);
        return -1;
    }

    wchar_t msg[MAX_PATH + 50] = L"Downloading zip file from server: ";
    wcscat_s(msg, MAX_PATH + 50, src);
    AddMessage(L"INFO", msg);

    SendMessage(hwndProgressBar, PBM_SETRANGE, 0, MAKELPARAM(0, 100));
    SendMessage(hwndProgressBar, PBM_SETPOS, 0, 0);

    wchar_t buffer[4096];
    size_t copiedSize = 0;
    size_t bytesRead;
    while ((bytesRead = fread(buffer, 1, sizeof(buffer), source)) > 0) {
        fwrite(buffer, 1, bytesRead, destination);
        copiedSize += bytesRead;
        int progress = (int)(((double)copiedSize / totalSize) * 100);
        SendMessage(hwndProgressBar, PBM_SETPOS, progress, 0);
        // For testing progress bar:
        //Delay(1);

        // Process pending messages to keep the UI responsive
        MSG msg;
        while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    fclose(source);
    fclose(destination);

    ShowWindow(hwndProgressBar, SW_HIDE);
    free(params);
    return 0;
    AddMessage(L"INFO", L"File copy completed");
}

//============================================================================

static DWORD ExtractZip(const wchar_t* zipfile, const wchar_t* outdir) {
    wchar_t msg[MAX_PATH + 30] = { 0 };
    char zipfile_mb[256];
    char outdir_mb[256];
    int err = 0;
    size_t output_size;

    StringCchPrintf(msg, MAX_PATH+30, L"Extracting files from %s", zipfile);
    AddMessage(L"INFO", msg);

    wcstombs_s(&output_size, zipfile_mb, 256, zipfile, 256);
    wcstombs_s(&output_size, outdir_mb, 256, outdir, 256);

    struct zip* z = zip_open(zipfile_mb, 0, &err);

    if (z == NULL) {
        zip_error_t ziperror;
        zip_error_init_with_code(&ziperror, err);
        AddMessage(L"ERROR", L"Failed to open ZIP file");
        zip_error_fini(&ziperror);
        return -1;
    }

    zip_uint64_t num_entries = zip_get_num_entries(z, 0);
    for (zip_uint64_t i = 0; i < num_entries; i++) {
        const char* name = zip_get_name(z, i, 0);
        if (name == NULL) {
            AddMessage(L"ERROR", L"Failed to get name for entry");
            continue;
        }

        struct zip_stat st;
        size_t output_size;
        if (zip_stat_index(z, i, 0, &st) == 0) {
            wchar_t wname[256];
            mbstowcs_s(&output_size, wname, 256, name, 256);

            wchar_t outpath[512];
            swprintf(outpath, sizeof(outpath) / sizeof(wchar_t), L"%ls/%ls",
                outdir, wname);

            CreateDirectories(outpath);
            if (wname[wcslen(wname) - 1] == L'/') {
                // This entry is a directory
                continue;
            }
            struct zip_file* zf = zip_fopen_index(z, i, 0);
            if (zf == NULL) {
                AddMessage(L"ERROR", L"Failed to open file in ZIP");
                continue;
            }

            FILE* outf = _wfopen(outpath, L"wb");
            if (outf == NULL) {
                zip_fclose(zf);
                wcscpy_s(msg, MAX_PATH + 30, L"Failed to open output file ");
                wcscat_s(msg, MAX_PATH + 30, outpath);
                AddMessage(L"ERROR", msg);
                continue;
            }

            char buffer[4096];
            zip_int64_t bytes_read;
            while ((bytes_read = zip_fread(zf, buffer, sizeof(buffer))) > 0) {
                fwrite(buffer, 1, bytes_read, outf);
            }

            fclose(outf); // Ensure the output file is closed
            zip_fclose(zf); // Ensure the zip file entry is closed
        }
        else {
            AddMessage(L"ERROR", L"Failed to get file information");
        }
    }

    if (zip_close(z) == 0 && DEBUG == TRUE) {
        AddMessage(L"DEBUG", L"753 ExtractZip: zip_close succeeded");
    }
    else {
        if (DEBUG == TRUE)
            AddMessage(L"DEBUG", L"755 ExtractZip: zip_close failed");
    }

    // Delete generally returns 0 which is a fail
    if (FileExists(zipfile)) {
        if (DeleteFile(zipfile) == 0 && DEBUG == TRUE) {
            swprintf(msg, MAX_PATH + 20, L"Unable to delete %s", zipfile);
            AddMessage(L"DEBUG", msg);
        }
    }
    return 0;
}

//============================================================================

static HRESULT CreateShortcut(LPCWSTR exePath, LPCWSTR cwdPath,
        LPCWSTR shortcutPath, LPCWSTR shortcutDescription) {

    HRESULT hres;
    IShellLink* psl = { 0 };
    /* DEBUGGING:
    wchar_t msg[MAX_PATH + 40] = { 0 };
    //StringCchCopy(msg, MAX_PATH+40, 
    StringCchCopy(msg, MAX_PATH+40, L"769 CreateShortcut: exePath: ");
    StringCchCat(msg, MAX_PATH+40, exePath);
    AddMessage(L"DEBUG", msg);
    StringCchCopy(msg, MAX_PATH+40, L"771 CreateShortcut: cwdPath: ");
    StringCchCat(msg, MAX_PATH+40, cwdPath);
    AddMessage(L"DEBUG", msg);
    StringCchCopy(msg, MAX_PATH+40, L"773 CreateShortcut: shortcutPath: ");
    StringCchCat(msg, MAX_PATH+40, shortcutPath);
    AddMessage(L"DEBUG", msg);
    StringCchCopy(msg, MAX_PATH+40, L"773 CreateShortcut: shortcutDescription: "); 
    StringCchCat(msg, MAX_PATH+40, shortcutDescription);
    AddMessage(L"DEBUG", msg);
    */

    // Get a pointer to the IShellLink interface
    hres = CoInitialize(NULL);
    if (SUCCEEDED(hres)) {
        hres = CoCreateInstance(&CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER,
            &IID_IShellLink, (LPVOID*)&psl);
        if (SUCCEEDED(hres)) {
            IPersistFile* ppf = { 0 };

            // Set the path to the shortcut target and add the description
            psl->lpVtbl->SetPath(psl, exePath);
            psl->lpVtbl->SetWorkingDirectory(psl, cwdPath);
            psl->lpVtbl->SetDescription(psl, shortcutDescription);

            // Query IShellLink for IPersistFile interface for saving shortcut
            hres = psl->lpVtbl->QueryInterface(psl, &IID_IPersistFile,
                (LPVOID*)&ppf);
            if (SUCCEEDED(hres)) {
                // Save the link by calling IPersistFile::Save
                hres = ppf->lpVtbl->Save(ppf, shortcutPath, TRUE);
                ppf->lpVtbl->Release(ppf);
            } else if (DEBUG == TRUE) {
                AddMessage(L"DEBUG", 
                        L"807 CreateShortcut: Failed while saving shortuct");
            }
            psl->lpVtbl->Release(psl);
        }
        CoUninitialize();
    }
    return hres;
}

//============================================================================

static int RegisterApp(wchar_t* executablePath_orig, wchar_t* folderPath, 
        wchar_t* appName) {

    wchar_t shortcutPath[MAX_PATH] = { 0 };
    wchar_t executablePath[MAX_PATH] = { 0 };
    wchar_t* shortcutName = AddSpaces(appName);

    AddMessage(L"INFO", L"Creating shortcut");
    wcscpy_s(executablePath, MAX_PATH, executablePath_orig);
    if (SHGetSpecialFolderPath(NULL, shortcutPath, CSIDL_PROGRAMS, TRUE)) {
        // Append the app name to the path
        wcscat_s(shortcutPath, MAX_PATH, L"\\MyApps\\");
        wcscat_s(shortcutPath, MAX_PATH, shortcutName);
        wcscat_s(shortcutPath, MAX_PATH, L".lnk");

        // Eg: Change name from "GroupManager" to "Group Manager"
        if (!shortcutName) {
            AddMessage(L"ERROR", L"Failed to get shortcut name");
            return -1;
        }

        // Create the shortcut
        if (!SUCCEEDED(CreateShortcut(executablePath, folderPath, shortcutPath,
                shortcutName))) {
            AddMessage(L"ERROR", L"Failed to create shortcut");
            return -1;
        }
    }
    else {
        AddMessage(L"ERROR", L"Failed to get the program directory");
        return -1;
    }

    return 0;
}

//============================================================================

static void UninstallApplication(wchar_t* appName, wchar_t* folderName) {
    // FolderName is MyOldApps for old installs, MyApps for new
    // Remove shortcut and existing folder
    wchar_t targetDir[MAX_PATH];
    wchar_t shortcutPath[MAX_PATH];
    size_t len = wcslen(appName) + 15;
    BOOL isDir = 0;

    if (DEBUG == TRUE)
        AddMessage(L"DEBUG", L"878 UninstallApplication...");
    wchar_t shortcutName[256] = { 0 };
	wcscpy_s(shortcutName, 256, folderName);
	wcscat_s(shortcutName, 256, L"\\");
	wcscat_s(shortcutName, 256, AddSpaces(appName));
    // This will find current apps in c:\MyOldApps
    if (FindShortcut(shortcutName, targetDir, MAX_PATH, shortcutPath, MAX_PATH)) {
        isDir = DirectoryExists(targetDir);
        // Expecting c:/Users/{username}/Appdata/Local/{ProgramName}  OR...
        //      c:/MyOldApps/{ProgramName}
        if (wcslen(targetDir) > 20 && isDir) {
            AddMessage(L"INFO", L"Deleting existing version...");
            DeleteDirectory(targetDir);
        }
		DeleteFile(shortcutPath);
        if (DEBUG == TRUE) {
            AddMessage(L"DEBUG", 
                    L"895 UninstallApplication: Deleted existing shortcut");
            AddMessage(L"DEBUG", shortcutPath);
        }
    } else if (DEBUG == TRUE) {
        wchar_t msg[MAX_PATH+30] = L"880 UninstallApplication: Did not find existing shortcut ";
        wcscat_s(msg, MAX_PATH+30, shortcutPath);
        AddMessage(L"DEBUG", msg);
    }
}

//============================================================================
// Get AppInstaller

static int GetInstaller(HWND hwnd, wchar_t* appdata) {

    AddMessage(L"INFO", L"Getting Installer...");
    int retval = 0;
    COPYFILEPARAMS* params = (COPYFILEPARAMS*)malloc(sizeof(
            COPYFILEPARAMS));
    if (params == NULL) {
        AddMessage(L"ERROR", L"Memory allocation failed");
        return -1;
    }

    // Copy AppInstaller and unzip it
    wchar_t srcDir[MAX_PATH] = { 0 };
    wcscpy_s(srcDir, MAX_PATH, PROGRAMDIR);
    wcscat_s(srcDir, MAX_PATH, L"AppInstaller2");
    wchar_t* remoteInstaller = GetNewestFileInDir(srcDir, L"\\*.zip");
    if (remoteInstaller == NULL) {
        AddMessage(L"ERROR", L"Couldn't find remote installer");
        return -1;
    }

    wchar_t localInstaller[MAX_PATH];
    wchar_t localInstallerDir[MAX_PATH];
    wcscpy_s(localInstaller, MAX_PATH, appdata);
    wcscpy_s(localInstallerDir, MAX_PATH, appdata);
    wcscat_s(localInstaller, MAX_PATH, L"\\MyApps\\AppInstaller.zip");
    wcscat_s(localInstallerDir, MAX_PATH, L"\\MyApps\\AppInstaller");
    params->hwnd = hwnd;
    wcscpy_s(params->src, MAX_PATH, remoteInstaller);
    wcscpy_s(params->dst, MAX_PATH, localInstaller);
    retval = CopyFileWithProgress(params);
    if (retval == 0) {
        if (ExtractZip(localInstaller, localInstallerDir) != 0) {
            AddMessage(L"ERROR", L"Couldn't extract installer");
            return -1;
        }
    } else {
        AddMessage(L"ERROR", L"Couldn't copy remote installer");
        return -1;
    }
    return retval;
}

//============================================================================
// Check / Install / Update AppInstaller locally

static int UpdateInstaller(HWND hwnd, wchar_t* appdata) {
    BOOL isDir = 0;
    int retval = 0;

    wchar_t srcDir[MAX_PATH] = { 0 };
    wcscpy_s(srcDir, MAX_PATH, PROGRAMDIR);
    wcscat_s(srcDir, MAX_PATH, L"\\AppInstaller2\\");

    wchar_t localDir[MAX_PATH];
    wcscpy_s(localDir, MAX_PATH, appdata);
    wcscat_s(localDir, MAX_PATH, L"\\MyApps\\");
    wcscat_s(localDir, MAX_PATH, L"\\AppInstaller\\");

    isDir = DirectoryExists(localDir);
    if (!isDir) {
        if (!SUCCEEDED(_wmkdir(localDir))) {
            AddMessage(L"ERROR", L"Unable to create 'AppInstaller' directory");
            return -1;
        } else {
            return GetInstaller(hwnd, appdata);
        }
    } else {
        // Is AppInstaller.exe older than the zip on the network?
        wchar_t* remoteInstaller = GetNewestFileInDir(srcDir, L"\\*.zip");
        wchar_t* localInstaller = GetNewestFileInDir(localDir, L"\\*.exe");
        if (localInstaller == NULL) {
            return GetInstaller(hwnd, appdata);
        } else {
            BOOL isNewer = IsFileNewer(localInstaller, remoteInstaller);
            if (isNewer) {
                // Rename local installer (can't delete it when it is running), 
                // Copy new zip and unzip it
                // FIXME: Delete old copies
                wchar_t newName[MAX_PATH] = { 0 };
                wchar_t nowStr[16];
                time_t now = time(0);
                struct tm *timeptr = localtime(&now);
                wcsftime(nowStr, sizeof(nowStr)-1, L"%m%d%Y_%H%M%S", timeptr);
                wcscpy_s(newName, MAX_PATH, localInstaller);
                wcscat_s(newName, MAX_PATH, L"_Old_");
                wcscat_s(newName, MAX_PATH, nowStr);
                AddMessage(L"INFO", newName);
                if (SUCCEEDED(_wrename(localInstaller, newName))) {
                    return GetInstaller(hwnd, appdata);
                }
            }
        }
    }
    return 0;
}

//============================================================================

static int ProcessInstall(HWND hwnd, wchar_t* appName) {
    // Get path to APPDATA local
    BOOL isDir = 0;
    int retval = 0;
    wchar_t appdata[MAX_PATH];
    wchar_t destFolderPath[MAX_PATH] = { 0 };
    wchar_t msg[75] = { 0 };
    wchar_t zipFilename[MAX_PATH] = { 0 };
    if (SUCCEEDED(SHGetFolderPath(NULL, CSIDL_LOCAL_APPDATA, NULL, 0, appdata)))
    {
        // Make sure Worley directory exists in %LocalAppData% 
        wchar_t localZipName[MAX_PATH];
        wcscpy_s(localZipName, MAX_PATH, appdata);
        wcscat_s(localZipName, MAX_PATH, L"\\Worley\\");
        isDir = DirectoryExists(localZipName);
        if (!isDir) {
            if (!SUCCEEDED(_wmkdir(localZipName))) {
                AddMessage(L"ERROR", 
                    L"Unable to create 'Worley' directory in LocalAppData");
                return -1;
            }
        }

        // STEP 1: Check / Install / Update Installer
        UpdateInstaller(hwnd, appdata);

        // -------------------------------------------------------------------
        // Get newest zip file from network program directory
        wchar_t searchPath[MAX_PATH] = { 0 };
        wcscpy_s(searchPath, MAX_PATH, PROGRAMDIR);
        if (wcslen(appName) > 0) {
            wcscat_s(searchPath, MAX_PATH, appName);
            isDir = DirectoryExists(searchPath);
            if (!isDir) {
                StringCchPrintf(msg, 75, 
                    L"Could not find an application folder with the name %s",
                    appName);
                AddMessage(L"ERROR", msg);
            } else {
                wcscpy_s(zipFilename, MAX_PATH,
                    GetNewestFileInDir(searchPath, L"\\*.zip"));
            }
        }
        if (wcslen(appName) < 1) {
            AddMessage(L"ERROR", L"No application specified to install");
		} else if (wcslen(zipFilename) < 1) {
            if (isDir)  // We found the directory but no zip files in it
                AddMessage(L"ERROR", L"No zip files found");
        } else {
            // ---------------------------------------------------------------
            // Copy file from server
            StringCchPrintf(msg, 75, L"Installing application %s", appName);
            AddMessage(L"INFO", msg);

            // Copy application zip from server
            wcscat_s(localZipName, MAX_PATH, appName);
            // We will make use of %LocalAppData%/MyApps/{appName}
            wcscpy_s(destFolderPath, MAX_PATH, localZipName);
            wcscat_s(localZipName, MAX_PATH, L".zip");
            // TODO: don't copy if timestamps match
            if (FileExists(localZipName) && DEBUG == TRUE) {
                AddMessage(L"DEBUG", L"1051 ProcessInstall: Local zip exists");
            } else if (DEBUG == TRUE) {
                AddMessage(L"DEBUG", L"1053 ProcessInstall: Local zip not found");
                AddMessage(L"DEBUG", localZipName);
            }
            COPYFILEPARAMS* params = (COPYFILEPARAMS*)malloc(sizeof(
                    COPYFILEPARAMS));
            if (params == NULL) {
                AddMessage(L"ERROR", L"Memory allocation failed");
                return -1;
            }
            params->hwnd = hwnd;
            wcscpy_s(params->src, MAX_PATH, zipFilename);
            wcscpy_s(params->dst, MAX_PATH, localZipName);
            // STEP 2: Copy file from server
            retval = CopyFileWithProgress(params);
            // ---------------------------------------------------------------
            // Is our program already runnning?
            if (retval == 0)
                retval = CheckIfRunning(localZipName);
            if (retval > 0)
                AddMessage(L"ERROR", 
                        L"CANNOT INSTALL: the program is already running!");
            // ---------------------------------------------------------------
            // Deregister existing version
            if (retval == 0) {
                wcscpy_s(searchPath, MAX_PATH, appdata);
                wcscat_s(searchPath, MAX_PATH, L"\\MyApps\\");
                wcscat_s(searchPath, MAX_PATH, appName);
                // STEP 3: Uninstall existing version
                // Also check if there is version in the MyOldApps directory
                UninstallApplication(appName, L"MyOldApps");
                UninstallApplication(appName, L"MyApps");
            }
            // ---------------------------------------------------------------
            // Extract new version
            if (retval == 0) {
                // STEP 4: Extract zip
                retval = ExtractZip(localZipName, destFolderPath);
            }
            // ---------------------------------------------------------------
            // Create shortcut
            if (retval == 0) {
                wchar_t* newestFile = GetNewestFileInDir(searchPath, L"\\*.exe");
				if (newestFile != NULL && wcslen(newestFile) > 0) {
					wcscpy_s(exeFileName, MAX_PATH, newestFile);
                    if (DEBUG == TRUE) {
                        AddMessage(L"DEBUG", 
                                L"1092 ProcessInstall: New unzipped executable:");
                        AddMessage(L"DEBUG", newestFile);
                    }
                } else {
                    AddMessage(L"ERROR", 
                            L"1096 ProcessInstall: Did not find unzipped executable");
                    if (DEBUG == TRUE)
                        AddMessage(L"DEBUG", searchPath);
                }
            }
            if (exeFileName != NULL && wcslen(exeFileName) > 0) {
                // STEP 5: Create shortcut
                retval = RegisterApp(exeFileName, destFolderPath, appName);
            }

            if (retval == 0)
                GOODTOLAUNCH = TRUE;
            AddMessage(L"INFO", L"Finished!");
        }
	} else {
		AddMessage(L"ERROR", L"Could not get LocalAppData directory");
        return 0;
	}
    return 0;
}

//============================================================================
// GUI Functions
//============================================================================

static void DrawColumnLines(HWND hwnd) {
    RECT rc;
    GetClientRect(hwnd, &rc);

    HDC hdc = GetDC(hwnd);
    // Light gray color with 1 pixel thickness
    HPEN hPen = CreatePen(PS_SOLID, 1, RGB(192, 192, 192));
    HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);

    HWND hHeader = ListView_GetHeader(hwnd);
    int columnCount = Header_GetItemCount(hHeader);
    RECT columnRect = { 0 };

    for (int i = 0; i < columnCount; i++) {
        ListView_GetSubItemRect(hwnd, 0, i, LVIR_BOUNDS, &columnRect);
        if (i > 0) { // Ensure we draw lines after every column
            MoveToEx(hdc, columnRect.left - 1, rc.top, NULL);
            LineTo(hdc, columnRect.left - 1, rc.bottom);
        }
        MoveToEx(hdc, columnRect.right - 1, rc.top, NULL);
        LineTo(hdc, columnRect.right - 1, rc.bottom);
    }

    SelectObject(hdc, hOldPen);
    DeleteObject(hPen);
    ReleaseDC(hwnd, hdc);
}

//============================================================================

static void CopyListViewToClipboard() {
    int rowCount = ListView_GetItemCount(hListView);
    int colCount = Header_GetItemCount(ListView_GetHeader(hListView));

    // Calculate the buffer size needed
    int bufferSize = (rowCount * (256 + 30)) * colCount;
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, bufferSize);
    if (!hMem) return;

    // Lock the global memory object
    wchar_t* buffer = (wchar_t*)GlobalLock(hMem);
    if (!buffer) {
        GlobalFree(hMem);
        return;
    }

    buffer[0] = L'\0';

    // Copy the ListView content to the buffer
    for (int i = 0; i < rowCount; i++) {
        for (int j = 0; j < colCount; j++) {
            wchar_t text[256] = { 0 };
            int x = sizeof(text) / sizeof(wchar_t);
            ListView_GetItemText(hListView, i, j, text, x);
            wcscat_s(buffer, bufferSize / sizeof(wchar_t), text);
            if (j < colCount - 1) wcscat_s(buffer, bufferSize /
                sizeof(wchar_t), L"\t");
        }
        wcscat_s(buffer, bufferSize / sizeof(wchar_t), L"\r\n");
    }

    // Unlock the global memory object
    GlobalUnlock(hMem);

    // Open the clipboard and copy the buffer
    if (OpenClipboard(NULL)) {
        EmptyClipboard();
        SetClipboardData(CF_UNICODETEXT, hMem);
        CloseClipboard();
    }
    else {
        GlobalFree(hMem);
    }
}

//============================================================================

LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_CREATE:
        AddControls(hwnd);
        break;
    case WM_SIZE:
        if (hListView != NULL && hExitButton != NULL && hCopyButton != NULL) {
            int windowWidth = LOWORD(lParam);
            int windowHeight = HIWORD(lParam);
            int buttonWidth = 100;
            int buttonHeight = 30;
            int buttonSpacing = 10;
            int totalButtonWidth = (buttonWidth * 2) + buttonSpacing;

            // Resize the ListView
            ListView_SetColumnWidth(hListView, 2, LVSCW_AUTOSIZE_USEHEADER);
            MoveWindow(hListView, 7, 10, windowWidth - 15, windowHeight - 60, 
                    TRUE);

            // Calculate new positions for the buttons
            int xStart = (windowWidth - totalButtonWidth) / 2;
            int yPosition = windowHeight - 40;

            // Move the buttons to the new positions
            MoveWindow(hExitButton, xStart, yPosition, buttonWidth, 
                    buttonHeight, TRUE);
            MoveWindow(hCopyButton, xStart + buttonWidth + buttonSpacing, 
                    yPosition, buttonWidth + 50, buttonHeight, TRUE);

            // Force a redraw of the window
            InvalidateRect(hwnd, NULL, TRUE);
            UpdateWindow(hwnd);
        }
        break;
    // Without this, window resize attempt would not catch
    case WM_NCHITTEST: {
        LRESULT hit = DefWindowProc(hwnd, uMsg, wParam, lParam);
        if (hit == HTCLIENT) {
            POINT pt;
            pt.x = GET_X_LPARAM(lParam);
            pt.y = GET_Y_LPARAM(lParam);
            ScreenToClient(hwnd, &pt);
            RECT rc;
            GetClientRect(hwnd, &rc);
            if (pt.x >= rc.right - 10 && pt.y >= rc.bottom - 10) {
                return HTBOTTOMRIGHT;
            }
            if (pt.x <= 10 && pt.y >= rc.bottom - 10) {
                return HTBOTTOMLEFT;
            }
            if (pt.x >= rc.right - 10 && pt.y <= 10) {
                return HTTOPRIGHT;
            }
            if (pt.x <= 10 && pt.y <= 10) {
                return HTTOPLEFT;
            }
            if (pt.x >= rc.right - 10) {
                return HTRIGHT;
            }
            if (pt.x <= 10) {
                return HTLEFT;
            }
            if (pt.y >= rc.bottom - 10) {
                return HTBOTTOM;
            }
            if (pt.y <= 10) {
                return HTTOP;
            }
        }
        return hit;
    }
    case WM_COMMAND:
        if (LOWORD(wParam) == IDC_EXIT_BUTTON) {
            // STEP 6: Start new application on Exit
            if (GOODTOLAUNCH == TRUE)
                ExecuteProgram(exeFileName);
            PostQuitMessage(0);
        }
        else if (LOWORD(wParam) == IDC_COPY_BUTTON) {
            CopyListViewToClipboard(hListView);
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    default:
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
    return 0;
}

//============================================================================

void AddControls(HWND hwnd) {
    InitCommonControls();

    hListView = CreateWindowExW(
        WS_EX_CLIENTEDGE, WC_LISTVIEWW, NULL,
        WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_AUTOARRANGE,
        10, 10, 768, 350,
        hwnd, (HMENU)IDC_LISTVIEW, GetModuleHandle(NULL), NULL
    );

    // Subclass the ListView
    wpOrigListViewProc = (WNDPROC)SetWindowLongPtr(hListView, GWLP_WNDPROC,
        (LONG_PTR)ListViewProc);

    // Add columns
    LVCOLUMN lvCol = { 0 };
    lvCol.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
    lvCol.pszText = L"Time";
    lvCol.cx = 70;
    ListView_InsertColumn(hListView, 0, &lvCol);

    lvCol.pszText = L"Type";
    lvCol.cx = 70;
    ListView_InsertColumn(hListView, 1, &lvCol);

    lvCol.pszText = L"Message";
    lvCol.cx = 70;
    ListView_InsertColumn(hListView, 2, &lvCol);
    ListView_SetColumnWidth(hListView, 2, LVSCW_AUTOSIZE_USEHEADER);
    wchar_t* itemText = { 0 };

    hExitButton = CreateWindowExW(
        0, L"BUTTON", L"Close",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        100, 260, 100, 30,
        hwnd, (HMENU)IDC_EXIT_BUTTON, GetModuleHandle(NULL), NULL
    );

    hCopyButton = CreateWindowExW(
        0, L"BUTTON", L"Copy to Clipboard",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        210, 260, 150, 30,
        hwnd, (HMENU)IDC_COPY_BUTTON, GetModuleHandle(NULL), NULL
    );

    hwndProgressBar = CreateWindowEx(
        0, PROGRESS_CLASS, NULL,
        WS_CHILD | WS_VISIBLE | PBS_SMOOTH,
        50, 150, 300, 30,
        hwnd, (HMENU)IDC_PROGRESS_BAR, GetModuleHandle(NULL), NULL);

    SendMessage(hwndProgressBar, PBM_SETRANGE, 0, MAKELPARAM(0, 100));
    ShowWindow(hwndProgressBar, SW_HIDE);
    EnableWindow(hExitButton, FALSE);
}

//============================================================================

LRESULT CALLBACK ListViewProc(HWND hwnd, UINT uMsg, WPARAM wParam, 
        LPARAM lParam) {
    switch (uMsg) {
    case WM_PAINT:
        CallWindowProc(wpOrigListViewProc, hwnd, uMsg, wParam, lParam);
        DrawColumnLines(hwnd);
        return 0;
    }
    return CallWindowProc(wpOrigListViewProc, hwnd, uMsg, wParam, lParam);
}

//============================================================================
// Entry point

int WINAPI wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, 
        _In_ LPWSTR lpCmdLine, _In_ int nCmdShow) {

    const wchar_t CLASS_NAME[] = L"Sample Window Class";
    HANDLE hTh = { 0 };
    WNDCLASS wc = { 0 };
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hbrBackground = CreateSolidBrush(RGB(240, 240, 240));
    setlocale(LC_ALL, "");

    RegisterClass(&wc);

    /*
    // Parse command line arguments
    int argc;
    wchar_t* appName = { 0 };
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (argv != NULL) {
        for (int i = 0; i < argc; i++) {
            if (wcscmp(argv[i], L"--debug") == 0) {
                DEBUG = TRUE;
            } else if (i > 0 && DEBUG == FALSE) {
                appName = argv[i];
            }
        }
    }
    // Free the memory allocated for CommandLineToArgvW
    LocalFree(argv);
    */

    HWND hwnd = CreateWindowExW(
        0,                              // Optional window styles.
        CLASS_NAME,                     // Window class
        L"EDS Edmonton App Installer",  // Window text
        WS_OVERLAPPEDWINDOW,            // Window style
        CW_USEDEFAULT, CW_USEDEFAULT, 800, 500, // Size and position
        NULL,                           // Parent window
        NULL,                           // Menu
        hInstance,                      // Instance handle
        NULL                            // Additional application data
    );

    if (hwnd == NULL) {
        return 0;
    }

    HICON hIcon = LoadIcon(GetModuleHandle(L"shell32.dll"), MAKEINTRESOURCE(13));
    if (hIcon) {
        SendMessage(hwnd, WM_SETICON, ICON_BIG, (LPARAM)hIcon);
        SendMessage(hwnd, WM_SETICON, ICON_SMALL, (LPARAM)hIcon);
    }

    ShowWindow(hwnd, nCmdShow);

    // Here is where the magic happens:
    ProcessInstall(hwnd, lpCmdLine);

    EnableWindow(hExitButton, TRUE);

    MSG msg = { 0 };
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return EXIT_SUCCESS;
}

