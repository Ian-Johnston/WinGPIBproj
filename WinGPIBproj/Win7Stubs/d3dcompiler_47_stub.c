#include <windows.h>
/* Stand-in for d3dcompiler_47.dll, for Windows 7 machines that don't have it (KB4019990). libSkiaSharp.dll imports D3DCompile but WinGPIB
   draws with Skia's CPU renderer and never calls it. */
HRESULT WINAPI D3DCompile(LPCVOID src, SIZE_T size, LPCSTR name, const void* defines, void* include, LPCSTR entry,
                          LPCSTR target, UINT flags1, UINT flags2, void** code, void** errors)
{
    if (code) *code = 0;
    if (errors) *errors = 0;
    return E_FAIL;
}
BOOL WINAPI DllMain(HINSTANCE h, DWORD reason, LPVOID reserved) { return TRUE; }
