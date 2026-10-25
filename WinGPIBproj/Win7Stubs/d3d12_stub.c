#include <windows.h>
/* Stand-in for the Windows 10 d3d12.dll, for Windows 7 only. libSkiaSharp.dll lists D3D12SerializeRootSignature in its import table, so
   Windows 7 refuses to load it if d3d12.dll is missing. WinGPIB draws with Skia's CPU renderer and never calls it. */
HRESULT WINAPI D3D12SerializeRootSignature(const void* desc, int version, void** blob, void** error)
{
    if (blob) *blob = 0;
    if (error) *error = 0;
    return E_NOTIMPL;
}
BOOL WINAPI DllMain(HINSTANCE h, DWORD reason, LPVOID reserved) { return TRUE; }
