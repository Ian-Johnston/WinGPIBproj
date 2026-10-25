WinGPIB on Windows 7 - experimental, untested on Windows 7

WinGPIB's charts use SkiaSharp 3.119. Its native libSkiaSharp.dll lists two Windows 10 / DirectX functions in its import table
(D3D12SerializeRootSignature from d3d12.dll, D3DCompile from d3dcompiler_47.dll), so Windows 7 refuses to load it:
"The program can't start because d3d12.dll is missing". WinGPIB draws with Skia's CPU renderer and never calls either function.

These two 32-bit stand-in DLLs (WinGPIB is an x86 program) only export those two names, so the loader is satisfied:
  d3d12.dll
  d3dcompiler_47.dll
Copy BOTH into the same folder as WinGPIB.exe (and, if it still complains, also into its x86 subfolder, next to libSkiaSharp.dll).

Use them on Windows 7 ONLY. On Windows 10/11 they would hide the real DirectX DLLs from this program.
Windows 7 also needs Service Pack 1 and .NET Framework 4.8. Other Windows 7 problems may still appear after this one.
Source (C) and the .def files are included; built with Visual Studio 2022 for x86: cl /LD /MT /GS- file.c /link /DEF:file.def
