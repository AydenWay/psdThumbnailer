Note: This is a personal project that is mostly vibe coded. It has allowed psd thumbnails for me, but I cannot guarantee success.

# PSD Thumbnail Provider — Setup Guide

All commands require an **Admin command prompt or PowerShell**.

---

## Step 1 — Generate a GUID

Run this in PowerShell to generate a unique GUID for your extension:
```powershell
[guid]::NewGuid()
```

Copy the output and paste it into `PsdThumbnailProvider.cs`, replacing the ("YOUR-GUID-HERE"):

---

## Step 2 — Compile to DLL
```cmd
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe /target:library /platform:x64 /out:"C:\Windows\PsdThumbnailProvider.dll" /reference:"C:\Windows\assembly\GAC_MSIL\System.Drawing\4.0.0.0__b03f5f7f11d50a3a\System.Drawing.dll" "C:\path\to\PsdThumbnailProvider.cs"
```
> Replace `C:\path\to\` with the folder containing your `.cs` file.

---

## Step 3 — Register the DLL
```cmd
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe /codebase "C:\Windows\PsdThumbnailProvider.dll"
```

---

## Step 4 — Fix ThreadingModel
```cmd
reg add "HKCR\CLSID\{YOUR-GUID-HERE}\InprocServer32" /v "ThreadingModel" /t REG_SZ /d "Apartment" /f
```

---

## Step 5 — Add HKCU Registration (Explorer looks here first)
```cmd
reg add "HKCU\Software\Classes\CLSID\{YOUR-GUID-HERE}" /ve /d "PSD Thumbnail Provider" /f

reg add "HKCU\Software\Classes\CLSID\{YOUR-GUID-HERE}\InprocServer32" /ve /d "mscoree.dll" /f

reg add "HKCU\Software\Classes\CLSID\{YOUR-GUID-HERE}\InprocServer32" /v "CodeBase" /t REG_SZ /d "file:///C:/Windows/PsdThumbnailProvider.dll" /f

reg add "HKCU\Software\Classes\CLSID\{YOUR-GUID-HERE}\InprocServer32" /v "Class" /t REG_SZ /d "CustomShellExtensions.PsdThumbnailProvider" /f

reg add "HKCU\Software\Classes\CLSID\{YOUR-GUID-HERE}\InprocServer32" /v "Assembly" /t REG_SZ /d "PsdThumbnailProvider, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null" /f

reg add "HKCU\Software\Classes\CLSID\{YOUR-GUID-HERE}\InprocServer32" /v "RuntimeVersion" /t REG_SZ /d "v4.0.30319" /f

reg add "HKCU\Software\Classes\CLSID\{YOUR-GUID-HERE}\InprocServer32" /v "ThreadingModel" /t REG_SZ /d "Apartment" /f
```

---

## Step 6 — Disable Process Isolation
 Explorer requires DisableProcessIsolation to be set for .NET shell extensions — without it, it tries to run your DLL in an isolated surrogate process (dllhost.exe) which breaks .NET COM loading entirely.
```cmd
reg add "HKCU\Software\Classes\CLSID\{YOUR-GUID-HERE}" /v "DisableProcessIsolation" /t REG_DWORD /d 1 /f
```

---

## Step 7 — Register as Thumbnail Handler for .psd Files

First check what ProgID your .psd files use:

You can either right click a psd file and check file properties, in which it will be shown next to "Type of file:"

Or you can run the following command:
```cmd
reg query "HKCR\.psd" /ve
```

If it returns a ProgID (e.g. `Photoshop.Image.27`), use that below. If nothing is returned, use `.psd` as the key name.

```cmd
reg add "HKCR\Photoshop.Image.27\ShellEx\{e357fccd-a995-4576-b01f-234630154e96}" /ve /d "{YOUR-GUID-HERE}" /f

reg add "HKCR\.psd\ShellEx\{e357fccd-a995-4576-b01f-234630154e96}" /ve /d "{YOUR-GUID-HERE}" /f
```

---

## Step 8 — Add to Approved Shell Extensions

Make sure you change the "YOUR-GUID-HERE" - keep brackets and quotes
```cmd
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved" /v "{YOUR-GUID-HERE}" /t REG_SZ /d "PSD Thumbnail Provider" /f
```

---

## Step 9 — Clear Cache and Restart Explorer
```powershell
taskkill /f /im dllhost.exe
taskkill /f /im explorer.exe
Remove-Item "$env:LocalAppData\Microsoft\Windows\Explorer\thumbcache_*.db" -Force
start explorer.exe
```

---

## Updating to a New DLL Version

If the DLL is locked and needs to be replaced, recompile with a new filename (e.g. `PsdThumbnailProvider2.dll`) then run:
```cmd
reg add "HKCU\Software\Classes\CLSID\{YOUR-GUID-HERE}\InprocServer32" /v "CodeBase" /t REG_SZ /d "file:///C:/Windows/PsdThumbnailProvider2.dll" /f
reg add "HKCU\Software\Classes\CLSID\{YOUR-GUID-HERE}\InprocServer32" /v "Assembly" /t REG_SZ /d "PsdThumbnailProvider2, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null" /f
reg add "HKCR\CLSID\{YOUR-GUID-HERE}\InprocServer32" /v "CodeBase" /t REG_SZ /d "file:///C:/Windows/PsdThumbnailProvider2.dll" /f
reg add "HKCR\CLSID\{YOUR-GUID-HERE}\InprocServer32" /v "Assembly" /t REG_SZ /d "PsdThumbnailProvider2, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null" /f
```

This is because once the dll is initialized, Windows will refuse to change or delete it. Simply creating a new one and redirecting the registry pointers to it I found to be easier.

Then repeat Step 9 to restart Explorer.

If this doess't work, you can try going to Disk Cleanup, selecting the C drive, selecting thumbnails, and pressing OK. This will clear the thumbnail cache which may allow the thumbnail to be updated.
