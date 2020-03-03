# Laserfiche_PageGenerator
This Application is a console app that generates Laserfiche pages (TIFFs) using the repository Id of the entry.

## Instructions :

1. Download Ghostscript
   * for Windows (32bit) : https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/download/gs950/gs950w32.exe
   * place the gsdll32.dll file into the project directory.
   * **File Properties :**
   * Choose Build Action : Embedded Resource
   * Copy to Output Directory : Copy always
2. Create Project in Visual Studio Enterprise 2017/2019.
3. Install Laserfiche SDK. *No link provided*
4. Add references to
   * Laserfiche.DocumentServices *10.2.0.0*
   * Laserfiche.RepositoryAccess *10.2.0.0*
5. Clean and build 32bit console application.
