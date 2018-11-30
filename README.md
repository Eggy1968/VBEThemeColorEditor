# VBEThemeColorEditor
The VBA / VB editor (or VBE) is limited in the 16 colors it can use to render code text. This application replaces the default colors with custom palettes that can be saved and loaded.

This application replaces the default colors with custom palettes that can be saved and loaded (2 themes are already provided).
The application can only run in administrator mode!

Download: https://github.com/gallaux/VBEThemeColorEditor/raw/master/VBEThemeColorEditor.zip

![alt tag](https://github.com/gallaux/VBEThemeColorEditor/blob/master/ThemeEditor.png?raw=true)

NOTE: Office 365 DLL Location and registry are slightly different.
The DLL for 32bit installtions (most common) is located by default here
C:\Program Files (x86)\Microsoft Office\root\VFS\ProgramFilesCommonX86\Microsoft Shared\VBA\VBA7.1

Once the theme is applied you will have to manually update 2 registry entries both in the following location
Computer\HKEY_CURRENT_USER\Software\Microsoft\VBA\7.1\Common
CodeBackColors =  2 7 1 13 15 2 2 2 11 9 0 0 0 0 0 0
CodeForeColors = 13 5 12 1 6 15 8 5 1 1 0 0 0 0 0 0
 
Based on the information found on: https://github.com/dimitropoulos/VBECustomColors
![alt tag](https://github.com/gallaux/VBEThemeColorEditor/blob/master/ExampleColors.png?raw=true)
