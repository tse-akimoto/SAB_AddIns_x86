rem objが残っていてInstaller Projectsがエラーになったことがあったのでお掃除するバッチ

rmdir /S /Q .\Addins\

rmdir /S /Q .\ExcelAddInSAB\obj

rmdir /S /Q .\PowerPointAddInSAB\bin
rmdir /S /Q .\PowerPointAddInSAB\obj

rmdir /S /Q .\WordAddInSAB\bin
rmdir /S /Q .\WordAddInSAB\obj

rmdir /S /Q .\AddInsLibrary\bin
rmdir /S /Q .\AddInsLibrary\obj

rmdir /S /Q .\Setup\Debug
rmdir /S /Q .\Setup\Release

pause
