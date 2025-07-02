:: Paths
set CSC="%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
set OUTPUT=OrderModule
set DLL=DocumentFormat.OpenXml.dll

:: Compile order.exe
%CSC% /reference:%DLL% /reference:"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\WPF\WindowsBase.dll" /out:%OUTPUT%\order.exe Program.cs

:: Copy runtime dependency
copy /Y %DLL% %OUTPUT%\DocumentFormat.OpenXml.dll

echo Done. Your PowerShell module is in %OUTPUT%\
