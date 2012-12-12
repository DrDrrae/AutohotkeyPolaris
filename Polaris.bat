taskkill /f PolarisAHK.exe

robocopy "\\ao-ntserv1\Departments\Parkville\Staff Folders\Andrea\Polaris" %USERPROFILE%\Polaris

START %USERPROFILE%\Polaris\PolarisAHK.exe