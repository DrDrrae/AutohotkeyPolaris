robocopy "\\ao-ntserv1\Departments\Parkville\Staff Folders\Andrea" %USERPROFILE%\Polaris Polaris.exe
robocopy "\\ao-ntserv1\Departments\Parkville\Staff Folders\Andrea" %USERPROFILE%\Polaris /xo /xn settings.ini

START %USERPROFILE%\Polaris\Polaris.exe