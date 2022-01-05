@echo off
set ubic=%cd%
set time1=%time:~0,2%
set /a time2=%time1% + 1
set time3=%time2%%time:~2,3%
at %time3% "%ProgramFiles%\Symantec\Symantec Endpoint Protection\Smc.exe" -start >> %userprofile%\null.log 2>&1
at %time3% "%ProgramFiles(x86)%\Symantec\Symantec Endpoint Protection\Smc.exe" -start >> %userprofile%\null.log 2>&1
cls
echo           =============================================================
echo          SYMANTEC ================================= PERMISOS INSTALACION
echo          ===============================================================
echo          ===                                                         ===
echo          ===                                                         ===
echo          ===                                                         ===
echo          ===                Procesando permisos, por favor espere... ===
echo          ===============================================================
echo          =============================== By: Richard Molina - INTERLAN
echo           =============================================================
"%ProgramFiles%\Symantec\Symantec Endpoint Protection\Smc.exe" -p obi-wan -stop >> %userprofile%\null.log 2>&1
"%ProgramFiles(x86)%\Symantec\Symantec Endpoint Protection\Smc.exe" -p obi-wan -stop >> %userprofile%\null.log 2>&1
timeout /t 5 /nobreak >> %userprofile%\null.log 2>&1
cls
echo           =============================================================
echo          SYMANTEC ================================= PERMISOS INSTALACION
echo          ===============================================================
echo          === Ahora puede realizar la instalacion.                    ===
echo          ===                                                         ===
echo          === Una vez terminada la instalación regrese a esta ventana ===
echo          ===                    y presione una tecla para continuar. ===
echo          ===============================================================
echo          =============================== By: Richard Molina - INTERLAN
echo           =============================================================
pause.
"%ProgramFiles%\Symantec\Symantec Endpoint Protection\Smc.exe" -start >> %userprofile%\null.log 2>&1
"%ProgramFiles(x86)%\Symantec\Symantec Endpoint Protection\Smc.exe" -start >> %userprofile%\null.log 2>&1
cls
echo           =============================================================
echo          SYMANTEC ================================= PERMISOS INSTALACION
echo          ===============================================================
echo          ===                                                         ===
echo          ===                                                         ===
echo          ===                                                         ===
echo          ===                                   Proceso finalizado... ===
echo          ===============================================================
echo          =============================== By: Richard Molina - INTERLAN
echo           =============================================================
"%ProgramFiles%\Symantec\Symantec Endpoint Protection\Smc.exe" -updateconfig >> %userprofile%\null.log 2>&1
"%ProgramFiles(x86)%\Symantec\Symantec Endpoint Protection\Smc.exe" -updateconfig >> %userprofile%\null.log 2>&1
timeout /t 3 /nobreak >> %userprofile%\null.log 2>&1
del %userprofile%\null.log
del "%ubic%\Installs.cmd"