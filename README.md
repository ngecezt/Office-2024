# Office-2024
Install &amp; Activate Microsoft Office LTSC Porfessional Plus 2024 &amp; Visio LTSC Profesional 2024
<br>
<br>
-> Download office installer from : https://github.com/asheroto/Deploy-Office and then open it.<br>
![install office pro plus 2024 volume edition](image.png) <br>
-> Click Start &amp; wait for Installation to finish<br>
<br>
-> Open Office installer again and install Visio Pro<br>
![install visio pro 2024 volume edition](image-1.png) <br>
-> Click Start &amp; wait for Installation to finish<br>
<br>
<br>
1. Open Terminal As Administrator and paste this to the Terminal<br>
2. cd /d %ProgramFiles%\Microsoft Office\Office16<br>
3. for /f %x in ('dir /b ..\root\Licenses16\ProPlus2024VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"<br>
4. for /f %x in ('dir /b ..\root\Licenses16\VisioPro2024VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"<br>
<br>
(Optional Steps : Uninstall Preview Key)<br>
<br>
4a. cscript ospp.vbs /dstatus (check What Key Is Installed)<br>
4b. cscript ospp.vbs /unpkey:WXGKQ >nul (Last 5 Keys Visio Pro 2024 Preview Edition)<br>
4c. cscript ospp.vbs /unpkey:TX3T2 >nul (Last 5 Keys Pro Plus 2024 Preview Edition)<br>
<br>
5. cscript ospp.vbs /setprt:1688<br>
6. cscript ospp.vbs /inpkey:XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB (Office LTSC Pro Plus 2024 VL Key)<br>
7. cscript ospp.vbs /inpkey:B7TN8-FJ8V3-7QYCP-HQPMV-YY89G (Visio Pro LTSC 2024 VL Key)<br>
8. cscript ospp.vbs /sethst:e8.us.to<br>
9. cscript ospp.vbs /act<br>
