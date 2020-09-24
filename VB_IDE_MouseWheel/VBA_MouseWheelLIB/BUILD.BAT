@echo off
\masm32\bin\ml /c /coff /Cp abfr.asm
\masm32\bin\link /DLL /DEF:abfr.def /SUBSYSTEM:WINDOWS /LIBPATH:\masm32\lib abfr.obj /COMMENT:"VB6 IDE-Mouse Wheel ! Created By Vanja Fuckar,19.11.2003 , email:inga@vip.hr,version 1.1" 
pause