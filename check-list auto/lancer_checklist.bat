@echo off 
title Script Checklist Automatique - CD13
color 0B

echo ==========================================================================
echo			LANCEMENT DE LA CHECKLIST AUTOMATIQUE 
echo ==========================================================================
echo.

cd /d "C:\Users\lbenadyext\Desktop\check-list-auto-main\check-list auto"
python "auto_checklist.py"

cd /d "C:\Users\lbenadyext\Desktop\check-list-auto-main\check-list auto"
python "prep_mail.py"

echo ==========================================================================
echo		    TRAVAIL TERMINE ! Verifie le dissier OUTPUT 
echo ==========================================================================
echo.
pause
