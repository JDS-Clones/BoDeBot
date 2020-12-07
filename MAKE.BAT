@echo off
echo Rebuilding Project Binaries ...
set PATH=%PATH%;e:\devstudio\vb
vb5.exe /make group1.vbg
echo Packaging Source Code ...
call del bodebot.zip
call rename required.scp VALID_required.scp
call rename DIST_required.scp required.scp
cd ".\bot object"
call rename pass.txt VALID_pass.txt
call rename DIST_pass.txt pass.txt
call cd..
call ntzip -add -dir=relative bodebot.zip *.*
call ntzip -del bodebot.zip "bot object\VALID_pass.txt"
cd ".\bot object"
call rename pass.txt DIST_pass.txt
call rename VALID_pass.txt pass.txt
call cd..
call rename required.scp DIST_required.scp
call rename VALID_required.scp required.scp
