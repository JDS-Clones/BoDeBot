call rename user.scp ORIG_user.scp
call rename required.scp ORIG_required.scp
call rename DIST_user.scp user.scp
call rename DIST_required.scp required.scp
cd "bot object"
call rename pass.txt ORIG_pass.txt
call rename DIST_pass.txt pass.txt
cd ..
