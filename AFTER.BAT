call rename user.scp DIST_user.scp
call rename required.scp DIST_required.scp
call rename ORIG_user.scp user.scp
call rename ORIG_required.scp required.scp
cd "bot object"
call rename pass.txt DIST_pass.txt
call rename ORIG_pass.txt pass.txt
cd ..
