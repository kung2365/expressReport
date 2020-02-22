@echo on


Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
call pip install virtualenv
call virtualenv venv
call ./venv/Scripts/activate
call pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
call pip install dbf-to-sqlite
