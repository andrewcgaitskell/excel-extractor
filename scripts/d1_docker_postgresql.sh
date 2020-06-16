apt update
apt upgrade -y

apt install sudo

apt install -y software-properties-common
apt install -y postgresql-12
apt install -y postgresql-client-12

psqluser="pythonuser"   # Database username
psqlpass="pythonuser"  # Database password
psqldb="data"   # Database name

printf "CREATE ROLE pythonuser WITH LOGIN SUPERUSER INHERIT CREATEDB CREATEROLE NOREPLICATION PASSWORD 'pythonuser';\nCREATE DATABASE $psqldb WITH OWNER $psqluser;" > createuserdb.sql

chown postgres createuserdb.sql

su postgres

