apt update
apt upgrade -y

apt install sudo

apt install -y software-properties-common
apt install -y postgresql-10
apt install -y postgresql-client-10

sudo -u postgres psql -c "ALTER USER postgres WITH PASSWORD 'postgres';"

sudo -u postgres psql -c "CREATE EXTENSION adminpack;"

psqluser="pythonuser"   # Database username
psqlpass="pythonuser"  # Database password
psqldb="data"   # Database name

printf "CREATE ROLE pythonuser WITH LOGIN SUPERUSER INHERIT CREATEDB CREATEROLE NOREPLICATION PASSWORD 'pythonuser';\nCREATE DATABASE $psqldb WITH OWNER $psqluser;" > createuserdb.sql

sudo -u postgres psql -f createuserdb.sql

echo "host all  all    0.0.0.0/0 md5" >> /etc/postgresql/10/main/pg_hba.conf

echo "listen_addresses='*'" >> /etc/postgresql/10/main/postgresql.conf

service postgresql restart
