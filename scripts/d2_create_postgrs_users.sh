psql -c "ALTER USER postgres WITH PASSWORD 'postgres';"

psql -c "CREATE EXTENSION adminpack;"

psql -f createuserdb.sql

exit
