echo "host all  all    0.0.0.0/0 md5" >> /etc/postgresql/12/main/pg_hba.conf

echo "listen_addresses='*'" >> /etc/postgresql/12/main/postgresql.conf

service postgresql restart
