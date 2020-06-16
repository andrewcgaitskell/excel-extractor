## building from a mac

### pre-requisites

docker desktop installed

## start point

https://hostpresto.com/community/tutorials/how-to-create-a-docker-container-using-an-interactive-shell/

## pull latest ubuntu image

	docker pull ubuntu

## connect to image

	docker run -i -t ubuntu bash

## update and install

	apt update
	apt upgrade
	apt install git

# 1137 committed

	mkdir /home/ubuntu; cd /home/ubuntu

	git clone https://github.com/andrewcgaitskell/celldb.git

cd /home/ubuntu/celldb/scripts

## update and install required software

	bash 1a_docker_postgresql.sh

 	pg_ctlcluster 12 main start

 	bash 1b_create_postgrs_users.sh 

## exit from postgrs user

	exit

## exit from image

	exit

## save image

	docker ps -a

find container id from above

use following with container_id found in the aboe

	docker commit container_id your_container_name

# created 1148 now connect

	run -i -t ubuntu_1148 bash

## open firewalls

	bash 1c_open_firewalls.sh

## install python

	bash 2_install-python3.sh

## install python libraries

	bash 3_environment.sh

	exit

	docker ps -a

	docker commit container_id your_container_name

# 1156 created

## config notebook security

apt install nano

jupyter notebook --generate-config

cd /home/ubuntu/celldb/scripts


nano /root/.jupyter/jupyter_notebook_config.py

paste

c.NotebookApp.token = ''
c.NotebookApp.password = ''

ctrl x - save

test run inside image

run

jupyter notebook --ip=0.0.0.0 --port=8080 --allow-root --no-browser

# commited as 1213

## make a folder to container dockerfile

	mkdir folder_name

## change to folder

	cd folder_name

## contents for Dockerfile

	FROM ubuntu_1213             
	EXPOSE 8080
	CMD ["jupyter", "notebook", "--ip=0.0.0.0", "--port=8080", "--allow-root", "--no-browser"]

## add docker file

	nano Dockerfile

FROM ubuntu_celldb_1011            
EXPOSE 8080
CMD ["bash","/home/ubuntu/voila/scripts/oncontaineropen.sh"]

============================

add data

manually started postgrs

ran load data shell script

committed image as ubuntu_1526

============================

created script file to acivate postgrs
then run notyebook server

committed as 1537

=====================

changed script file to executable

committed as 1542

## list available images

	docker images


docker remove

remove all containers

docker rm $(docker ps -a -q)

remove images

docker images

docker image rm [OPTIONS] IMAGE [IMAGE...]


=======================

FROM ruby:2.5
RUN apt-get update -qq && apt-get install -y nodejs postgresql-client
RUN mkdir /myapp
WORKDIR /myapp
COPY Gemfile /myapp/Gemfile
COPY Gemfile.lock /myapp/Gemfile.lock
RUN bundle install
COPY . /myapp

# Add a script to be executed every time the container starts.
COPY entrypoint.sh /usr/bin/
RUN chmod +x /usr/bin/entrypoint.sh
ENTRYPOINT ["entrypoint.sh"]
EXPOSE 3000

# Start the main process.
CMD ["rails", "server", "-b", "0.0.0.0"]

=======================

FROM voila_sql_1641_container 
EXPOSE 5432
CMD ["jupyter", "notebook", "--ip=0.0.0.0", "--port=8080", "--allow-root"]

jupyter notebook --ip=0.0.0.0 --port=8080 --allow-root


==================
starting at 1542 image

docker run -i -t ubuntu_1542 bash

cd /home/ubuntu

git clone https://github.com/andrewcgaitskell/celldb.git

cd celldb

bash 3_environment.sh

exit

docker ps -a

find container id from above

df51322ebc35

use following with container_id found in the aboe


	docker commit container_id your_container_name
	
	docker commit df51322ebc35 ubuntu_celldb_1011

Dockerfile

FROM ubuntu_celldb_1011            
EXPOSE 8080
CMD ["bash","/home/ubuntu/voila/scripts/oncontaineropen.sh"]

