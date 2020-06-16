# https://docs.docker.com/engine/install/ubuntu/
apt remove docker docker-engine docker.io containerd runc
apt update
apt install \
    apt-transport-https \
    ca-certificates \
    curl \
    gnupg-agent \
    software-properties-common
curl -fsSL https://download.docker.com/linux/ubuntu/gpg | sudo apt-key add -
add-apt-repository \
   "deb [arch=amd64] https://download.docker.com/linux/ubuntu \
   $(lsb_release -cs) \
   stable"
apt update
apt install docker-ce docker-ce-cli containerd.io
docker run hello-world
docker pull ubuntu
docker run -i -t ubuntu bash
