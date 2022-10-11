#!/bin/bash

if [ "$1" == "--help" ]
then
    echo " "
	echo "Usage: generateDockerImage.sh [OPTIONS]"
    echo " "
    echo "Options:"
    echo "   --push     Will push the docker image to the docker.hcllabs.net repository"
    echo " "
	exit 0
fi

cd ..

# Build a docker image
timestamp=$(date +%Y%m%d%H%M)
#docker build --build-arg BUILD_NUMBER=$timestamp -f deploy/dockerfile -t hcllabs/ews .
DOCKER_BUILDKIT=1 docker build --no-cache --progress=plain --secret id=npmrcpat,src=$HOME/.npmrc --build-arg BUILD_NUMBER=$timestamp -f deploy/dockerfile -t hcllabs/ews .

# Create a tag to push to the HCLLabs docker repository
docker tag hcllabs/ews docker.hcllabs.net/ews:ews-$timestamp

# Push the image to the HCLLabs docker repository
if [ "$1" == "--push" ]
then
    docker push docker.hcllabs.net/ews:ews-$timestamp
fi

cd deploy 