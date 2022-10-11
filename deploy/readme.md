# Building a docker image

To build a docker image that can be used to create docker container that runs the EWS middleware, run the following script:

1. Open a terminal in the ews-middleware/deploy directory. 

1. Create the docker image:
    ```
    ./generateDockerImage.sh 
    ```

    If you want to create the image and upload it to the HCLLabs docker registery:
    ```
    ./generateDockerImage.sh --push
    ```

## Manual steps

If you don't use the generateDockerImage.sh script, you can build the docker image manually by following these steps:

1. Open a terminal in the ews-middleware directory. 

1. Create the docker image: 
    ```
    DOCKER_BUILDKIT=1 docker build --no-cache --progress=plain --secret id=npmrcpat,src=$HOME/.npmrc -f ./deploy/dockerfile -t hcllabs/ews .
    ```
1. You will see the docker image listed:
    ```
    docker images
    ```
1. Now create a tag to push to our docker repository. "latest" is the tag. You can tag it with something different if needed. 
    ```
    docker tag hcllabs/ews docker.hcllabs.net/ews:latest
    ```
1. Show list of images. You will see two my hcllabs/ews and one with docker.hcllabs.net. Note they both have the same id, so they point to the same image. 
    ```
    docker images
    ```

## Creating a container from the image

This is optional, but suggested in order to insure your image is correct. 

1. To create a container in your local docker installation to test the image (8085 is the external port):
    ```
    docker run -p 8085:3000 --name ews -e KEEP_BASE_URL=http://<keep_server>:8880 -d hcllabs/ews
    ```

    You can point the EWS to use an instance of Keep by setting the KEEP_BASE_URL environment variable when creating the image. If you are pointing to Keep running in another container on your machine, then the keep_server should be the ip address of your machine. For example -e KEEP_BASE_URL=http://10.0.1.16:8880. Don't use 127.0.0.1 since it will resolve to the EWS containers local host address.

1. You should see the container running:
    ```
    docker ps -a
    ```
1. You should be able to access the container via the external port:
    ```
    curl -i localhost:8085/EWS/Services.wsdl
    ```

Other useful docker commands
- Delete a docker image:
    ```
    docker image rm <image_id>
    ```

- To point to a different instance of keep you can set an environment variable using the -e option:
    ```
    docker run -e KEEP_BASE_URL=http://<keep_host>  -p 8085:3000 -d hcllabs/ews
    ```
    Note: You can also change the port EWS middleware listens on by setting the PORT environment variable. 

- To list the containers (will show container id):
    ```
    docker container ls -a
    ```

- To show the logs for a container:
    ```
    docker logs <container_id>
    ```

- To start a command line for the container:
    ```
    docker exec -it <container_id> /bin/bash
    ```
    Type exit to exit the container.

## Upload the image to the docker repo

If you want to deploy the image to OpenShift, you can upload the image to the docker repo:
```
docker push docker.hcllabs.net/ews:latest
```

# Deploy EWS image to OpenShift

Note: You will need to install the OpenShift client on your machine. You can download it from [OpenShift 4 clients](https://mirror.openshift.com/pub/openshift-v4/clients/oc/latest/). You need to put the oc command executable in your Path. On Mac you can place it in /usr/local/bin. 

1. Login OKD at https://okd.hcllabs.net:8443/ 

    If you don't have an OKD account, contact George Fridrich.

    Note: You should be able to do these steps from this website if you want to use a GUI instead of the command line. 

1. In the top right hand corner click on your login name and select "Copy Login Command".

1. Open a terminal in the ews-middleware directory. 

1. Paste the login command to the terminal window. It will look something like this: 
    ```
    oc login https://okd.hcllabs.net:8443 --token=<generated_token>
    ```

1. If you don't have a project, create a new project (only do this once):
    ```
    oc new-project <project_name>
    ```

1. Create a secret for accessing the image to the docker registry:
    ```
    oc create secret docker-registry docker.hcllabs.net   --docker-server=docker.hcllabs.net   --docker-username=hcllabs   --docker-password=Images4Labs   --docker-email=none

    oc secrets link default docker.hcllabs.net --for=pull
    ```

1. Create the pod:
    ```
    oc create -f deploy/statefulset.yaml
    ```

1. List information about the pods to get the pod name. It should be something like ews-0. 
    ```
    oc get pods

    oc describe pods

    oc describe pod <pod_name>
    ```

1. When you first create the pod, if it does not start correctly, you can delete a pod. It is setup to automatically get recreated. 
    ```
    oc delete pod <pod_name> 
    ```

1. You can see the system out with this command. You should see the ews server started on port 3000. 
    ```
    oc logs <pod_name>
    ```

1. Create service:
    ```
    oc create -f deploy/services.yaml
    ```

1. List service in yaml format: 
    ```
    oc get service -o yaml
    ```

1. Create your route:
    ```
    oc create -f deploy/route.yaml
    ```

1. List routes:
    ```
    oc get routes
    ```

## Other useful OpenShift commands

- Open a terminal to your pod:
    ```
    oc rsh <pod_name>
    ```
    To exit the terminal, type `exit`

- List the environment variables for the pod:
    ```
    oc set env pods <pod_name> --list
    ```

    Note: To change the environment variable, login using the web UI and goto Stateful Set for your pod and click on the Environment tab. 