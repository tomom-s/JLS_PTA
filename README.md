# JLS_PTA
## Description
This tool has been developed for Boston Japanese Language School PTA to automate the process.

### 台帳作成
Selecting the csv exported from WES will generate the data that needs to be shared for the further PTA processing.

### 台帳から名簿の作成
Export CSV from 台帳 and select it to generate each class's data to share with the classes.


## How to run locally

1. Install docker desktop (one time only)
   https://www.docker.com/products/docker-desktop/
1. Build docker image (one time only)

    ```shell
    # From terminal, run the following command at /app folder:
    docker build -t pta-app .
    ```

1. Run

    ```shell
    docker run --rm -p 8080:5000 pta-app
    ```
    if you want to enable debug mode, run the following command instead:
    ```shell
    docker run --rm -it -p 8080:5000 -e FLASK_DEBUG=True pta-app
    ```

1. Open browser and go to http://localhost:8080
