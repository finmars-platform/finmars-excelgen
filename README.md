Excel generator

Docker image name registry.finmars.com/excelgen

How to start a service

    docker-compose up -d
    
Test request:
    
    curl -X POST -vv  0.0.0.0:80 -o ./file.xlsx
    
JSON request:

    curl --header "Content-Type: application/json" \
      -X POST \
      --data @data.json \
      -vv 0.0.0.0:80 -o ./file.xlsx
      
JSON request for windows:

    curl --header "Content-Type: application/json" \
      -X POST \
      --data @data.json \
      -vv 127.0.0.1:80 -o ./file.xlsx