Excel generator

Docker image name registry.finmars.com/excelgen

How to start a service

    docker-compose up -d
    
How to convert html page to pdf

    curl -X POST -vv -F 'file=@portal.html' 0.0.0.0:80 -o ./file.pdf
   
Test request:
    
    curl -X POST -vv  0.0.0.0:80 -o ./file.pdf
    
JSON request:

    curl --header "Content-Type: application/json" \
      -X POST \
      --data @data.json \
      -vv 0.0.0.0:80 -o ./file.pdf
      
JSON request for windows:

    curl --header "Content-Type: application/json" \
      -X POST \
      --data @data.json \
      -vv 127.0.0.1:80 -o ./file.pdf