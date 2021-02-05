FROM ubuntu:20.04

RUN sed 's/main$/main universe/' -i /etc/apt/sources.list
RUN apt-get update
RUN DEBIAN_FRONTEND=noninteractive apt-get upgrade -y --no-install-recommends

RUN DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends build-essential xorg libssl-dev libxrender-dev wget gdebi

RUN apt-get install -y python3-pip
RUN pip3 install werkzeug executor gunicorn
RUN pip3 install openpyxl
RUN pip3 install pyyaml

WORKDIR /var/www/excelgen
COPY . .
EXPOSE 80

ENTRYPOINT ["/usr/local/bin/gunicorn"]

CMD [ "-b", "0.0.0.0:80", "--log-file", "-", "app:application", "--reload", "--log-level", "DEBUG", "--timeout", "90"]