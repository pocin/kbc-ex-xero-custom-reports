FROM ubuntu:18.04

RUN apt-get update && \
    apt-get install -y python3 python3-pip curl chromium-browser && \
    rm -rf /var/lib/apt/lists/* && \
    curl https://chromedriver.storage.googleapis.com/2.40/chromedriver_linux64.zip -o ~/chromedriver.zip && \
    python3 -m zipfile -e ~/chromedriver.zip /usr/local/bin/ && \
    chmod a+x /usr/local/bin/chromedriver && \
    rm ~/chromedriver.zip

RUN pip3 install selenium requests pandas openpyxl xlrd

WORKDIR /code/
COPY . /code/
CMD python3 -u /code/main.py
