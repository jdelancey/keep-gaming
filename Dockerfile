FROM python

WORKDIR /usr/local/dk

# https://grigorkh.medium.com/fix-tzdata-hangs-docker-image-build-cdb52cc3360d
ENV TZ=America/New_York
RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo $TZ > /etc/timezone

# ensure package manager is up to date
RUN apt update -y

# install google chrome
RUN wget -q -O - https://dl-ssl.google.com/linux/linux_signing_key.pub | apt-key add -
RUN sh -c 'echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" >> /etc/apt/sources.list.d/google-chrome.list'
RUN apt-get -y update
RUN apt-get install -y google-chrome-stable

# install chromedriver
RUN apt-get install -yqq unzip
RUN wget -O /tmp/chromedriver.zip http://chromedriver.storage.googleapis.com/`curl -sS chromedriver.storage.googleapis.com/LATEST_RELEASE`/chromedriver_linux64.zip
RUN unzip /tmp/chromedriver.zip chromedriver -d /usr/local/bin/

RUN pip3 install --no-cache-dir bs4
RUN pip3 install --no-cache-dir google-api-python-client
RUN pip3 install --no-cache-dir google-auth-oauthlib
RUN pip3 install --no-cache-dir requests
RUN pip3 install --no-cache-dir selenium

# set display port to avoid crash
ENV DISPLAY=:99

COPY *.py /usr/local/dk/
COPY ./keys/ /usr/local/dk/keys/
RUN chmod 755 /usr/local/dk/dk.py

CMD ["python3", "-u", "/usr/local/dk/dk.py", "--new", "--ncaam"]
