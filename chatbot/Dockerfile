FROM python:3.10

RUN apt-get update && \
    apt-get install -y wget curl && \
    apt-get clean && \
    apt-get remove --purge -y


WORKDIR /chatbot
RUN export PYTHONPATH=${PWD}

COPY ./requirements.txt /tmp/requirements.txt

RUN python3 -m pip install --upgrade pip && \
    pip install --no-cache-dir -r /tmp/requirements.txt && \
    rm -r /tmp/requirements.txt

RUN mkdir data
COPY ./src src
COPY ./main.py main.py

EXPOSE 8000
CMD ["--host", "0.0.0.0", "--port", "8000"]
ENTRYPOINT [ "uvicorn", "main:app" ]
