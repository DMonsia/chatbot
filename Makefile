.ONESHELL:

.PHONY: build
build:
	docker build -t dmonsia/api-chatbot ./chatbot/

.PHONY: push
push:
	docker login -u ${DOCKER_USER} -p ${DOCKER_PWD} && \
	docker push ${DOCKER_USER}/api-chatbot

.PHONY: run
run:
	docker run -d \
		--name api-chatbot \
		-p 8000:8000 \
		-v ${DATA_FILE}:/chatbot/data \
		dmonsia/api-chatbot
