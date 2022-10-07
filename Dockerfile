FROM public.ecr.aws/lambda/python:3.9

COPY ./src ./src
COPY ./requirements.txt ./requirements.txt

RUN pip install -r ./requirements.txt

CMD ["src.server.handler"]