FROM python:3.9-alpine
WORKDIR /app
COPY . /app
COPY templates/ /app/templates
RUN pip install -r requirements.txt
EXPOSE 3000
CMD python ./app.py