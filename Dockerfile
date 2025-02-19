FROM python:3.11

# Обновляем apt и ставим системные пакеты (ldap + oracle + компилятор)
RUN apt-get update && apt-get install -y \
    libldap2-dev \
    libsasl2-dev \
    libssl-dev \
    build-essential \
    libaio1 \
    wget \
    unzip

# Устанавливаем Oracle Instant Client
WORKDIR /opt/oracle
RUN wget https://download.oracle.com/otn_software/linux/instantclient/2370000/instantclient-basic-linux.x64-23.7.0.25.01.zip \
    && unzip instantclient-basic-linux.x64-23.7.0.25.01.zip

# Настраиваем переменные окружения для Oracle
ENV LD_LIBRARY_PATH=/opt/oracle/instantclient_23_7:${LD_LIBRARY_PATH}
ENV ORACLE_HOME=/opt/oracle/instantclient_23_7

# Переходим в /app
WORKDIR /app

# Копируем файл requirements.txt и устанавливаем зависимости
COPY requirements.txt .
RUN pip install --upgrade pip \
    && pip install -r requirements.txt

# Копируем папку Main_Modul в /app/Main_Modul
COPY Main_Modul /app/Main_Modul 

# Запускаем main.py, который лежит в Main_Modul
CMD ["python", "Main_Modul/main.py"]