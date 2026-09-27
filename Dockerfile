# Python 3.10 का बेस इमेज लें
FROM python:3.10-slim

# LibreOffice और ज़रूरी सिस्टम फाइल्स इंस्टॉल करें
RUN apt-get update && apt-get install -y libreoffice default-jre

# काम करने का फोल्डर सेट करें
WORKDIR /app

# रिक्वायरमेंट्स फाइल कॉपी करके सारे पैकेजेस इंस्टॉल करें
COPY requirements.txt .
RUN pip install --no-cache-dir gunicorn
RUN pip install --no-cache-dir -r requirements.txt

# बाकी सारा कोड सर्वर पर कॉपी करें
COPY . .

# ग यूनिकॉर्न (Gunicorn) सर्वर चालू करें
CMD ["gunicorn", "-b", "0.0.0.0:10000", "app:app", "--timeout", "120"]
