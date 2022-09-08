@ECHO OFF

echo Installing Email Processor Service...
sc create EmailProcessor binPath= "C:\Users\KenNguyen\Desktop\api\sp4a-api\EmailProcessor\bin\Debug"
echo Done