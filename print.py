import socket              
mysocket = socket.socket(socket.AF_INET,socket.SOCK_STREAM)         
host = "192.168.1.8" 
port = 9100   
try:           
	mysocket.connect((host, port)) #connecting to host
	mysocket.send(b"^XA^A0N,50,50^FO50,50^FDSocket Test^FS^XZ")#using bytes
	mysocket.close () #closing connection
except:
	print("Error with the connection")