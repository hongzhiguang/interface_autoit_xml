from utils.MouseControl import *
from config.ProjVar import *
import time

def moveServerIPEntry():
    MouseControl.move(server_ip_locate)
    MouseControl.click(server_ip_locate)
    time.sleep(1)

def connectNMSServer():
    MouseControl.move(connect_locate)
    MouseControl.click(connect_locate)

def DisconnectNMSServer():
    MouseControl.move(disconnect_locate)
    MouseControl.click(disconnect_locate)

def moveToRequest():
    MouseControl.move(request_param_locate)
    MouseControl.click(request_param_locate)
    time.sleep(1)

def moveToResponse():
    MouseControl.move(response_data_locate)
    MouseControl.click(response_data_locate)
    time.sleep(1)

def send():
    MouseControl.move(send_button_locate)
    MouseControl.click(send_button_locate)
    time.sleep(1)

def clear():
    MouseControl.move(send_button_locate)
    MouseControl.click(send_button_locate)
    time.sleep(1)