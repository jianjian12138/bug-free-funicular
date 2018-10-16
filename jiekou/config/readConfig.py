# coding:utf-8
import os
import ConfigParser
      #这个模块定义了一个ConfigParser类，该类的作用是使用配置文件生效，配置文件的格式和windows的INI文件的格式相同

cur_path = os.path.dirname(os.path.realpath(__file__))
configPath = os.path.join(cur_path, "cfg.ini")
conf = ConfigParser.ConfigParser()
conf.read(configPath)

smtp_server = conf.get("email", "smtp_server")

sender = conf.get("email", "sender")

psw = conf.get("email", "psw")

receiver = conf.get("email", "receiver")

port = conf.get("email", "port")
