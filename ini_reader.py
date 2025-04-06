
# -*- coding: utf-8 -*-
import configparser
from charset_normalizer import from_path

def read_ini(ini_path):
    
    result = from_path(ini_path).best()
    if result is None:
        raise ValueError("Кодировка ini файла не определена или проблемы с файлом.")

    decoded_content = str(result)
    config = configparser.ConfigParser()
    config.read_string(decoded_content)
    metadata = {}
    for section in config.sections():
        metadata.update({key: value for key, value in config.items(section) if value.strip()})
    return metadata