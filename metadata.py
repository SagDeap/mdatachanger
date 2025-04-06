# -*- coding: cp1251 -*-
import os
import ezdxf
from ini_reader import read_ini
from odf.opendocument import load as load_odt
from odf.meta import UserDefined
from docx import Document

def load_extensions(file_name='extensions.txt'):
    
    base_path = os.getcwd()
    extensions_path = os.path.join(base_path, file_name)

    if os.path.exists(extensions_path):
        with open(extensions_path, 'r', encoding='cp1251') as file:
            extensions = file.readlines()
        return [ext.strip() for ext in extensions]
    else:
        raise FileNotFoundError(f"Файл {extensions_path} не найден.")

def update_docx_metadata(file_path, metadata):
    
    MAX_LENGTH = 255

    def truncate(value):
        return value[:MAX_LENGTH] if len(value) > MAX_LENGTH else value

    try:
        doc = Document(file_path)
    except Exception:
        raise ValueError(f"Не удалось открыть файл {file_path}. Убедитесь, что это корректный .docx файл.")

    core_props = doc.core_properties

    
    for key in ['title', 'subject', 'comments']:
        if key in metadata:
            try:
                setattr(core_props, key, truncate(metadata[key]))
            except (UnicodeEncodeError, UnicodeDecodeError):
                setattr(core_props, key, truncate(metadata[key])) 

   
    if 'tags' in metadata:
        core_props.keywords = truncate(metadata['tags'])

    # Обновляем категории, если они присутствуют
    if 'categories' in metadata:
        core_props.category = truncate(metadata['categories'])  

    doc.save(file_path)

def update_odt_metadata(file_path, metadata):
    
    doc = load_odt(file_path)
    meta = doc.meta

    def update_or_create_element(meta, name, value):
        existing_elements = meta.getElementsByType(UserDefined)
        for element in existing_elements:
            if element.getAttribute("name") == name:
                element.firstChild.data = value
                return
        new_element = UserDefined(name=name)
        new_element.addText(value)
        meta.addElement(new_element)

    if 'title' in metadata:
        update_or_create_element(meta, "title", metadata['title'])
    if 'subject' in metadata:
        update_or_create_element(meta, "subject", metadata['subject'])
    if 'tags' in metadata:
        update_or_create_element(meta, "keywords", metadata['tags'])
    if 'comments' in metadata:
        update_or_create_element(meta, "description", metadata['comments'])

    doc.save(file_path)

def update_dxf_metadata(file_path, metadata):
    
    try:
        # Открываем DXF файл
        doc = ezdxf.readfile(file_path)
    except Exception as e:
        raise ValueError(f"Не удалось открыть файл {file_path} как DXF: {e}")

    
    metadata_dict_name = "ACAD_METADATA"
    try:
        
        if metadata_dict_name in doc.rootdict:
            metadata_dict = doc.rootdict.get(metadata_dict_name)
            print(f"Словарь {metadata_dict_name} найден в rootdict.")
        else:
            
            metadata_dict = doc.objects.add_dictionary()
            doc.rootdict[metadata_dict_name] = metadata_dict
            print(f"Словарь {metadata_dict_name} создан и добавлен в rootdict.")
    except Exception as e:
        raise ValueError(f"Ошибка создания или получения словаря метаданных: {e}")

    
    try:
        for key, value in metadata.items():
            if isinstance(value, str):
               
                xrecord = doc.objects.add_xrecord(owner=metadata_dict.dxf.handle)
                xrecord_data = [(1, value)]  
                xrecord.extend(xrecord_data) 
                metadata_dict[key.upper()] = xrecord  
                print(f"Метаданные для {key} добавлены в словарь.")
            else:
                raise ValueError(f"Значение для ключа {key} должно быть строкой.")
    except Exception as e:
        raise ValueError(f"Ошибка добавления данных в словарь: {e}")

    # Сохраняем изменения
    try:
        doc.saveas(file_path)
        print(f"Метаданные успешно обновлены в файле: {file_path}")
    except Exception as e:
        raise ValueError(f"Ошибка сохранения файла {file_path}: {e}")

def attempt_generic_update(file_path, metadata):
    """Попытка обновить метаданные в неизвестном формате файла."""
    try:
        with open(file_path, "r+b") as f:
            content = f.read()
            for key, value in metadata.items():
                if isinstance(value, str):
                    try:
                        value_bytes = value.encode('cp1251')
                    except UnicodeEncodeError:
                        value_bytes = value.encode('cp1251')
                    key_bytes = key.encode('cp1251')
                    if key_bytes in content:
                        content = content.replace(key_bytes, value_bytes[:255])
            f.seek(0)
            f.write(content)
            f.truncate()
        print(f"Попытка обновления метаданных завершена для: {file_path}")
    except Exception as e:
        print(f"Не удалось обновить файл {file_path}: {e}")

def process_files(directory, ini_file):
    """Рекурсивная обработка файлов в указанной папке."""
    metadata = read_ini(ini_file)

    base_path = os.getcwd() 
    extensions_file = os.path.join(base_path, "extensions.txt")

    allowed_extensions = load_extensions(extensions_file)
    extensions = {
        '.docx': update_docx_metadata,
        '.odt': update_odt_metadata,
        '.dxf': update_dxf_metadata  
    }

    processed_files = []
    failed_files = []

    for root, _, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            ext = os.path.splitext(file)[-1].lower()

            if ext in extensions and ext.lstrip('.') in allowed_extensions:
                try:
                    extensions[ext](file_path, metadata)
                    processed_files.append(file_path)
                except Exception as e:
                    print(f"Ошибка обработки файла {file_path}: {e}")
                    failed_files.append(file_path)
            elif ext.lstrip('.') in allowed_extensions:
                try:
                    attempt_generic_update(file_path, metadata)
                    processed_files.append(file_path)
                except Exception as e:
                    print(f"Ошибка обработки файла {file_path}: {e}")
                    failed_files.append(file_path)

    return processed_files, failed_files
