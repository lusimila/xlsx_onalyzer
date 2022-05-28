# Из библиотеки YAML импортируем одну функцию safe_load,
# которая принимает yaml файл и превращает его в обычный словарь

from yaml import safe_load

def load_config() -> dict:
    with open('config.yaml', 'r', encoding='utf-8') as file:
        result = safe_load(file)
        return result

CONFIG = load_config()
