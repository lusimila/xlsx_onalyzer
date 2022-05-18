from yaml import safe_load


def load_config() -> dict:
    with open('config.yaml', 'r', encoding='utf-8') as file:
        return safe_load(file)


CONFIG = load_config()
