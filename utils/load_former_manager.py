def get_former_manager(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        former_manager = f.read().split('\n')
    return former_manager
