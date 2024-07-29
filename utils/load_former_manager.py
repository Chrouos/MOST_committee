
def get_former_manager(file_path) -> list:
    with open(file_path, 'r') as f:
        former_manager = f.read().split('\n')
    return former_manager