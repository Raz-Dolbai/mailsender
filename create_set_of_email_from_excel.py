import pandas as pd


def create_list_excel(path):
    """Open file EXCEL and save data in list"""
    with pd.ExcelFile(path) as code_image:
        email = set(map(str, (pd.read_excel(code_image))['e-mail'].tolist()))
    return list(email)


