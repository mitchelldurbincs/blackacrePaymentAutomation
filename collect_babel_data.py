import babel
import os
import shutil

def collect_babel_data():
    print("Collecting Babel data...")  # Add this line
    babel_root = os.path.dirname(babel.__file__)
    data_dir = os.path.join(babel_root, 'locale-data')
    return [(data_dir, 'babel/locale-data')]

datas = collect_babel_data()
print("Babel data collected")  # Add this line