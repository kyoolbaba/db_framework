# from pudbo_v11 import generate_dashboard
from db_fw12 import generate_dashboard

import pickle

# with open("../AIDER_AT/PICKLE_FILES/data.pkl", "rb") as f:
#     data = pickle.load(f)

with open("../HEALTHIUM/CODE BASE/PICKLE_FILES/data_1.pkl", "rb") as f:
    data = pickle.load(f)


dd=[]
page_name=[]
for i in data.keys():
    print(i)
    dd.append(data[i])
    page_name.append(i)


generate_dashboard(
    dicts         = dd,
    names         = page_name,
    output        = 'healht_data_.html',
    default_theme = 'Dark Blue',
    rows_per_page = 25,
    title         = 'Healthium — Charts Report',
)

# generate_dashboard(
#     dicts         = [dict1, dict2, dict3, dict4],
#     names         = ['Color By', 'Bar Modes', 'Combo', 'Annotate'],
#     output        = 'chart_features_v6_test.html',
#     default_theme = 'Dark Blue',
#     rows_per_page = 25,
#     title         = 'Chart Features v6',
# )
