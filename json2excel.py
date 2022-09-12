import openpyxl
import json
import sys
sys.path.append('/Users/meursault/OneDrive/ffs/')

wb = openpyxl.Workbook()

type_dict = {
    "<class 'int'>":"int",
    "<class 'str'>":"str",
    "<class 'dict'>":"dict",
    "<class 'NoneType'>":"NoneType"
}

def handle_data(data, key_list):
    for x in data:
        item = data[x]
        for key in key_list:
            try:
                item[key]
            except KeyError:
            # key_list1.append(key)
                item[key] = None
        data.update({x:item})
    # print(data)
    return data
def get_key_list(data):
    key_list1 = list(list(data.values())[0].keys())
    value_list = []
    for i, id in enumerate(data):
        # i是key的index，id是key
        item = data[id]    
        key_list = list(item.keys())
        values = list(item.values())
        if i == 0:
            value_list = values
        # 如果字典第一组key比后续的少
        for j , key in enumerate(key_list):
            if key not in key_list1:
                key_list1.append(key)
            # if values[j] not in value_list:
                value_list.append(values[j])
        # 如果字典第一组key比后续的多
        # for j , key in enumerate(key_list1):
        #     # j是key的index，key是key
        #     try:
        #         item[key]
        #     except KeyError:
        #         # key_list1.append(key)
        #         item[key] = None
    type_list = list(str(type(v)) for v in value_list)
    type_list = [str(type(1))] + type_list
    type_list = list(type_dict[x] for x in type_list)
    return key_list1, type_list

def get_data_attribute(data):
    data_attribute = []
    value_list = []
    type_list = []
    for x in data:
        keys = list(data[x].keys())
        values = list(data[x].values())
        for i, key in enumerate(keys):
            if key not in data_attribute:
                data_attribute.append(key)
                value_list.append(values[i])
    # print(key_list)
    type_list = list(str(type(v)) for v in value_list)
    data_attribute = ['id'] + data_attribute
    type_list = [str(type(1))] + type_list
    type_list = list(type_dict[x] for x in type_list)
    # print(type_list)
    return data_attribute

def get_key_index(data_attribute):
    key_index = {}
    for i, d in enumerate(data_attribute):
        key_index[d] = i +1
    # print(key_index)
    return key_index

def json2excel(path, data, sheet_name):
    key_list_data = get_key_list(data)
    key_list = key_list_data[0]
    type_list = key_list_data[1]
    data = handle_data(data, key_list)
    data_attribute = get_data_attribute(data)
    key_index = get_key_index(data_attribute)
    # key_index ={'series':2, 'index':3, 'score':4, 'color':5, 'temperature':6, 'sweetness':7, 'calorie':8, 'layer':9, 'add_y':10, 'animation':11, 'coloring':12}
    # key_index = {**{'id':1}, **key_index}
    # data_attribute = list(key_index.keys())
    # type_list = ['str', 'int', 'int', 'str', 'str', 'int', 'int', 'int', 'int', 'str', 'str']
    # type_list = ['int']+type_list
    # row_m = len(data_attribute)
    # column_m = len(material)
    wb.create_sheet(title=sheet_name,index=-1)
    ws = wb[sheet_name]

    ws.cell(row=1, column=1, value=data_attribute[0])
    ws.cell(row=2, column=1, value=type_list[0])
    for i, id in enumerate(data):
        # print(i,id)
        ws.cell(row=i+3, column=1, value=id)
        item = data[id]
        for j in item:
            ws.cell(row=1, column=key_index[j], value=data_attribute[key_index[j]-1])
            ws.cell(row=2, column=key_index[j], value=type_list[key_index[j]-1])
            if key_index[j] == 1:
                c_value = int(id)
            else:
                c_value = item.get(j) 
            try:           
                ws.cell(row=i+3, column=key_index[j], value=c_value)
            except ValueError:
                ws.cell(row=i+3, column=key_index[j], value=str(c_value))

    wb.save(path)
if __name__ == "__main__":
    path = "./Excel/json2excel.xlsx"
    material = {"81001": {'series': 'fruit', "id":666,'index': 28, 'score': 120, 'color': 'colourful', 'temperature': 'ruan', 'sweetness': 15, 'calorie': 10, 'layer': 4, 'add_y': 15, 'animation': 'bc_cake_front'},
        "81002": {'series': 'candy', 'it':222,'index': 27, 'score': 120, 'color': 'colourful', 'temperature': 'cui', 'sweetness': 20, 'calorie': 35, 'layer': 3, 'add_y': 25, 'animation': 'bc_cake_center'},
        "81003": {'series': 'fruit', "id":666,'index': 28, 'score': 120, 'color': 'colourful', 'temperature': 'ruan', 'sweetness': 15, 'calorie': 10, 'layer': 4, 'add_y': 15, 'animation': 'bc_cake_front'}}
    json2excel(path, material, "material")
