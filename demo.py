key_index = {
    'index': 1,
    'series': 2,
    'score': 3,
    'color': 4,
    'temperature': 5,
    'sweetness': 6,
    'calorie': 7,
    'layer': 8,
    'add_y': 9,
    'animation': 10
}

def handle_data(data):
    key_list1 = list(list(data.values())[0].keys())
for i, id in enumerate(data):
    # i是key的index，id是key
    # print(i,id)
    item = data[id]
    # print(data.values())
    
    key_list = list(item.keys())
    # 如果字典第一组key比后续的少
    for j , key in enumerate(key_list):
        if key not in key_list1:
            key_list1.append(key)
    # 如果字典第一组key比后续的多
    for j , key in enumerate(key_list1):
        # j是key的index，key是key
        try:
            item[key]
        except KeyError:
            # key_list1.append(key)
            item[key] = None


    print(key_list1)
    # return data
data = {"81001": {'series': 'fruit', "id":666,'index': 28, 'score': 120, 'color': 'colourful', 'temperature': 'ruan', 'sweetness': 15, 'calorie': 10, 'layer': 4, 'add_y': 15, 'animation': 'bc_cake_front'},
        "81002": {'series': 'candy', 'it':222,'index': 27, 'score': 120, 'color': 'colourful', 'temperature': 'cui', 'sweetness': 20, 'calorie': 35, 'layer': 3, 'add_y': 25, 'animation': 'bc_cake_center'},
        "81003": {'series': 'fruit', "id":666,'index': 28, 'score': 120, 'color': 'colourful', 'temperature': 'ruan', 'sweetness': 15, 'calorie': 10, 'layer': 4, 'add_y': 15, 'animation': 'bc_cake_front'}}
key_list1 = list(list(data.values())[0].keys())
for i, id in enumerate(data):
    # i是key的index，id是key
    # print(i,id)
    item = data[id]
    # print(data.values())
    
    key_list = list(item.keys())
    # 如果字典第一组key比后续的少
    for j , key in enumerate(key_list):
        if key not in key_list1:
            key_list1.append(key)
    # 如果字典第一组key比后续的多
    for j , key in enumerate(key_list1):
        # j是key的index，key是key
        try:
            item[key]
        except KeyError:
            # key_list1.append(key)
            item[key] = None


    print(key_list1)
    # for k in item:
    #     print(item[k])
    #     sheet.write(i+1, key_index[k], item[k])
