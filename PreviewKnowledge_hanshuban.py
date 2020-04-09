import xlrd


def read_xlrd(excelFile, colNum):
    """
    :param excelFile: 输入的excel文档
    :param colNum: 飞机型号所在的列号
    :return: dataFile整体为一个列表，而每一行的飞机型号集合为一子列表
            table为读取的excel表
    """
    data = xlrd.open_workbook(excelFile)
    table = data.sheet_by_index(0)
    dataFile = []

    for rowNum in range(table.nrows):
        if rowNum > 0:
            dataFile.append(table.cell(rowNum, colNum).value.split(";"))

    return table, dataFile

def make_dict(names, values):
    """
    :param names: 网络的输出的飞机列表集合
    :param values: 与飞机集合一一对应的网络预测概率值
    :return:将两者统一起来形成一个字典
    """
    nvs = zip(names, values)
    nvDict = dict((name, value) for name, value in nvs)

    return nvDict

def max_dict(intersection, nvDict):
    """
    :param intersection: 其数据形式是一个集合，是网络输出的列表名称与先验知识列表的一个集合
    :param nvDict: 是网络输出的列表名称与其预测数值所形成的字典
    :return: 返回值是一个字典，是将先验知识与网络输出的飞机列表的交集intersection,对应的字典中值
    最大的键的输出
    """
    p2 = {k: v for k, v in nvDict.items() if k in intersection}
    return max(p2, key=p2.get)


def process_previewknowledge(Longtitude, Atitude, longtitude, atitude, excelFile, colNum,
                             names, values):

    table, dataFile =read_xlrd(excelFile, colNum)
    nvDict = make_dict(names, values)
    if longtitude == table.cell(1, Longtitude).value and atitude == table.cell(1, Atitude).value:
        previewknowledge = dataFile[1-1]
        intersection = set(previewknowledge) & set(names)
        #先验机型和预测的机型无交集，按照预测的最大值输出
        if intersection == set():
            return print(max(nvDict, key=nvDict.get))

        else:

            return print(max_dict(intersection, nvDict))

    elif longtitude == table.cell(2, Longtitude).value and atitude == table.cell(2, Atitude).value:
        previewknowledge = dataFile[2 - 1]
        intersection = set(previewknowledge) & set(names)
        # 先验机型和预测的机型无交集，按照预测的最大值输出
        if intersection == set():
            return print(max(nvDict, key=nvDict.get))

        else:
            return print(max_dict(intersection, nvDict))

    elif longtitude == table.cell(3, Longtitude).value and atitude == table.cell(3, Atitude).value:
        previewknowledge = dataFile[3 - 1]
        intersection = set(previewknowledge) & set(names)
        # 先验机型和预测的机型无交集，按照预测的最大值输出
        if intersection == set():
            return print(max(nvDict, key=nvDict.get))

        else:
            return print(max_dict(intersection, nvDict))

    elif longtitude == table.cell(3, Longtitude).value and atitude == table.cell(3, Atitude).value:
        previewknowledge = dataFile[3 - 1]
        intersection = set(previewknowledge) & set(names)
        # 先验机型和预测的机型无交集，按照预测的最大值输出
        if intersection == set():
            return print(max(nvDict, key=nvDict.get) )
        else:
            return print(max_dict(intersection, nvDict))


if __name__ == '__main__':
    #以下变量均是需要由网络输入,name是网络输入的飞机类别的集合，
    # values是其一一对应的值，应该是从网络反馈回来的向量，
    # nvDict是由这两个列表生成的字典
    names = ['An-12', 'B-1B', 'B-52H', 'EF-55', 'EF-67', 'EF-12', 'C-17', 'C-130', 'CH-47', 'E-3',
             'E-8', 'IL-76', 'KC-10', 'KC-135', 'P-3', 'F-22', 'J-20']
    values = [0.05, 0.1, 0.05, 10, 0.05, 0.05, 0.05, 0.05, 0.05, 0.05, 0.05, 0.05,
              0.05, 0.05, 0.05, 0, 0.5]
    #longtitude, atitude为输入切片的经度值和纬度值
    longtitude, atitude = [-147.4855, 19.2429]
    excelFile = 'test.xlsx'
    colNum=4
    #excel表格中的经度Longtitude,和纬度Atitude所在的列数
    Longtitude = 2
    Atitude = 3
    prediction = process_previewknowledge(Longtitude=Longtitude, Atitude=Atitude,longtitude=longtitude,
                                          atitude=atitude, excelFile=excelFile,
                                          colNum=colNum, names=names,values=values)

