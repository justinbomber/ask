import os


def shop_move_pos(counts, example_sheet, final_sheet, sample_sheet):  # 排版
    dict_shop, final_sheet = create_shop_dict(counts, example_sheet, final_sheet)
    sample_sheet.move_range('F1:F100', 0, 19, True)
    sample_sheet.move_range('P1:P100', 0, 10, True)
    sample_sheet.move_range('I1:I100', 0, 14, True)
    sample_sheet.move_range('Q1:Q100', 0, 12, True)
    sample_sheet.delete_cols(2, 21)
    for i in range(2, counts):
        s = 'A' + str(i) + ':AZ' + str(i)
        sample_sheet.move_range(s, rows=300, cols=0, translate=True)
    for i in range(302, counts + 320):
        s = 'A' + str(i) + ':AZ' + str(i)
        if sample_sheet.cell(row=i, column=1).value:
            sample_sheet.move_range(s, rows=dict_shop[sample_sheet.cell(row=i, column=1).value] - i, cols=0)

    return sample_sheet, final_sheet


def create_shop_dict(counts, example_sheet, final_sheet):  # 创建商家字典
    dict_shop = {}
    for i in range(2, counts + 1):
        if example_sheet.cell(row=i, column=1).value and example_sheet.cell(row=i, column=1).value != '总计':
            dict_shop[example_sheet.cell(row=i, column=1).value] = i
            for j in range(2, 6):
                final_sheet.cell(row=i, column=j).value = None
    return dict_shop, final_sheet


def read_shop_list_profit(data_directory):
    profit_shop_list = []
    with open(data_directory + "/毛利润商家列表早餐.txt", encoding='UTF-8', errors='ignore') as f:
        for line in f:
            profit_shop_list.append(line.strip('\n'))
        profit_shop_list = set(profit_shop_list)
    return profit_shop_list
