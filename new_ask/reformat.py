import os
from formula import change_the_formula



def creat_shop_dict(counts, example_sheet, final_sheet):  # 创建商家字典
    dict_shop = {}
    for i in range(2, counts + 1):
        if (example_sheet.cell(row=i, column=1).value and example_sheet.cell(row=i, column=1).value != '饮料分区'):
            dict_shop[example_sheet.cell(row=i, column=1).value] = i
        for j in range(2, 6):
            final_sheet.cell(row=i, column=j).value = None
        for c in range(9, 12):
            final_sheet.cell(row=i, column=c).value = None

    return dict_shop, final_sheet


def change_name(counts, sample_sheet):
    for i in range(2, counts+10):
        if sample_sheet.cell(row=i, column=1).value:
            if sample_sheet.cell(row=i, column=1).value == '彩虹绵绵冰':
                sample_sheet.cell(row=i, column=1).value = '杨小贤 芒果绵绵冰'
    return sample_sheet


def shop_move_pos(counts, example_sheet, final_sheet, sample_sheet, food):  # 排版
    dict_shop, final_sheet = creat_shop_dict(counts, example_sheet, final_sheet)
    sample_sheet.move_range('F1:F100', 0, 17, True)
    sample_sheet.move_range('P1:P100', 0, 6, True)
    sample_sheet.move_range('I1:I100', 0, 16, True)
    sample_sheet.move_range('Q1:Q100', 0, 15, True)
    sample_sheet.move_range('H1:H100', 0, 21, True)
    sample_sheet.delete_cols(2, 20)
    counts, food, final_sheet, sample_sheet, dict_shop = insert_new_shop(counts, food, final_sheet, sample_sheet,
                                                                         dict_shop)
    for i in range(2, counts):
        s = 'A' + str(i) + ':AZ' + str(i)
        sample_sheet.move_range(s, rows=300, cols=0, translate=True)
    for i in range(302, counts + 320):
        s = 'A' + str(i) + ':AZ' + str(i)
        if sample_sheet.cell(row=i, column=1).value:
            sample_sheet.move_range(s, rows=dict_shop[sample_sheet.cell(row=i, column=1).value] - i, cols=0)

    return sample_sheet, final_sheet



data_directory = os.path.dirname(os.path.abspath(__file__))
data_directory = data_directory.strip('\\new_ask')


def insert_new_shop(count_drink, count_food, be_change, change_u, dict_shop):
    for r in range(2, count_drink+10):
        if change_u.cell(row=r, column=1).value:
            if change_u.cell(row=r, column=1).value in dict_shop:
                continue
            else:
                print('------------->', change_u.cell(row=r, column=1).value)
                shop_type = int(input('侦测到新商家，饮品类输入1，食品类输入2：'))
                if shop_type == 2:
                    dict_shop[change_u.cell(row=r, column=1).value] = count_food + 1
                    count_food += 1
                    count_drink += 1
                    shop_money_type = int(input('协议价输入1，15趴类输入2:'))
                    if shop_money_type == 1:
                        be_change.insert_rows(count_food)
                        be_change.cell(row=count_food, column=1).value = change_u.cell(row=r, column=1).value
                        s = '=B' + str(count_food)
                        be_change.cell(row=count_food, column=6).value = s
                        s = '=(F' + str(count_food) + '-I' + str(count_food) + ')-L' + str(count_food)
                        be_change.cell(row=count_food, column=7).value = s
                        s = '=F' + str(count_food) + '*0.15'
                        be_change.cell(row=count_food, column=8).value = s
                        s = '=L' + str(count_food) + '-H' + str(count_food)
                        be_change.cell(row=count_food, column=13).value = s
                        s = '=C' + str(count_food) + '/20'
                        be_change.cell(row=count_food, column=14).value = s
                        s = '=(H' + str(count_food) + '+M' + str(count_food) + ')/F' + str(count_food)
                        be_change.cell(row=count_food, column=15).value = s
                        with open(data_directory + "/ask/毛利润商家列表.txt", 'a', encoding='UTF-8') as f:
                            f.write('\n' + change_u.cell(row=r, column=1).value)

                    if shop_money_type == 2:
                        be_change.insert_rows(count_food)
                        be_change.cell(row=count_food, column=1).value = change_u.cell(row=r, column=1).value
                        s = '=B' + str(count_food)
                        be_change.cell(row=count_food, column=6).value = s
                        s = '=(F' + str(count_food) + '-I' + str(count_food) + ')*0.85'
                        be_change.cell(row=count_food, column=7).value = s
                        s = '=F' + str(count_food) + '*0.15'
                        be_change.cell(row=count_food, column=8).value = s
                        s = '=C' + str(count_food) + '/20'
                        be_change.cell(row=count_food, column=14).value = s
                        s = '=H' + str(count_food) + '/F' + str(count_food)
                        be_change.cell(row=count_food, column=15).value = s

                    be_change = change_the_formula(count_drink, count_food, be_change, 0)



                else:
                    dict_shop[change_u.cell(row=r, column=1).value] = count_food + 6
                    a = count_food
                    count_food = count_food + 6
                    count_drink += 1
                    shop_money_type = int(input('协议价输入1，15趴类输入2:'))
                    if shop_money_type == 1:
                        be_change.insert_rows(count_food)
                        be_change.cell(row=count_food, column=1).value = change_u.cell(row=r, column=1).value
                        s = '=B' + str(count_food)
                        be_change.cell(row=count_food, column=6).value = s
                        s = '=(F' + str(count_food) + '-I' + str(count_food) + ')-L' + str(count_food)
                        be_change.cell(row=count_food, column=7).value = s
                        s = '=F' + str(count_food) + '*0.15'
                        be_change.cell(row=count_food, column=8).value = s
                        s = '=L' + str(count_food) + '-H' + str(count_food)
                        be_change.cell(row=count_food, column=13).value = s
                        s = '=C' + str(count_food) + '/20'
                        be_change.cell(row=count_food, column=14).value = s
                        s = '=(H' + str(count_food) + '+M' + str(count_food) + ')/F' + str(count_food)
                        be_change.cell(row=count_food, column=15).value = s
                        with open(data_directory + "/ask/毛利润商家列表.txt", 'a', encoding='UTF-8') as f:
                            f.write('\n' + change_u.cell(row=r, column=1).value)
                    if shop_money_type == 2:
                        be_change.insert_rows(count_food)
                        be_change.cell(row=count_food, column=1).value = change_u.cell(row=r, column=1).value
                        s = '=B' + str(count_food)
                        be_change.cell(row=count_food, column=6).value = s
                        s = '=(F' + str(count_food) + '-I' + str(count_food) + ')*0.85'
                        be_change.cell(row=count_food, column=7).value = s
                        s = '=F' + str(count_food) + '*0.15'
                        be_change.cell(row=count_food, column=8).value = s
                        s = '=C' + str(count_food) + '/20'
                        be_change.cell(row=count_food, column=14).value = s
                        s = '=H' + str(count_food) + '/F' + str(count_food)
                        be_change.cell(row=count_food, column=15).value = s
                    count_food = a
                    be_change = change_the_formula(count_drink, count_food, be_change, 1)

    dict_new_shop = {}
    for i in range(2, count_drink + 1):
        if (be_change.cell(row=i, column=1).value and be_change.cell(row=i, column=1).value != '饮料分区'):
            dict_new_shop[be_change.cell(row=i, column=1).value] = i
    dict_shop = dict_new_shop
    return count_drink, count_food, be_change, change_u, dict_shop

def read_shop_list_profit(data_directory):
    profit_shop_list = []
    with open(data_directory + "/ask/毛利润商家列表.txt", encoding='UTF-8', errors='ignore') as f:
        for line in f:
            profit_shop_list.append(line.strip('\n'))
        profit_shop_list = set(profit_shop_list)
    return profit_shop_list




