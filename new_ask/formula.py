from itertools import groupby



def change_formula(num_more, formula, SUM):
    ss = [''.join(list(g)) for k, g in groupby(formula, key=lambda x: x.isdigit())]
    print(ss)
    for i in range(len(ss)):
        if ss[i].isdigit():
            if ss[i - 1][-1] == '+' or ss[i - 1][-1] == '*' or ss[i - 1][-1] == '-' or ss[i - 1][-1] == '/':
                break
            elif SUM:
                ss[-2] = int(ss[-2])
                ss[-2] += num_more
                ss[-2] = str(ss[-2])
                break

            else:
                ss[i] = int(ss[i])
                ss[i] += num_more
                ss[i] = str(ss[i])
        else:
            continue
    aa = ''.join(ss)
    return aa


def eval_all(counts, sheet):
    for r in range(2, counts+10):
        for c in range(2, 13):
            if c == 9:
                continue
            elif sheet.cell(row=r, column=c).value:
                sheet.cell(row=r, column=c).value = round(eval(sheet.cell(row=r, column=c).value), 2)
    return sheet


title_set = {'手续费', '退款', '运费', '毛利', '差额结余'}
title_set2 = {'手续费', '退款', '运费', '毛利', '主食订单数', '差额结余'}

def change_the_formula(count_drink, count_food, be_change, drink_bool):
    if drink_bool:
        for i in range(count_food+7, count_food+count_drink):
            for j in range(1, 25):
                if be_change.cell(row=i, column=j).value and isinstance(be_change.cell(row=i, column=j).value, str) and be_change.cell(row=i, column=j).value[0] == '=':
                    if be_change.cell(row=i, column=j - 1).value in title_set or i == count_drink+2:
                        be_change.cell(row=i, column=j).value = change_formula(1, be_change.cell(row=i, column=j).value, 1)
                    elif be_change.cell(row=i, column=j - 1).value == '饮料杯数':
                        s = '=SUM(C' + str(count_food+5) + ':C' + str(count_drink-1) + ')'
                        be_change.cell(row=i, column=j).value = s
                    elif be_change.cell(row=i, column=j - 1).value == '主食订单数':
                        continue
                    else:
                        be_change.cell(row=i, column=j).value = change_formula(1, be_change.cell(row=i, column=j).value, 0)

    else:
        for i in range(count_food+1, count_food+count_drink):
            for j in range(1, 25):
                if be_change.cell(row=i, column=j).value and isinstance(be_change.cell(row=i, column=j).value, str) and be_change.cell(row=i, column=j).value[0] == '=':
                    if be_change.cell(row=i, column=j - 1).value in title_set2 or i == count_drink+2:
                        be_change.cell(row=i, column=j).value = change_formula(1, be_change.cell(row=i, column=j).value, 1)
                    else:
                        be_change.cell(row=i, column=j).value = change_formula(1, be_change.cell(row=i, column=j).value, 0)
    return be_change





