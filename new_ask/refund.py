def refund(Count_shop_new, sample_sheet):  # 退款转换
    recall = 0
    for i in range(2, Count_shop_new):
        for j in range(len(sample_sheet.cell(row=i, column=8).value)):
            if sample_sheet.cell(row=i, column=8).value[j] == '：':
                a = sample_sheet.cell(row=i, column=8).value[j + 1:]
                sample_sheet.cell(row=i, column=8).value = a
                if a:
                    a = a[1:]
                    recall = recall + float(a)
                    sample_sheet.cell(row=i, column=8).value = round(float(a), 2)
                    break
    return sample_sheet


def s_define(sheet, count):  # 自定义计算
    sum_s_define = 0
    g_s_define = 0
    uncle_s_define = 0
    sf_s_define = 0
    for r in range(2, count + 10):
        if sheet.cell(row=r, column=11).value:
            if sheet.cell(row=r, column=1).value == '阿叔猪扒包':
                uncle_s_define = round(eval(sheet.cell(row=r, column=11).value), 2)
                sum_s_define = sum_s_define + uncle_s_define + g_s_define
            if sheet.cell(row=r, column=1).value == '居肉町·极炙烧肉饭':
                g_s_define = round(eval(sheet.cell(row=r, column=11).value), 2)
                sum_s_define = sum_s_define + g_s_define + uncle_s_define
            if sheet.cell(row=r, column=1).value == '丰顺捆粄（客家小吃）':
                sf_s_define = round(eval(sheet.cell(row=r, column=11).value), 2)
                sum_s_define = sum_s_define + g_s_define + uncle_s_define + sf_s_define
            else:
                sum_s_define = sum_s_define + round(eval(sheet.cell(row=r, column=11).value), 2)
    return round(sum_s_define, 2), uncle_s_define, g_s_define, sf_s_define


def refund_combine(counts, final_sheet):  # 退款合拼
    for r in range(2, counts + 10):
        if final_sheet.cell(row=r, column=1).value and final_sheet.cell(row=r + 1, column=1).value:
            a = final_sheet.cell(row=r, column=1).value
            b = final_sheet.cell(row=r + 1, column=1).value
            if final_sheet.cell(row=r, column=1).value[0] == final_sheet.cell(row=r + 1, column=1).value[0]:
                if final_sheet.cell(row=r, column=9).value and final_sheet.cell(row=r + 1, column=9).value:
                    final_sheet.cell(row=r + 1, column=9).value = final_sheet.cell \
                        (row=r, column=9).value + final_sheet.cell(row=r + 1, column=9).value
                    final_sheet.cell(row=r, column=9).value = None
                elif final_sheet.cell(row=r, column=9).value:
                   final_sheet.cell(row=r+1, column=9).value = final_sheet.cell(row=r, column=9).value
                   final_sheet.cell(row=r, column=9).value = None
        else:
            continue
    return final_sheet


def profit_combine(counts, final_sheet):
    for r in range(2, counts + 10):
        if final_sheet.cell(row=r, column=12).value and final_sheet.cell(row=r + 1, column=12).value:
            if final_sheet.cell(row=r, column=1).value[0] == final_sheet.cell(row=r + 1, column=1).value[0]:
                final_sheet.cell(row=r + 1, column=12).value = \
                    final_sheet.cell(row=r, column=12).value + final_sheet.cell(row=r + 1, column=12).value
                final_sheet.cell(row=r, column=12).value = None
        else:
            continue

    return final_sheet


def clean_profit(counts, profit_shop_list, sheet, output_s):
    for r in range(2, counts + 10):
        if sheet.cell(row=r, column=12).value:
            if sheet.cell(row=r, column=1).value in profit_shop_list:
                output_s.cell(row=r, column=12).value = None
                continue
            else:
                sheet.cell(row=r, column=12).value = None
    return sheet, output_s
