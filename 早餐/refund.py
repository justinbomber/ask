def refund(count_shop, change_u):  # 退款转换
    recall = 0
    for i in range(2, count_shop):
        for j in range(len(change_u.cell(row=i, column=8).value)):
            if change_u.cell(row=i, column=8).value[j] == '：':
                a = change_u.cell(row=i, column=8).value[j + 1:]
                change_u.cell(row=i, column=8).value = a
                if a:
                    a = a[1:]
                    recall = recall + float(a)
                    change_u.cell(row=i, column=8).value = round(float(a), 2)
                    break
    return recall


def clean_profit(counts, profit_shop_list, sheet, output_s):
    for r in range(2, counts + 10):
        if sheet.cell(row=r, column=8).value:
            if sheet.cell(row=r, column=1).value in profit_shop_list:
                output_s.cell(row=r, column=8).value = None
                continue
            else:
                sheet.cell(row=r, column=8).value = None
    return sheet, output_s
