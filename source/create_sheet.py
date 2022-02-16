import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

def create_excel(data_dir, data, symbol):
    workbook = xlsxwriter.Workbook(data_dir + '/' + symbol + '.xlsx')
    worksheet = workbook.add_worksheet()

    million_format = workbook.add_format({'num_format': ' #,##0.0,, "M";(#,##0.0,, "M")'})
    percent_format = workbook.add_format({'num_format': '#,##0.0%_);(#,##0.0%);-_);@_)'})
    bold = workbook.add_format({'bold': True})
    result = workbook.add_format({'num_format': ' #,##0.0,, "M";(#,##0.0,, "M")', 'bold': True, 'top': 2})

    worksheet.write(2, 0, 'Q', bold)
    worksheet.write(3, 0, 'Revenue', bold)
    worksheet.write(4, 0, 'Cost of Revenue', bold)
    worksheet.write(5, 0, 'Gross Profit', result)
    worksheet.write(6, 0, 'Gross Profit Ratio', bold)
    worksheet.write(8, 0, 'Expenses', bold)
    worksheet.write(9, 0, 'Research & Development', bold)
    worksheet.write(10, 0, 'Selling and Marketing', bold)
    worksheet.write(11, 0, 'General & Admin', bold)
    worksheet.write(12, 0, 'Other', bold)
    worksheet.write(13, 0, 'Total OpEx', result)
    worksheet.write(14, 0, 'Operating loss', result)
    worksheet.write(15, 0, 'Operating margin', bold)
    worksheet.write(17, 0, 'EPS', bold)
    worksheet.write(19, 0, 'EBITDA', bold)
    
    row = 0
    col = 1

    for x in reversed(data['income']):
        worksheet.set_column(col, col, 10)
        worksheet.write(row + 2, col, x['date'])
        worksheet.write(row + 3, col, x['revenue'], million_format)
        worksheet.write(row + 4, col, x['costOfRevenue'], million_format)
        worksheet.write(row + 5, col, x['grossProfit'], result)
        worksheet.write(row + 6, col, x['grossProfitRatio'], percent_format)
        worksheet.write(row + 9, col, x['sellingAndMarketingExpenses'], million_format)
        worksheet.write(row + 10, col, x['generalAndAdministrativeExpenses'], million_format)
        worksheet.write(row + 11, col, x['researchAndDevelopmentExpenses'], million_format)
        worksheet.write(row + 12, col, x['otherExpenses'], million_format)
        worksheet.write(row + 13, col, x['operatingExpenses'], result)


        # Calculate OpEx
        cell_for_opex = xl_rowcol_to_cell(row + 14, col)
        gross_profit = xl_rowcol_to_cell(row + 5, col)
        op_ex = xl_rowcol_to_cell(row + 13, col)
        worksheet.write_dynamic_array_formula('%s' % cell_for_opex, '=%s-%s' %(gross_profit, op_ex), result)
        # Add operating margin
        cell_for_op_margin = xl_rowcol_to_cell(row + 15, col)
        revenue = xl_rowcol_to_cell(row + 3, col)
        worksheet.write_dynamic_array_formula('%s' % cell_for_op_margin, '=%s/%s' %(cell_for_opex, revenue), percent_format)

        worksheet.write(row + 17, col, x['eps'])
        worksheet.write(row + 19, col, x['ebitda'], million_format)

        col += 1

    workbook.close()