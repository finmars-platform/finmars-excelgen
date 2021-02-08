#! /usr/bin/env python

import sys
import os
import time

import json
import yaml
from io import BytesIO


from werkzeug.wsgi import wrap_file
from werkzeug.wrappers import Request, Response
from executor import execute

from werkzeug.datastructures import Headers


from pyexcelerate import Workbook, Color, Style, Font, Fill, Format


import logging

logger = logging.getLogger()

gunicorn_error_logger = logging.getLogger('gunicorn.error')
logger.handlers.extend(gunicorn_error_logger.handlers)
logger.setLevel(logging.DEBUG)



@Request.application
def application(request):

    if request.method == 'GET':
        return Response("Excel generator is online")

    if request.method != 'POST':
        return

    st = time.perf_counter()

    # logger.info("Request data %s" % request.data)

    # print('request.data %s' %request.data)

    # requestData = json.loads(request.data)

    requestData = yaml.safe_load(request.data)

    contentSettings = requestData['contentSettings']
    content = requestData['content']

    print("Creating empty workbook")
    wb = Workbook()
    ws = wb.new_sheet("General")

    columns = contentSettings["columns"]

    # colorFill = PatternFill(start_color='F05A22',
    #                         end_color='F05A22',
    #                         fill_type='solid')
    #
    # thin_border = Border(left=Side(style='thin'),
    #                      right=Side(style='thin'),
    #                      top=Side(style='thin'),
    #                      bottom=Side(style='thin'))
    #
    # ft = Font(color="FFFFFF", name='Arial', size=14)

    column_index = 0
    for column in columns:
        ws[1][column_index + 1].value = column['name']
        ws[1][column_index + 1].style.font.color = Color(255, 255, 255)
        ws[1][column_index + 1].style.fill.background = Color(240, 90, 34)
        column_index = column_index + 1


    # for col in ws.iter_cols(min_row=1, max_col=len(columns), max_row=1):
    #     for cell in col:
    #         cell.fill = colorFill
    #         cell.font = ft
    #         cell.border = thin_border

    print("Adding data")

    row_index = 1 # first for header

    for row in content:

        column_index = 0

        skip = False

        if '___type' in row and row['___type'] == 'subtotal':
            skip = True

        if not skip:

            for column in columns:

                value = None

                try:

                    value = row[column["key"]]

                except Exception as e:
                    value = None

                ws[row_index + 1][column_index + 1] = value

                column_index = column_index + 1

            row_index = row_index + 1


    print("Saving to excel file")

    # wb.save(filename='report.xlsx')

    headers = Headers()
    headers.add('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    headers.add('Content-Disposition', 'attachment', filename='report.xslx')

    # file = open('report.xslx')
    #
    # os.remove('report.xslx')

    stream = BytesIO()
    wb.save(stream)

    response = Response(stream.getvalue(),
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', headers=headers
                        )

    print('generating excel done: %s', "{:3.3f}".format(time.perf_counter() - st))

    # file.close()

    return response


if __name__ == '__main__':
    from werkzeug.serving import run_simple

    run_simple(
        '127.0.0.1', 5000, application, use_debugger=True, use_reloader=True
    )
