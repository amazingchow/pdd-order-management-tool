# -*- coding: utf-8 -*-
from __future__ import annotations
import os
import glob
import pathlib
import xlrd
import xlwt


if __name__ == "__main__":
    src_dir = "{}/Desktop/发货专用-待转换的表格".format(os.path.expanduser("~"))
    if not os.path.exists(src_dir):
        raise Exception("请先创建文件夹'{}'".format(src_dir))
    dst_dir = "{}/Desktop/发货专用-转换后的表格".format(os.path.expanduser("~"))
    if not os.path.exists(dst_dir):
        raise Exception("请先创建文件夹'{}'".format(dst_dir))
    
    for fpath in glob.glob(src_dir + "/*"):
        fname = pathlib.Path(fpath).stem
        fext = pathlib.Path(fpath).suffix

        wb_w = xlwt.Workbook()
        sheet_w = wb_w.add_sheet("待发货信息")
        header_font = xlwt.Font()
        header_font.name = "Arial"
        header_font.bold = True
        header_style = xlwt.XFStyle()
        header_style.font = header_font
        sheet_w.write(0, 0, "订单编号", header_style)
        sheet_w.write(0, 1, "收件人", header_style)
        sheet_w.write(0, 2, "手机", header_style)
        sheet_w.write(0, 3, "地址", header_style)
        sheet_w.write(0, 4, "发货信息", header_style)
        sheet_w.write(0, 5, "商品ID", header_style)
        sheet_w.write(0, 6, "发货数量", header_style)
        sheet_w.write(0, 7, "备注", header_style)

        wb_r = xlrd.open_workbook(fpath)
        sheet_r = wb_r.sheet_by_index(0)
        for row_idx in range(sheet_r.nrows):
            if len(sheet_r.cell_value(row_idx, 0)) == 0:
                continue

            if len(sheet_r.cell_value(row_idx, 0).split("\n")) >= 3:
                parts = sheet_r.cell_value(row_idx, 0).split("\n")
                sheet_w.write(row_idx + 1, 1, parts[0])
                sheet_w.write(row_idx + 1, 2, parts[1])
                if len(parts) > 3:
                    item = ""
                    for x in parts[2:]:
                        item += x
                    sheet_w.write(row_idx + 1, 3, item)
                else:
                    sheet_w.write(row_idx + 1, 3, parts[2])
            else:
                parts = sheet_r.cell_value(row_idx, 0).split(", ")
                sheet_w.write(row_idx + 1, 1, parts[1])
                sheet_w.write(row_idx + 1, 2, parts[2])
                if len(parts) > 3:
                    item = parts[0]
                    for x in parts[2:]:
                        item += x
                    sheet_w.write(row_idx + 1, 3, item)
                else:
                    sheet_w.write(row_idx + 1, 3, parts[0])

            item_num = sheet_r.cell_value(row_idx, 1)
            sheet_w.write(row_idx + 1, 6, item_num)

            annotation = sheet_r.cell_value(row_idx, 2)
            if len(annotation) > 0:
                sheet_w.write(row_idx + 1, 7, annotation)

        wb_w.save("{}/{}.xls".format(dst_dir, fname))
