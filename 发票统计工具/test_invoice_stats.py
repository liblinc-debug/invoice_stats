import unittest
from datetime import date
from decimal import Decimal

from invoice_stats import normalize_pdf_text, parse_amount, parse_date, parse_invoice_number, parse_item_name


class InvoiceParserTest(unittest.TestCase):
    def test_standard_invoice(self):
        text = """电子发票（普通发票） 发票号码：26952000000168054676
开票日期：2026年01月15日
项目名称 规格型号 单位 数量 单价 金额 税率 税额
*计算机外部设备*摄像头模组 S1M01 个 1 45.25 45.25 1% 0.45
价税合计（小写） ¥45.70"""
        self.assertEqual(parse_invoice_number(text), "26952000000168054676")
        self.assertEqual(parse_date(text), date(2026, 1, 15))
        self.assertEqual(parse_item_name(text), "*计算机外部设备*摄像头模组 S1M01")
        self.assertEqual(parse_amount(text), Decimal("45.70"))

    def test_spaced_amount(self):
        text = "价税合计（小写） ¥ 1 9 0 . 0 0"
        self.assertEqual(parse_amount(text), Decimal("190.00"))

    def test_toll_invoice_with_broken_font_mapping(self):
        raw = """项目名称 车牌号 车辆类型 通行日期起 通行日期止 金额 税率/征收率 税额
通行费
\x002\x006\x004\x004\x007\x009\x000\x002\x001\x008\x001\x000\x000\x000\x006\x006\x009\x001\x008\x005
\x002\x000\x002\x006^t\x001\x000g\x08\x002\x009eå
\x00*u\x1fN§u\x1fm;g\rR¡\x00*\x90\x1a\x88L\x8d9
\x00¥\x002\x008.\x008\x007
\x00¥\x000.\x008\x007
\x00¥\x002\x009.\x007\x004
"""
        text = normalize_pdf_text(raw)
        self.assertEqual(parse_invoice_number(text), "26447902181000669185")
        self.assertEqual(parse_date(text), date(2026, 10, 29))
        self.assertEqual(parse_item_name(text), "通行费")
        self.assertEqual(parse_amount(text), Decimal("29.74"))


if __name__ == "__main__":
    unittest.main()
