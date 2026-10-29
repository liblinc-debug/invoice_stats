import unittest
from datetime import date
from decimal import Decimal

from invoice_stats import parse_amount, parse_date, parse_invoice_number, parse_item_name


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


if __name__ == "__main__":
    unittest.main()
