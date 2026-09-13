import unittest
from services.product_weight_service import product_weights


class ProductWeightTests(unittest.TestCase):
    def test_pallet_name_and_notes(self):
        for name, note in [('振动球（卡板装）', ''), ('球', '一卡装180箱'),
                           ('球', '1卡板装90箱'), ('球', '卡板140箱')]:
            values = {'产品名称': name, '备注': note, '整箱净重kg': .556,
                      '整箱毛重kg': .716, '整个卡板的净重KG': 100.08,
                      '整个卡板毛重KG': 130.88}
            self.assertEqual(product_weights(values), (100.08, 130.88, '整个卡板'))

    def test_carton_and_missing_pallet_weights(self):
        values = {'产品名称': '球', '备注': '一箱装1个', '整箱净重kg': .556, '整箱毛重kg': .716}
        self.assertEqual(product_weights(values), (.556, .716, '整箱'))
        values['产品名称'] = '球（卡板装）'
        self.assertEqual(product_weights(values), (None, None, '整个卡板'))
