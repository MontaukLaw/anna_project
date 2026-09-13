from decimal import Decimal
import unittest
from models.packing import OrderItem
from services.packing_recommendation_service import recommend_candidate


class RecommendationTests(unittest.TestCase):
    def setUp(self):
        self.item=OrderItem('33465',Decimal(120),Decimal(12),'1','Duck','',
            packaging='HEADER CARD WITH PDQ',packing='12 PCS / 12.5 X 8.75 X 8 INCH')

    def record(self,note='1PDQ 12袋',dims=(31.5,22.5,20.5)):
        return {'values':dict(zip(('长','宽','高'),dims),备注=note,装箱数量=12)}

    def test_both_match_with_two_percent_tolerance(self):
        self.assertEqual(recommend_candidate(self.item,self.record()).level,'strong')

    def test_two_percent_boundary_and_above(self):
        self.item.packing = '10 X 20 X 30 CM'
        for dims in [(10.2, 20.4, 30.6), (9.8, 19.6, 29.4)]:
            self.assertEqual(recommend_candidate(self.item,self.record(dims=dims), tolerance_percent=2).level,'strong')
        self.assertEqual(recommend_candidate(self.item,self.record(dims=(10.201,20,30)), tolerance_percent=2).level,'partial')

    def test_height_difference_only_packaging_matches(self):
        result=recommend_candidate(self.item,self.record(dims=(31.5,22.8,21)), tolerance_percent=2)
        self.assertEqual(result.level,'partial')
        self.assertIn('尺寸未匹配',result.reason)

    def test_configured_four_percent_matches_example(self):
        result = recommend_candidate(self.item, self.record(dims=(31.5,22.8,21)))
        self.assertEqual(result.level, 'strong')
        self.assertIn('误差≤4%', result.reason)

    def test_dimension_only_and_order_of_dimensions(self):
        self.assertEqual(recommend_candidate(self.item,self.record(note='散装')).level,'partial')
        self.assertEqual(recommend_candidate(self.item,self.record(note='散装',dims=(22.8,31.5,20.9))).level,'normal')

    def test_bag_and_bad_values(self):
        self.item.packaging='Polybag'
        self.assertEqual(recommend_candidate(self.item,self.record(note='胶袋装')).level,'strong')
        self.assertEqual(recommend_candidate(self.item,self.record(note='',dims=(None,'NaN',0))).level,'normal')
