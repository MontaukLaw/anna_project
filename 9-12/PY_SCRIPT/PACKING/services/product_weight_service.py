"""按资料行的装箱方式选取重量列，界面与导出共用。"""
import re


def product_weights(values):
    name = str(values.get('产品名称') or '')
    note = str(values.get('备注') or '')
    pallet = '卡板装' in name or bool(re.search(
        r'(?:一|1|每)\s*(?:卡板?|托盘?|板)\s*装\s*\d+\s*箱|卡板\s*\d+\s*箱', note))
    normalized = {re.sub(r'\s+|的', '', str(key)).casefold(): value for key, value in values.items()}
    prefix = '整个卡板' if pallet else '整箱'
    return (normalized.get(f'{prefix}净重kg'), normalized.get(f'{prefix}毛重kg'), prefix)
