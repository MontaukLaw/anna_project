"""装箱单批处理，独立于 GUI；选择有歧义的装箱记录由调用方负责。"""
from pathlib import Path
import re
from models.packing import PackingDataError, PackingRow
from services.order_pdf_service import read_order_items
from services.packing_lookup_service import PackingLookup, number
from services.packing_excel_service import build_packing_excel
from services.date_service import parse_schedule_date
from services.product_weight_service import product_weights
from services.slip_sheet_service import is_slip_item, build_slip_rows


def output_paths(matches, directory):
    paths = {}
    for order in matches:
        if not re.fullmatch(r"POHK-\d+-\d+-\d+", order, re.I):
            raise PackingDataError(f"订单号格式不正确：{order}")
        paths[order] = directory / f"{order}_装箱单.xlsx"
    return paths


def generate_packing_lists(matches, catalog, schedule, template, directory, choose, log, on_generated=None):
    paths = output_paths(matches, directory)
    lookup = None
    generated, failed, skipped = [], [], []
    selected_records = {}
    for order, pdfs in matches.items():
        try:
            if paths[order].exists():
                skipped.append(order)
                log(f"订单 {order} 的装箱单已存在，已跳过：{paths[order]}", "WARNING")
                continue
            if lookup is None:
                lookup = PackingLookup(catalog, schedule)
            if len(pdfs) != 1:
                raise PackingDataError(f"匹配到 {len(pdfs)} 个 PDF，需要唯一订单文件")
            log(f"开始读取订单 {order}：{pdfs[0].name}")
            items, customer = read_order_items(pdfs[0], order)
            log(f"订单 {order}：PDF 共 {len(items)} 个 ITEM，将逐项生成。")
            rows = []
            for item in items:
                code = lookup.date_code(order, item.item_no)
                try:
                    delivery_date = parse_schedule_date(item.cargo_date)
                except PackingDataError as exc:
                    raise PackingDataError(f"ITEM {item.item_no}：PDF Cargo Ready Date 为空或日期格式无法确定：{item.cargo_date}") from exc
                candidates = lookup.product_candidates(item.item_no)
                if is_slip_item(item):
                    rows.extend(build_slip_rows(item, candidates, code, delivery_date,
                        lookup.packaging_hint(order, item.item_no), choose, log))
                    continue
                key = (order, item.item_no, item.case_pack)
                if key in selected_records:
                    record, english = selected_records[key]
                else:
                    if len(candidates) > 1:
                        log(f"ITEM {item.item_no}：找到 {len(candidates)} 条装箱资料，请在弹窗中确认资料行。", "WARNING")
                        hint = lookup.packaging_hint(order, item.item_no)
                        log(f"订单排期表包装要求：\n{hint}")
                        choice = choose(item, candidates, hint)
                        if choice is None:
                            raise PackingDataError(f"用户取消选择 ITEM {item.item_no} 的装箱资料")
                        record, english = choice
                    else:
                        record = candidates[0]
                    selected_records[key] = (record, item.description)
                english = item.description
                v = record['values']
                log(f"ITEM {item.item_no}：使用 {record['sheet']} 第 {record['row']} 行；DATE CODE={code}；PDF Cargo Ready Date={delivery_date:%Y-%m-%d}；PDF DESCRIPTION={english}")
                if str(v.get('装箱数量')) != str(item.case_pack):
                    log(f"ITEM {item.item_no}：资料表装箱数量 {v.get('装箱数量')} 与 PDF {item.case_pack} 不同；按所选资料读取尺寸重量，箱数按 PDF 计算。", "WARNING")
                net, gross, weight_source = product_weights(v)
                log(f"ITEM {item.item_no}：重量来源={weight_source}，净重={net} KG，毛重={gross} KG")
                rows.append(PackingRow(item, code, number(net, f'{weight_source}净重'),
                    number(gross, f'{weight_source}毛重'), number(v.get('长'), '长'), number(v.get('宽'), '宽'),
                    number(v.get('高'), '高'), english, f"{catalog.source.name} / {record['sheet']} / 行 {record['row']}", delivery_date))
                if rows[-1].gross < rows[-1].net:
                    raise PackingDataError(f"ITEM {item.item_no} 毛重小于净重")
            content = build_packing_excel(template, order, customer, rows)
            directory.mkdir(parents=True, exist_ok=True)
            # Exclusive creation: never overwrite even if another process creates it after preflight.
            destination = paths[order]
            with destination.open('xb') as stream:
                try:
                    stream.write(content)
                except Exception:
                    stream.close()
                    destination.unlink(missing_ok=True)
                    raise
            generated.append(destination)
            log(f"装箱单已生成：{destination}（{len(rows)} 个 ITEM）")
            if on_generated is not None:
                try:
                    on_generated(destination)
                    log(f"已请求 Microsoft Excel 打开：{destination.name}")
                except Exception as exc:
                    log(f"装箱单已保存，但 Office 打开失败：{destination}；{exc}", "ERROR")
        except FileExistsError:
            skipped.append(order)
            log(f"订单 {order} 的装箱单已被创建，已跳过，继续下一订单。", "WARNING")
        except Exception as exc:
            failed.append(order)
            log(f"订单 {order} 未生成：{exc}", "ERROR")
    return generated, failed, skipped
