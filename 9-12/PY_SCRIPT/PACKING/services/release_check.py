"""Explicit --self-test checks for the packaged executable; never run during normal use."""
import io
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import traceback


def run_release_check(report_path):
    result = {'passed': False, 'checks': []}
    app = None
    try:
        from config import PROJECT_DIR, RESOURCE_DIR, SETTINGS_FILE, PACKING_RULES_FILE, ASN_TEMPLATE, PACKING_TEMPLATE, FACTORY_TEMPLATE
        from services.settings_service import read_settings
        from ui.app import PackingApp
        app = PackingApp()
        app.withdraw()
        app.update_idletasks()
        assert app.title().casefold() == 'Love Anna, 加油!'.casefold()
        assert app.asn_tab.select_button.cget('text') == '选择三文件所在的目录'
        assert app.tabs.get() == '入仓预申报ASN生成器'
        assert SETTINGS_FILE.parent == PROJECT_DIR and SETTINGS_FILE.is_file()
        assert PACKING_RULES_FILE.is_file()
        result.update(app_dir=str(PROJECT_DIR), resource_dir=str(RESOURCE_DIR), window_title=app.title(),
                      font_size=read_settings(SETTINGS_FILE)['ui_font_size'])
        result['checks'].append('GUI, external JSON creation/loading, button text')

        import xlrd
        from openpyxl import load_workbook
        for template in (PACKING_TEMPLATE, FACTORY_TEMPLATE):
            book = load_workbook(template, read_only=True)
            assert book.sheetnames
            book.close()
        book = xlrd.open_workbook(str(ASN_TEMPLATE))
        assert '报关资料与ASN' in book.sheet_names()
        book.release_resources()
        result['checks'].append('All three embedded Excel templates and XLS/XLSX readers')

        from pypdf import PdfReader, PdfWriter
        pdf = io.BytesIO()
        writer = PdfWriter()
        writer.add_blank_page(width=100, height=100)
        writer.write(pdf)
        pdf.seek(0)
        assert len(PdfReader(pdf).pages) == 1
        result['checks'].append('PDF reader/writer')

        from services.asn_generation_service import write_asn_workbook
        row = [None] * 47
        row[0], row[1], row[5], row[26], row[28] = 1, 'RELEASE-CHECK', '00123', 8, 7.5
        plan = {'template': str(ASN_TEMPLATE), 'sheet': '报关资料与ASN', 'scope': 'asn_only',
                'contract': 'RELEASE-CHECK', 'headers': {'C8': 'Release test customer'}, 'rows': [row],
                'formulas': [{'cell': 'AD26', 'value': '=ROUND(AA26*AC26,2)'}]}
        with TemporaryDirectory(prefix='love-anna-check-') as temp:
            output = write_asn_workbook(plan, temp)
            book = xlrd.open_workbook(str(output))
            assert book.sheet_names() == ['报关资料与ASN']
            sheet = book.sheet_by_index(0)
            assert sheet.cell_value(7, 2) == 'Release test customer'
            assert sheet.cell_value(25, 5) == '00123'
            assert sheet.cell_value(25, 29) == 60
            book.release_resources()
        result['checks'].append('Embedded PowerShell writer and native Excel XLS generation')
        result['passed'] = True
    except Exception:
        result['error'] = traceback.format_exc()
    finally:
        if app is not None:
            for timer in app.tk.call('after', 'info'):
                app.tk.call('after', 'cancel', timer)
            app.destroy()
        Path(report_path).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    return 0 if result['passed'] else 1
