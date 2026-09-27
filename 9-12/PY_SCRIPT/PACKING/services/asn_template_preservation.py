"""Restore stored print scales that Excel normalizes while saving fit-to-page sheets."""
from pathlib import Path
import posixpath
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile


MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def sheet_parts(archive):
    links = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
    targets = {node.attrib['Id']: posixpath.normpath(posixpath.join('xl', node.attrib['Target']))
               if not node.attrib['Target'].startswith('/') else node.attrib['Target'].lstrip('/')
               for node in links}
    workbook = ET.fromstring(archive.read('xl/workbook.xml'))
    return {node.attrib['name']: targets[node.attrib[f'{{{REL}}}id']]
            for node in workbook.find(f'{{{MAIN}}}sheets')}


def preserve_template_metadata(template, output, *, preserve_vba=False):
    """Keep print scales and optional original VBA bytes without changing calculated cells."""
    output = Path(output)
    with ZipFile(template) as source, ZipFile(output) as generated:
        originals, outputs = sheet_parts(source), sheet_parts(generated)
        replacements = {}
        for name, original_path in originals.items():
            original_sheet = ET.fromstring(source.read(original_path))
            page = original_sheet.find(f'{{{MAIN}}}pageSetup')
            scale = page.get('scale') if page is not None else None
            path = outputs[name]
            content = generated.read(path)
            new_sheet = ET.fromstring(content)
            new_page = new_sheet.find(f'{{{MAIN}}}pageSetup')
            if preserve_vba:
                before = original_sheet.find(f'{{{MAIN}}}sheetPr')
                after = new_sheet.find(f'{{{MAIN}}}sheetPr')
                if (None if before is None else before.get('codeName')) != (None if after is None else after.get('codeName')):
                    raise ValueError(f'{name} 的宏工作表标识发生变化，未输出文件')
            actual = new_page.get('scale') if new_page is not None else None
            if actual == scale:
                continue
            # Editing the original bytes avoids reserializing namespace prefixes and extensions.
            def restore(match):
                tag = re.sub(rb'\s+scale="[^"]*"', b'', match.group())
                if scale is not None:
                    closing = b'/>' if tag.endswith(b'/>') else b'>'
                    tag = tag[:-len(closing)] + b' scale="' + scale.encode('ascii') + b'"' + closing
                return tag
            changed, count = re.subn(rb'<(?:[A-Za-z_][\w.-]*:)?pageSetup\b[^>]*>', restore, content)
            if count != 1:
                raise ValueError(f'无法保留模板 {name} 的打印缩放设置')
            replacements[path] = changed
        if preserve_vba:
            project = 'xl/vbaProject.bin'
            if project not in source.namelist() or project not in generated.namelist():
                raise ValueError('星辉仓宏项目缺失，未输出文件')
            original_workbook = ET.fromstring(source.read('xl/workbook.xml')).find(f'{{{MAIN}}}workbookPr')
            generated_workbook = ET.fromstring(generated.read('xl/workbook.xml')).find(f'{{{MAIN}}}workbookPr')
            if original_workbook.get('codeName') != generated_workbook.get('codeName'):
                raise ValueError('宏工作簿标识发生变化，未输出文件')
            # SaveAs can change VBA container timestamps even with macros disabled.
            # Restore the unchanged project after confirming all module host identifiers.
            replacements[project] = source.read(project)
        if not replacements:
            return
        temporary = output.with_suffix('.print-settings.tmp')
        try:
            with ZipFile(temporary, 'w') as restored:
                for item in generated.infolist():
                    restored.writestr(item, replacements.get(item.filename, generated.read(item.filename)))
                restored.comment = generated.comment
        except Exception:
            temporary.unlink(missing_ok=True)
            raise
    temporary.replace(output)
