from pathlib import Path
from PIL import Image, ImageOps
from models.table import TableData
from services.table_layout import TextBox, reconstruct_table


class TableRecognitionService:
    """本地 CPU OCR；模型由 rapidocr==3.4.0 的 wheel 随依赖一起安装。"""

    def __init__(self):
        self._engine = None

    def _get_engine(self):
        if self._engine is None:
            try:
                from rapidocr import RapidOCR
                import onnxruntime  # noqa: F401 -- report missing runtime before inference
            except ImportError as exc:
                raise RuntimeError("缺少 OCR 依赖，请在运行本程序的 Python 环境执行：python -m pip install -r requierement.txt") from exc
            self._engine = RapidOCR(params={
                "Global.log_level": "warning",
                "Global.max_side_len": 4096,
                "EngineConfig.onnxruntime.intra_op_num_threads": 4,
                "EngineConfig.onnxruntime.inter_op_num_threads": 1,
            })
        return self._engine

    def recognize(self, image_path: Path) -> TableData:
        engine = self._get_engine()
        import numpy as np

        with Image.open(image_path) as source:
            image = ImageOps.exif_transpose(source).convert("RGB")
            scale = min(3.0, 3900 / max(image.size))
            image = image.resize((max(1, round(image.width * scale)), max(1, round(image.height * scale))), Image.Resampling.BICUBIC)
            image = ImageOps.expand(image, border=30, fill="white")
            # RapidOCR's ndarray input expects OpenCV BGR order.
            result = engine(np.asarray(image)[:, :, ::-1].copy())
        if result.boxes is None or not result.txts:
            raise ValueError("未识别到文字，请选择清晰的表格图片。")
        boxes = [TextBox(float(box[:, 0].min()), float(box[:, 1].min()),
                         float(box[:, 0].max()), float(box[:, 1].max()), text)
                 for box, text in zip(result.boxes, result.txts)]
        return reconstruct_table(boxes)
