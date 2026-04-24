from __future__ import annotations

from pathlib import Path
import threading
from typing import Any


class OcrDependencyError(RuntimeError):
    pass


_rapid_ocr_engine_lock = threading.Lock()
_rapid_ocr_engine: Any | None = None
_rapid_ocr_engine_init_error: Exception | None = None


def _get_rapid_ocr_engine() -> Any:
    global _rapid_ocr_engine, _rapid_ocr_engine_init_error
    with _rapid_ocr_engine_lock:
        if _rapid_ocr_engine is not None:
            return _rapid_ocr_engine

        if _rapid_ocr_engine_init_error is not None:
            raise OcrDependencyError("OCR 依赖不可用，请先安装识别依赖。") from _rapid_ocr_engine_init_error

        try:
            from rapidocr_onnxruntime import RapidOCR
        except Exception as exc:
            _rapid_ocr_engine_init_error = exc
            raise OcrDependencyError("未检测到 rapidocr-onnxruntime / onnxruntime。") from exc

        try:
            _rapid_ocr_engine = RapidOCR()
        except Exception as exc:
            _rapid_ocr_engine_init_error = exc
            raise RuntimeError("OCR 引擎初始化失败。") from exc

        return _rapid_ocr_engine


def _extract_ocr_line_text(item: Any) -> str:
    if item is None:
        return ""

    if isinstance(item, str):
        return item.strip()

    if isinstance(item, dict):
        for key in ("text", "txt", "rec_txt", "ocr_text", "label"):
            value = item.get(key)
            if isinstance(value, str) and value.strip():
                return value.strip()
        for value in item.values():
            nested_text = _extract_ocr_line_text(value)
            if nested_text:
                return nested_text
        return ""

    if isinstance(item, (list, tuple, set)):
        sequence = list(item)
        if len(sequence) >= 2 and isinstance(sequence[1], str) and sequence[1].strip():
            return sequence[1].strip()
        if sequence and isinstance(sequence[0], str) and sequence[0].strip():
            return sequence[0].strip()
        for value in sequence:
            nested_text = _extract_ocr_line_text(value)
            if nested_text:
                return nested_text
        return ""

    return ""


def _normalize_ocr_output(result: Any) -> str:
    payload = result[0] if isinstance(result, tuple) and result else result
    if isinstance(payload, dict):
        payload = payload.get("rec_res", payload.get("result", payload))

    if isinstance(payload, str):
        return payload.strip()

    if isinstance(payload, (list, tuple, set)):
        lines: list[str] = []
        for item in payload:
            line_text = _extract_ocr_line_text(item)
            if line_text:
                lines.append(line_text)
        return "\n".join(lines).strip()

    return _extract_ocr_line_text(payload)


def recognize_image_text(image_path: Path) -> str:
    if not image_path.exists():
        raise FileNotFoundError("图片文件不存在。")

    ocr_engine = _get_rapid_ocr_engine()
    try:
        result = ocr_engine(str(image_path))
    except Exception as exc:
        raise RuntimeError("图片文字识别失败。") from exc

    return _normalize_ocr_output(result)
