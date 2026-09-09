"""智能克隆引擎主入口"""
import os
from .extractor import extract_raw_style_data
from .translator import translate_to_yu_works_schema
from core.scene.schema import SceneConfig


def create_scene_from_clone(docx_path: str) -> SceneConfig:
    if not os.path.exists(docx_path):
        return SceneConfig(name="智能克隆模板")

    raw_style_data = extract_raw_style_data(docx_path)
    return translate_to_yu_works_schema(raw_style_data)
