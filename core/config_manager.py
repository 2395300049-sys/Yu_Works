"""Yu_Works 基础排版配置。

2.0 精简版只保存一级至三级标题的可视化设置。旧版的 AI、模板克隆、
自定义模式和供应商密钥不会再读取或迁移。
"""

from __future__ import annotations

import copy
import json
import os
from pathlib import Path


_CONFIG_PATH = Path.home() / ".yu_works_config.json"

CHINESE_FONT_SIZES = {
    "初号": 42.0,
    "小初": 36.0,
    "一号": 26.0,
    "小一": 24.0,
    "二号": 22.0,
    "小二": 18.0,
    "三号": 16.0,
    "小三": 15.0,
    "四号": 14.0,
    "小四": 12.0,
    "五号": 10.5,
    "小五": 9.0,
    "六号": 7.5,
    "小六": 6.5,
    "七号": 5.5,
    "八号": 5.0,
}


DEFAULT_HEADING_SETTINGS = {
    "1": {
        "font_cn": "黑体",
        "size_name": "小三",
        "size_pt": 15.0,
        "space_before_pt": 40.0,
        "space_after_pt": 20.0,
        "center": True,
    },
    "2": {
        "font_cn": "黑体",
        "size_name": "四号",
        "size_pt": 14.0,
        "space_before_pt": 24.0,
        "space_after_pt": 6.0,
        "center": False,
    },
    "3": {
        "font_cn": "黑体",
        "size_name": "小四",
        "size_pt": 12.0,
        "space_before_pt": 12.0,
        "space_after_pt": 6.0,
        "center": False,
    },
}


def load_full_config() -> dict:
    if not _CONFIG_PATH.exists():
        return {}
    try:
        with _CONFIG_PATH.open("r", encoding="utf-8-sig") as stream:
            data = json.load(stream)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError, TypeError):
        return {}


def save_full_config(config_dict: dict) -> None:
    """原子保存配置，避免程序异常退出时留下半个 JSON 文件。"""
    _CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    temp_path = _CONFIG_PATH.with_suffix(".tmp")
    with temp_path.open("w", encoding="utf-8") as stream:
        json.dump(config_dict, stream, ensure_ascii=False, indent=2)
    os.replace(temp_path, _CONFIG_PATH)


def update_config(key_or_dict, value=None) -> None:
    data = load_full_config()
    if isinstance(key_or_dict, dict):
        data.update(key_or_dict)
    else:
        data[key_or_dict] = value
    save_full_config(data)


def _number(value, *, default: float, minimum: float, maximum: float) -> float:
    try:
        result = float(value)
    except (TypeError, ValueError):
        return default
    return min(maximum, max(minimum, result))


def _font_size_name(source: dict, default: dict) -> str:
    """读取中文字号；旧版数值配置自动转换到最接近的标准字号。"""
    requested = str(source.get("size_name", "")).strip()
    if requested in CHINESE_FONT_SIZES:
        return requested
    old_size = _number(
        source.get("size_pt"),
        default=default["size_pt"],
        minimum=min(CHINESE_FONT_SIZES.values()),
        maximum=max(CHINESE_FONT_SIZES.values()),
    )
    return min(CHINESE_FONT_SIZES, key=lambda name: abs(CHINESE_FONT_SIZES[name] - old_size))


def normalize_heading_settings(settings: dict | None) -> dict:
    """只接受界面公开的标题字段，并限制异常数值。"""
    incoming = settings if isinstance(settings, dict) else {}
    normalized = copy.deepcopy(DEFAULT_HEADING_SETTINGS)
    for level in ("1", "2", "3"):
        source = incoming.get(level, {})
        if not isinstance(source, dict):
            continue
        default = DEFAULT_HEADING_SETTINGS[level]
        font_cn = str(source.get("font_cn", default["font_cn"])).strip()
        size_name = _font_size_name(source, default)
        normalized[level] = {
            "font_cn": font_cn or default["font_cn"],
            "size_name": size_name,
            "size_pt": CHINESE_FONT_SIZES[size_name],
            "space_before_pt": _number(
                source.get("space_before_pt"), default=default["space_before_pt"],
                minimum=0.0, maximum=200.0,
            ),
            "space_after_pt": _number(
                source.get("space_after_pt"), default=default["space_after_pt"],
                minimum=0.0, maximum=200.0,
            ),
            "center": bool(source.get("center", default["center"])),
        }
    return normalized


def get_heading_settings() -> dict:
    return normalize_heading_settings(load_full_config().get("heading_settings"))


def save_heading_settings(settings: dict) -> dict:
    normalized = normalize_heading_settings(settings)
    save_full_config({"heading_settings": normalized})
    return normalized


def get_active_scene_config():
    """加载唯一的基础排版预设，再覆盖界面中的三级标题设置。"""
    from core.scene.manager import load_scene_config

    preset_path = _resolve_preset_path("thesis_strict.json")
    scene_cfg = load_scene_config(preset_path)
    heading_settings = get_heading_settings()

    for level in ("1", "2", "3"):
        style = scene_cfg.styles[f"heading{level}"]
        values = heading_settings[level]
        style.font_cn = values["font_cn"]
        style.size_pt = values["size_pt"]
        style.space_before_pt = values["space_before_pt"]
        style.space_after_pt = values["space_after_pt"]
        style.alignment = "center" if values["center"] else "left"
        style.first_line_indent_cm = 0.0
        style.keep_with_next = True

    # 一级标题的段前分页是基础版固定规则，不允许旧配置关闭。
    scene_cfg.styles["heading1"].page_break_before = True
    return scene_cfg


def _resolve_preset_path(filename: str) -> Path:
    import sys

    if hasattr(sys, "_MEIPASS"):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).resolve().parent.parent
    return base / "presets" / filename
