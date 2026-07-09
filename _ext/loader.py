# -*- coding: utf-8 -*-
"""
process_data 外部脚本加载器

外部脚本放在 _ext/process_data.py，直接 copy 覆盖不用改。
本 loader 负责：
  1. 执行外部脚本
  2. 把输出从 src/data/dashboard.json 移到实际位置
"""
import importlib.util
import os
import shutil

_HERE = os.path.dirname(os.path.abspath(__file__))
_PROJECT = os.path.dirname(_HERE)
_SCRIPT = os.path.join(_HERE, "process_data.py")
_EXT_OUTPUT_FILE = os.path.join(_PROJECT, "src", "data", "dashboard.json")


def _resolve_output_dir():
    cfg_path = os.path.join(_PROJECT, "config.yaml")
    if os.path.exists(cfg_path):
        try:
            import yaml
            with open(cfg_path, "r", encoding="utf-8") as f:
                cfg = yaml.safe_load(f) or {}
            if "dataview" in cfg:
                return cfg["dataview"]
        except Exception:
            pass
    return os.path.join(_PROJECT, "result")


def run():
    output_dir = _resolve_output_dir()
    os.makedirs(output_dir, exist_ok=True)

    saved_cwd = os.getcwd()
    os.chdir(_PROJECT)
    try:
        spec = importlib.util.spec_from_file_location(
            "process_data._ext", _SCRIPT,
            submodule_search_locations=[],
        )
        mod = importlib.util.module_from_spec(spec)
        mod.__file__ = _SCRIPT
        spec.loader.exec_module(mod)

        if os.path.exists(_EXT_OUTPUT_FILE):
            dest = os.path.join(output_dir, "dashboard.json")
            shutil.move(_EXT_OUTPUT_FILE, dest)
            print(f"[loader] 输出已生成: {dest}")
            return dest
        return None

    finally:
        os.chdir(saved_cwd)
