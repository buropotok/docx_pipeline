#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import json
import os
import sys
from typing import Dict, List, Any

FONTS_DIR = os.path.dirname(__file__)
FONTS_PARAMS_PATH = os.path.join(FONTS_DIR, "fonts_params.json")

with open(FONTS_PARAMS_PATH, "r", encoding="utf-8") as f:
    FONTS_PARAMS = json.load(f)

DEFAULT_FONT = "Times New Roman"
DEFAULT_SIZE = 24  # half-points

def get_char_width(char: str, font_name: str, font_size_half: int) -> int:
    font_data = FONTS_PARAMS.get(font_name, {}).get("12pt", {})
    base = font_data.get(char, 120)  # средняя ширина
    return round(base * font_size_half / 24.0)

def text_width(text: str, font_name: str, font_size_half: int) -> int:
    return sum(get_char_width(ch, font_name, font_size_half) for ch in text)

def get_font_for_run(run: Dict, para_style_id: str, styles: Dict, char_styles: Dict) -> tuple:
    char_id = run.get("char_style_id")
    if char_id and char_id in char_styles:
        rfmt = char_styles[char_id].get("r_format", {})
        rfonts = rfmt.get("rFonts", {})
        font = rfonts.get("ascii") or rfonts.get("hAnsi") or DEFAULT_FONT
        size = rfmt.get("font_size_half_points", DEFAULT_SIZE)
        return font, size
    style = styles.get(para_style_id, {})
    rfmt = style.get("r_format", {})
    rfonts = rfmt.get("rFonts", {})
    font = rfonts.get("ascii") or rfonts.get("hAnsi") or DEFAULT_FONT
    size = rfmt.get("font_size_half_points", DEFAULT_SIZE)
    return font, size

def get_paragraph_indent(para: Dict, styles: Dict) -> int:
    """Возвращает отступ слева (indentStartTwip) из локального p_format или из стиля."""
    local_pf = para.get('p_format', {})
    if 'indentStartTwip' in local_pf:
        return local_pf['indentStartTwip']
    style_id = para.get('style_id')
    if style_id and style_id in styles:
        style_pf = styles[style_id].get('p_format', {})
        return style_pf.get('indentStartTwip', 0)
    return 0

def find_tab_groups(runs: List[Dict]) -> List[Dict]:
    groups = []
    i = 0
    while i < len(runs):
        if runs[i].get('type') != 'tab':
            i += 1
            continue
        start = i
        leading = runs[i].get('meta', {}).get('leading', False)
        while i < len(runs) and runs[i].get('type') == 'tab':
            i += 1
        groups.append({
            'start': start,
            'end': i,
            'count': i - start,
            'leading': leading
        })
    return groups

def optimize_paragraph(para: Dict, default_tab_stop: int, styles: Dict, char_styles: Dict) -> Dict:
    runs = para.get('runs', [])
    if not runs:
        return para

    # текущая позиция = отступ абзаца
    current_pos = get_paragraph_indent(para, styles)
    groups = find_tab_groups(runs)
    if not groups:
        return para

    new_runs = []
    i = 0
    group_idx = 0
    # копируем p_format, чтобы не менять оригинал раньше времени
    p_format = para.get('p_format', {}).copy()
    tabs_added = False

    while i < len(runs):
        if group_idx < len(groups) and i == groups[group_idx]['start']:
            grp = groups[group_idx]
            pos_before = current_pos
            target_pos = pos_before + grp['count'] * default_tab_stop

            # добавляем позицию табуляции (если ещё нет близкой)
            if 'tabs' not in p_format:
                p_format['tabs'] = []
            exists = any(abs(t['posTwip'] - target_pos) < 5 for t in p_format['tabs'])
            if not exists:
                p_format['tabs'].append({"posTwip": target_pos, "val": "left"})
                tabs_added = True

            # заменяем группу одним табом
            new_tab = {"type": "tab"}
            if grp['leading']:
                new_tab['meta'] = {"leading": True}
            new_runs.append(new_tab)

            current_pos = target_pos
            i = grp['end']
            group_idx += 1
        else:
            run = runs[i]
            if run['type'] == 'text':
                font, size = get_font_for_run(run, para.get('style_id'), styles, char_styles)
                current_pos += text_width(run.get('text', ''), font, size)
            elif run['type'] == 'tab':
                # одиночный таб вне группы (не должен встречаться)
                current_pos += default_tab_stop
            new_runs.append(run)
            i += 1

    para['runs'] = new_runs
    if tabs_added:
        para['p_format'] = p_format
    return para

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--in-json", required=True)
    parser.add_argument("--out-json", required=True)
    args = parser.parse_args()

    with open(args.in_json, "r", encoding="utf-8") as f:
        data = json.load(f)

    default_tab_stop = data['document_info']['settings'].get('defaultTabStopTwip', 708)
    styles = data.get('styles', {})
    char_styles = data.get('character_styles', {})

    for para in data['content']:
        optimize_paragraph(para, default_tab_stop, styles, char_styles)

    with open(args.out_json, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print("[optimize_tabs] done")

if __name__ == "__main__":
    main()