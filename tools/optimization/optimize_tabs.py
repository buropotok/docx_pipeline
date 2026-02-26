#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import json
import os
import sys
from typing import Dict, List, Any, Optional

# Загрузка метрик шрифтов
FONTS_DIR = os.path.dirname(__file__)
FONTS_PARAMS_PATH = os.path.join(FONTS_DIR, "fonts_params.json")

with open(FONTS_PARAMS_PATH, "r", encoding="utf-8") as f:
    FONTS_PARAMS = json.load(f)

DEFAULT_FONT = "Times New Roman"
DEFAULT_SIZE = 24  # half-points

def get_char_width(char: str, font_name: str, font_size_half: int) -> int:
    """Возвращает ширину символа в твипах для заданного шрифта и размера."""
    font_data = FONTS_PARAMS.get(font_name, {}).get("12pt", {})
    base_width = font_data.get(char, 213)  # средняя ширина по умолчанию
    scale = font_size_half / 24.0
    return round(base_width * scale)

def text_width(text: str, font_name: str, font_size_half: int) -> int:
    """Суммарная ширина текста в твипах."""
    return sum(get_char_width(ch, font_name, font_size_half) for ch in text)

def get_font_for_run(run: Dict[str, Any], para_style_id: str,
                     styles: Dict, char_styles: Dict) -> tuple:
    """
    Определяет имя шрифта и размер (в half-points) для данного run.
    Возвращает (font_name, font_size_half).
    """
    # Приоритет: символьный стиль -> стиль абзаца -> умолчание
    char_id = run.get("char_style_id")
    if char_id and char_id in char_styles:
        rfmt = char_styles[char_id].get("r_format", {})
        rfonts = rfmt.get("rFonts", {})
        font_name = rfonts.get("ascii") or rfonts.get("hAnsi") or DEFAULT_FONT
        size = rfmt.get("font_size_half_points", DEFAULT_SIZE)
        return font_name, size

    style = styles.get(para_style_id, {})
    rfmt = style.get("r_format", {})
    rfonts = rfmt.get("rFonts", {})
    font_name = rfonts.get("ascii") or rfonts.get("hAnsi") or DEFAULT_FONT
    size = rfmt.get("font_size_half_points", DEFAULT_SIZE)
    return font_name, size

def find_tab_groups(runs: List[Dict]) -> List[Dict]:
    """Находит непрерывные последовательности табов в списке runs."""
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

def optimize_paragraph(para: Dict, default_tab_stop: int,
                       styles: Dict, char_styles: Dict) -> Dict:
    """Оптимизирует один абзац: заменяет группы табов и добавляет позиции."""
    runs = para.get('runs', [])
    if not runs:
        return para

    # Начальная позиция = отступ слева (indentStartTwip)
    p_format = para.get('p_format', {}).copy()
    current_pos = p_format.get('indentStartTwip', 0)

    groups = find_tab_groups(runs)
    if not groups:
        return para

    new_runs = []
    i = 0
    group_idx = 0
    debug_info = []

    while i < len(runs):
        if group_idx < len(groups) and i == groups[group_idx]['start']:
            grp = groups[group_idx]

            # Сохраняем позицию ДО группы (In + L)
            pos_before_group = current_pos

            # Вычисляем желаемую позицию после группы по формуле In + L + N*708
            target_pos = pos_before_group + grp['count'] * default_tab_stop

            debug_info.append(
                f"Group at {grp['start']}: count={grp['count']}, "
                f"pos_before={pos_before_group}, target_pos={target_pos}"
            )

            # Добавляем позицию табуляции в p_format (если ещё нет такой)
            if 'tabs' not in p_format:
                p_format['tabs'] = []
            # Проверим, нет ли уже такой позиции (с допуском 5 твипов)
            exists = any(abs(t['posTwip'] - target_pos) < 5 for t in p_format['tabs'])
            if not exists:
                p_format['tabs'].append({"posTwip": target_pos, "val": "left"})
                debug_info.append(f"  Added tab stop at {target_pos}")
            else:
                debug_info.append(f"  Tab stop at {target_pos} already exists (skipped)")

            # Создаём один таб для замены группы
            new_tab = {"type": "tab"}
            if grp['leading']:
                new_tab['meta'] = {"leading": True}
            new_runs.append(new_tab)

            # Обновляем текущую позицию до значения после группы
            current_pos = target_pos

            i = grp['end']
            group_idx += 1
        else:
            run = runs[i]
            if run['type'] == 'text':
                font_name, font_size = get_font_for_run(run, para.get('style_id'), styles, char_styles)
                text = run.get('text', '')
                width = text_width(text, font_name, font_size)
                debug_info.append(f"Text at {i}: '{text[:20]}...' width={width}")
                current_pos += width
            elif run['type'] == 'tab':
                # Одиночный таб вне группы (не должен встречаться, если группы выделены верно)
                current_pos += default_tab_stop
                debug_info.append(f"Lonely tab at {i}, pos now {current_pos}")
            new_runs.append(run)
            i += 1

    # Для отладки выведем информацию
    if debug_info:
        style_id = para.get('style_id', 'unknown')
        print(f"[debug] Paragraph {style_id}:")
        for line in debug_info:
            print(f"  {line}")

    para['runs'] = new_runs
    if p_format.get('tabs'):
        para['p_format'] = p_format
    return para

def main():
    parser = argparse.ArgumentParser(description="Optimize tabs using font metrics")
    parser.add_argument("--in-json", required=True, help="Input JSON file")
    parser.add_argument("--out-json", required=True, help="Output JSON file")
    args = parser.parse_args()

    in_json = os.path.abspath(args.in_json)
    out_json = os.path.abspath(args.out_json)

    print(f"[optimize_tabs] in-json: {in_json}")
    print(f"[optimize_tabs] out-json: {out_json}")

    with open(in_json, "r", encoding="utf-8") as f:
        data = json.load(f)

    default_tab_stop = data['document_info']['settings'].get('defaultTabStopTwip', 708)
    styles = data.get('styles', {})
    char_styles = data.get('character_styles', {})

    for para in data['content']:
        optimize_paragraph(para, default_tab_stop, styles, char_styles)

    os.makedirs(os.path.dirname(out_json) or ".", exist_ok=True)
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print("[optimize_tabs] done")
    return 0

if __name__ == "__main__":
    sys.exit(main())