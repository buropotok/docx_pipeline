import json
import configparser
from datetime import datetime
from pathlib import Path

from openai import OpenAI


def load_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")



def load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def save_text(path: Path, content: str) -> None:
    path.write_text(content, encoding="utf-8")


def save_json(path: Path, data) -> None:
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def extract_json_from_text(text: str):
    """
    Пытается распарсить JSON:
    1) как есть
    2) из ```json ... ```
    3) из блока между первой { / [ и последней } / ]
    """
    text = text.strip()

    # 1. Прямой JSON
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # 2. JSON в markdown fenced block
    if "```" in text:
        parts = text.split("```")
        for part in parts:
            candidate = part.strip()
            if candidate.startswith("json"):
                candidate = candidate[4:].strip()
            try:
                return json.loads(candidate)
            except json.JSONDecodeError:
                continue

    # 3. Поиск JSON-объекта/массива внутри текста
    start_obj = text.find("{")
    end_obj = text.rfind("}")
    if start_obj != -1 and end_obj != -1 and end_obj > start_obj:
        candidate = text[start_obj:end_obj + 1]
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            pass

    start_arr = text.find("[")
    end_arr = text.rfind("]")
    if start_arr != -1 and end_arr != -1 and end_arr > start_arr:
        candidate = text[start_arr:end_arr + 1]
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            pass

    raise ValueError("Не удалось извлечь валидный JSON из ответа модели.")


def main():
    # fingerprint_test.py лежит в: docx_pipeline/ai/test/fingerprint_test.py
    current_file = Path(__file__).resolve()
    test_dir = current_file.parent                    # docx_pipeline/ai/test
    ai_dir = test_dir.parent                         # docx_pipeline/ai
    project_root = ai_dir.parent                     # docx_pipeline

    config_path = test_dir / "config.ini"
    donor_path = test_dir / "donor.json"
    prompt_path = ai_dir / "fingerprint_prompt_LLM_safe.txt"
    schema_path = project_root / "schema" / "fingerprint_schema_v2_2.json"

    # Проверка файлов
    required_files = [config_path, donor_path, prompt_path, schema_path]
    missing_files = [str(p) for p in required_files if not p.exists()]
    if missing_files:
        raise FileNotFoundError(
            "Не найдены обязательные файлы:\n" + "\n".join(missing_files)
        )

    # Конфиг
    config = configparser.ConfigParser()
    config.read(config_path, encoding="utf-8")

    or_api_key = config.get("API", "OR_API_KEY")
    model_name = config.get("SETTINGS", "MODEL_NAME_gpt54pro")

    client = OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=or_api_key
    )

    # Данные
    donor_data = load_json(donor_path)
    prompt_text = load_text(prompt_path)
    schema_data = load_json(schema_path)

    # Сообщения
    system_message = (
        "Ты должен вернуть строго валидный JSON без пояснений, без markdown, "
        "без обертки в ```json. "
        "Ответ должен соответствовать переданной JSON Schema."
    )

    user_message = (
        f"{prompt_text.strip()}\n\n"
        f"=== JSON SCHEMA ===\n"
        f"{json.dumps(schema_data, ensure_ascii=False, indent=2)}\n\n"
        f"=== INPUT JSON ===\n"
        f"{json.dumps(donor_data, ensure_ascii=False, indent=2)}\n"
    )

    # Запрос
    response = client.chat.completions.create(
        model=model_name,
        temperature=0,
        messages=[
            {"role": "system", "content": system_message},
            {"role": "user", "content": user_message},
        ],
    )

    content = response.choices[0].message.content or ""

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    raw_response_path = test_dir / f"fingerprint_response_raw_{timestamp}.txt"
    parsed_response_path = test_dir / f"fingerprint_response_{timestamp}.json"
    latest_raw_path = test_dir / "fingerprint_response_raw_latest.txt"
    latest_json_path = test_dir / "fingerprint_response_latest.json"

    # Сохраняем raw
    save_text(raw_response_path, content)
    save_text(latest_raw_path, content)

    # Пытаемся распарсить JSON
    parsed_json = extract_json_from_text(content)

    # Сохраняем JSON
    save_json(parsed_response_path, parsed_json)
    save_json(latest_json_path, parsed_json)

    print("Готово.")
    print(f"Raw response:   {raw_response_path}")
    print(f"Parsed JSON:    {parsed_response_path}")
    print(f"Latest raw:     {latest_raw_path}")
    print(f"Latest parsed:  {latest_json_path}")


if __name__ == "__main__":
    main()