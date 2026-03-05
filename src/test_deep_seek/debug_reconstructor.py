import subprocess
import json
import tempfile
import os

# Создаём временный JSON
test_json = tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False)
json_path = test_json.name

data = {
    "meta": {"schema_version": "2.15"},
    "document_info": {},
    "numbering_definitions": {},
    "doc_defaults": {"p_format": {}, "r_format": {}},
    "latent_styles": {},
    "styles": {},
    "content": [
        {
            "type": "paragraph",
            "id": "p_1.before1",
            "p_style_id": "Normal",
            "runs": [],
            "anchor": "p_1",
            "position": "before",
            "derive_from": "p_1"
        }
    ]
}

json.dump(data, test_json, indent=2)
test_json.close()

print(f"Created test JSON: {json_path}")

cmd = [
    "python", "src/reconstructor.py",
    "--in-json", json_path,
    "--out-docx", "src/test_deep_seek/test_output.docx",
    "--donor-docx", "src/test_deep_seek/donor.materialized.docx"
]

print(f"Running: {' '.join(cmd)}")
result = subprocess.run(cmd, capture_output=True, text=True)

print(f"Return code: {result.returncode}")
print(f"STDOUT: {result.stdout}")
print(f"STDERR: {result.stderr}")

# Удаляем временный файл
os.unlink(json_path)