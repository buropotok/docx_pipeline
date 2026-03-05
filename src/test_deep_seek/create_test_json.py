# create_test_json.py
import json

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

with open("src/test_deep_seek/test_minimal.json", "w") as f:
    json.dump(data, f, indent=2)