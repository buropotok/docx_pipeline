from dagster import Definitions, define_asset_job, AssetSelection

from .assets import (
    input_docx_path,
    saveas_materialized,
    parse_raw_json,
    materialize_effective,
    optimize_tabs,          # новый asset
    reconstruct_docx,
)

full_run_job = define_asset_job(
    name="full_run_job",
    selection=AssetSelection.assets(
        "input_docx_path",
        "saveas_materialized",
        "parse_raw_json",
        "materialize_effective",
        "optimize_tabs",      # добавлен в job
        "reconstruct_docx",
    ),
)

defs = Definitions(
    assets=[
        input_docx_path,
        saveas_materialized,
        parse_raw_json,
        materialize_effective,
        optimize_tabs,        # добавлен в список
        reconstruct_docx,
    ],
    jobs=[full_run_job],
)