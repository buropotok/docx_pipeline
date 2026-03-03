from dagster import Definitions, define_asset_job, AssetSelection

from .assets import (
    input_docx_path,
    saveas_materialized,
    parse_raw_json,
    materialize_effective,
    optimize_tabs,          # новый asset
    reconstruct_docx_opt,
    reconstruct_docx,
    full_run_pipeline,
    full_run_pipeline_gpt
)

full_run_job = define_asset_job(
    name="full_run_job",
    selection=AssetSelection.assets(
        "input_docx_path",
        "saveas_materialized",
        "parse_raw_json",
        "reconstruct_docx",
    ),
)

full_run_job_opt = define_asset_job(
    name="full_run_job_opt",
    selection=AssetSelection.assets(
        "input_docx_path",
        "saveas_materialized",
        "parse_raw_json",
        "materialize_effective",
        "optimize_tabs",      # добавлен в job
        "reconstruct_docx_opt",
    ),
)

full_run_job_fast = define_asset_job(
    name="full_run_job_fast",
    selection=AssetSelection.assets(
        "full_run_pipeline",
    ),
)

full_run_job_fast_gpt = define_asset_job(
    name="full_run_job_fast_gpt",
    selection=AssetSelection.assets(
        "full_run_pipeline_gpt",
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
        reconstruct_docx_opt,
        full_run_pipeline,
        full_run_pipeline_gpt
    ],
    jobs=[full_run_job_fast, full_run_job, full_run_job_opt,full_run_job_fast_gpt],
)



