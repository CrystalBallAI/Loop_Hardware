"""
Subsystem registry — desktop edition.

Same registry as the dev server, with one structural change: a subsystem's
"codebase root" is no longer the source *_CodeBase folder but the per-user
writable workspace (app-data), materialized at startup from the bundled
runtime_data/ (paths.json, spec, recommendations library, caches). The
pipeline CODE ships compiled inside the binary (cbmi_pipelines.*); only the
DATA the pipelines read relative to their root lives in the workspace.

Layout contract (kept from the dev server): paths.run.json is written at the
workspace root with LEXICAL ../input_data/<key>/ input paths, so the
pipelines' `root / value` and `.relative_to(root)` both keep working —
input_data/ sits beside the per-subsystem workspace roots.
"""
from __future__ import annotations

from pathlib import Path

from cbmi import appdata
from cbmi.resources import webui_dir

APP_DIR = webui_dir()

STEP_ORDER = ["validate", "pipeline", "recommendations", "excel", "provenance"]
STEP_LABELS = {
    "validate": "Validate inputs",
    "pipeline": "Process & Score",
    "recommendations": "Recommendations",
    "excel": "Excel report",
    "provenance": "Provenance",
}


class FileSlot:
    """Describes how one Load-page input id is staged into the workspace layout.

    kind:
      'dir'  -> all uploaded files copied into <input_root>/<dest> (names kept)
      'file' -> single uploaded file written as <input_root>/<dest> (renamed)
    """

    def __init__(self, input_id, kind, dest, required=False, also_copy_to=None):
        self.input_id = input_id
        self.kind = kind
        self.dest = dest
        self.required = required
        self.also_copy_to = also_copy_to  # workspace-root-relative extra copy


class PointSlot:
    """Per-point file slot (GCP / CheckPoint), staged into a point folder."""

    def __init__(self, slot_id, kind, dest, required=False):
        self.slot_id = slot_id
        self.kind = kind
        self.dest = dest
        self.required = required


class Subsystem:
    def __init__(self, key, label, module, source_codebase, score_file, score_key,
                 xlsx_name, per_point=False, slots=None, point_slots=None,
                 point_folder_prefix=None, inputs_override=None):
        self.key = key
        self.label = label
        self.module = module
        self.source_codebase = source_codebase  # dev-only: sample staging source
        self.score_file = score_file
        self.score_key = score_key
        self.xlsx_name = xlsx_name
        self.per_point = per_point
        self.slots = slots or []
        self.point_slots = point_slots or []
        self.point_folder_prefix = point_folder_prefix
        self.inputs_override = inputs_override or {}

    # ---- paths (all in the writable per-user workspace) --------------------
    @property
    def codebase(self) -> Path:
        """The pipeline's runtime root (spec, libs, cache, outputs, paths.json)."""
        return appdata.pipeline_root(self.key)

    @property
    def input_root(self) -> Path:
        return appdata.input_data_root() / self.key

    @property
    def run_paths_file(self) -> Path:
        return self.codebase / "paths.run.json"

    @property
    def outputs_dir(self) -> Path:
        return self.codebase / "outputs"

    def output_path(self, rel: str) -> Path:
        return self.codebase / rel


SUBSYSTEMS = {
    "drone": Subsystem(
        key="drone", label="Drone", module="drone",
        source_codebase="Drone_CodeBase",
        score_file="outputs/06_drone_score.json", score_key="drone_score",
        xlsx_name="outputs/drone_results.xlsx",
        slots=[
            FileSlot("drone_images",      "dir",  "images",                 required=True),
            FileSlot("drone_rover_rinex", "dir",  "rover_rinex",            required=True),
            FileSlot("drone_bin",         "dir",  "telemetry",              required=True),
            FileSlot("drone_form",        "file", "user_input/form.json",   required=True),
            FileSlot("drone_hardware",    "file", "user_input/hardware.json"),
        ],
        inputs_override={
            "images_folder":      "images",
            "rinex_folder":       "rover_rinex",
            "bin_folder":         "telemetry",
            "user_input_file":    "user_input/form.json",
            "user_hardware_file": "user_input/hardware.json",
        },
    ),
    "basestation": Subsystem(
        key="basestation", label="Base Station", module="base",
        source_codebase="BaseStation_CodeBase",
        score_file="outputs/06_base_station_score.json", score_key="base_station_score",
        xlsx_name="outputs/base_station_results.xlsx",
        slots=[
            FileSlot("base_rinex",      "dir",  "base_rinex",                      required=True),
            FileSlot("base_oplog",      "file", "operator_log/operation_log.json", required=True),
            FileSlot("base_user_input", "file", "user_input/user_input.json",      required=True),
            # stage1_inventory consults <root>/sample_data/hardware.json for the
            # hardware override, so mirror the upload there too.
            FileSlot("base_hardware",   "file", "hardware.json",
                     also_copy_to="sample_data/hardware.json"),
        ],
        inputs_override={
            "rinex_folder":        "base_rinex",
            "operator_log_folder": "operator_log",
            "user_input_folder":   "user_input",
        },
    ),
    "gcp": Subsystem(
        key="gcp", label="Control Point", module="gcp",
        source_codebase="GCP_CodeBase",
        score_file="outputs/06_gcp_score.json", score_key="gcp_score",
        xlsx_name="outputs/gcp_results.xlsx",
        per_point=True, point_folder_prefix="gcp_rinex_point_",
        point_slots=[
            PointSlot("rinex",      "dir",  "",                required=True),
            PointSlot("user_input", "file", "user_input.json", required=True),
            PointSlot("hardware",   "file", "hardware.json"),
            PointSlot("oplog",      "file", "oplog.json"),
        ],
        inputs_override={"points_root": "."},
    ),
    "checkpoint": Subsystem(
        key="checkpoint", label="Check Point", module="checkpoint",
        source_codebase="CheckPoint_CodeBase",
        score_file="outputs/06_check_point_score.json", score_key="check_point_score",
        xlsx_name="outputs/check_point_results.xlsx",
        per_point=True, point_folder_prefix="checkpoint_rtk_point_",
        point_slots=[
            PointSlot("rtk_export", "file", "cp_rtk_export.csv",  required=True),
            PointSlot("user_input", "file", "cp_user_input.json", required=True),
            PointSlot("oplog",      "file", "cp_oplog.json"),
        ],
        inputs_override={"points_root": "."},
    ),
}

MODULE_TO_KEY = {s.module: s.key for s in SUBSYSTEMS.values()}


def get(key_or_module: str) -> Subsystem:
    if key_or_module in SUBSYSTEMS:
        return SUBSYSTEMS[key_or_module]
    if key_or_module in MODULE_TO_KEY:
        return SUBSYSTEMS[MODULE_TO_KEY[key_or_module]]
    raise KeyError(f"unknown subsystem: {key_or_module}")
