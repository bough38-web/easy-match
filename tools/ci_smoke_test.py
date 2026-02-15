import os
from matcher import match_universal

def main():
    base = {
        "type": "file",
        "path": os.path.join("samples", "base.csv"),
        "sheet": "CSV",
        "header": 1,
        "book": ""
    }
    target = {
        "type": "file",
        "path": os.path.join("samples", "target.csv"),
        "sheet": "CSV",
        "header": 1,
        "book": ""
    }
    out_dir = "outputs"
    out_path, summary = match_universal(
        base, target,
        key_cols=["사번"],
        take_cols=["부서","입사일"],
        out_dir=out_dir,
        options={"fuzzy": False, "color": False},
        replacement_rules=None,
        progress=None
    )
    assert os.path.exists(out_path), f"Output missing: {out_path}"
    print(summary)
    print("CI smoke OK")

if __name__ == "__main__":
    main()
