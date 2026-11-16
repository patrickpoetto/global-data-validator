import polars as pl
import pandas as pd
from pathlib import Path

# =======================
# CONFIGURAZIONE UTENTE
# =======================

KEY_COLS = ['Sales Doc.', 'Item', 'Internal counter']


CSV_SEPARATOR = ";"   

file_path = Path(
    "path/to/your_data.csv"
    # Example alternative formats:
    # "path/to/your_data_with_key.csv"
    # "path/to/your_data.xlsx"
)

print("=== VERIFICA E CREAZIONE COLONNA KEY ===\n")

# =======================
# I/O FUNCTIONS
# =======================

def read_input_file(path: Path) -> pl.DataFrame:
    ext = path.suffix.lower()

    if ext in {".xlsx", ".xls"}:
        try:
            pdf = pd.read_excel(path, dtype=str) 
            df = pl.from_pandas(pdf)
        except Exception:
            df = pl.read_excel(path, sheet_id=0)

        for col in KEY_COLS:
            if col in df.columns:
                df = df.with_columns(pl.col(col).cast(pl.Utf8))

        return df

    if ext == ".csv":
        df = pl.read_csv(
            path,
            separator=CSV_SEPARATOR,
            has_header=True,
            infer_schema_length=0,
            encoding="utf8",
            quote_char = '"',
            ignore_errors=True,
            null_values=["", "NULL", "null", "None", "nan", "NaN"],
            try_parse_dates=False,
        )
        if df.width == 1:
            raise ValueError(
                f"CSV letto con una sola colonna. "
                f"Probabile separatore errato: {CSV_SEPARATOR}"
            )
        return df

    raise ValueError(f"Formato non supportato: {ext}")


def write_output_file(df: pl.DataFrame, source_path: Path) -> Path:
    ext = source_path.suffix.lower()

    if ext in {".xlsx", ".xls"}:
        out_path = source_path.parent / f"{source_path.stem}_with_key{ext}"
        try:
            df.write_excel(out_path)
        except Exception:
            df.to_pandas().to_excel(out_path, index=False)
        return out_path

    if ext == ".csv":
        out_path = source_path.parent / f"{source_path.stem}_with_key.csv"
        df.write_csv(out_path, separator=CSV_SEPARATOR)
        return out_path

    raise ValueError(f"Formato output non supportato: {ext}")

# =======================
# MAIN LOGIC
# =======================

try:
    print("Caricamento file...")
    df = read_input_file(file_path)
    print(f"File caricato: {file_path.name}")
    print(f"Dimensioni: {df.shape[0]:,} righe, {df.shape[1]} colonne")
    print(f"Colonne chiave: {KEY_COLS}")

    # -----------------------
    # CLEAN TARGET COLUMNS
    # -----------------------
    for target in ["key_col"]:
        if target in df.columns:
            df = df.with_columns(
                pl.col(target).cast(pl.Utf8)
            ).with_columns(
                pl.when(pl.col(target).is_null() | (pl.col(target) == ""))
                .then(pl.lit("na"))
                .otherwise(pl.col(target))
                .alias(target)
            )

    # -----------------------
    # CHECK REQUIRED COLUMNS
    # -----------------------
    missing_cols = [c for c in KEY_COLS if c not in df.columns]
    if missing_cols:
        print(f"❌ Colonne mancanti: {missing_cols}")
        print(f"Colonne disponibili: {df.columns}")
        raise SystemExit(1)

    print("✓ Tutte le colonne richieste presenti")

    # =======================
    # 1) NULL CHECK
    # =======================
    print("\n=== VERIFICA VALORI NULL ===")
    total_rows = df.height

    null_counts = {
        c: df.select(pl.col(c).is_null().sum()).item()
        for c in KEY_COLS
    }

    for c in KEY_COLS:
        print(
            f"{c}: tot={total_rows:,} | "
            f"null={null_counts[c]:,} | "
            f"non-null={total_rows-null_counts[c]:,}"
        )

    # Strict rule: key columns must not contain null values.
    condition1_met = all(null_counts[c] == 0 for c in KEY_COLS)

    # =======================
    # 2) UNICITÀ KEY
    # =======================
    print("\n=== VERIFICA UNICITÀ CONCATENAZIONE ===")

    all_not_null = pl.all_horizontal([pl.col(c).is_not_null() for c in KEY_COLS])
    concat_expr = pl.concat_str(
        [pl.col(c).cast(pl.Utf8) for c in KEY_COLS],
        separator="-"
    )

    df_with_key = df.with_columns(
        pl.when(all_not_null).then(concat_expr).otherwise(None).alias("Key_temp")
    )

    key_nulls = df_with_key.select(pl.col("Key_temp").is_null().sum()).item()
    key_non_nulls = total_rows - key_nulls
    unique_keys = df_with_key.select(
        pl.col("Key_temp")
        .filter(pl.col("Key_temp").is_not_null())
        .n_unique()
    ).item()

    print(f"Totali: {total_rows:,}")
    print(f"Key null: {key_nulls:,}")
    print(f"Key non-null: {key_non_nulls:,}")
    print(f"Key uniche: {unique_keys:,}")

    condition2_met = (unique_keys == key_non_nulls) and (key_non_nulls > 0)

    if not condition2_met and key_non_nulls > unique_keys:
        print("\n--- DETTAGLIO DUPLICATI ---")
        dup = (
            df_with_key
            .filter(pl.col("Key_temp").is_not_null())
            .group_by("Key_temp")
            .agg(pl.len().alias("count"))
            .filter(pl.col("count") > 1)
            .sort("count", descending=True)
        )
        #print(dup.head(20))
        with pl.Config(tbl_rows=-1, tbl_cols=-1, fmt_str_lengths=100):
            print(dup)

    # =======================
    # 3) DECISIONE FINALE
    # =======================
    print("\n=== RISULTATO FINALE ===")
    if condition1_met and condition2_met:
        print("🎯 CONDIZIONI SODDISFATTE → creo Key_col")

        df_final = df.with_columns(
            pl.when(all_not_null)
            .then(concat_expr)
            .otherwise(None)
            .alias("Key_col")
        )

        out_path = write_output_file(df_final, file_path)
        print(f"✅ File salvato: {out_path}")
        print(
            f"Dimensioni finali: "
            f"{df_final.shape[0]:,} righe, "
            f"{df_final.shape[1]} colonne"
        )
    else:
        print("❌ CONDIZIONI NON SODDISFATTE")
        if not condition1_met:
            print(" - Troppi null in una o più colonne chiave")
        if not condition2_met:
            print(" - Key non univoca")

except Exception as e:
    print(f"\n❌ ERRORE: {e}")
    import traceback
    traceback.print_exc()
