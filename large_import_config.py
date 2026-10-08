"""Configuration helpers for large academic imports."""
import os

IMPORT_BATCH_SIZE = max(100, int(os.getenv("IMPORT_BATCH_SIZE", "2000")))
MAX_IMPORT_ROWS = max(0, int(os.getenv("MAX_IMPORT_ROWS", "500000")))

def validate_import_row_count(row_count: int) -> None:
    if row_count < 0:
        raise ValueError("Invalid row count.")
    if MAX_IMPORT_ROWS and row_count > MAX_IMPORT_ROWS:
        raise ValueError(
            f"This import contains {row_count:,} rows; the configured maximum is "
            f"{MAX_IMPORT_ROWS:,}. Increase MAX_IMPORT_ROWS or split the import."
        )
