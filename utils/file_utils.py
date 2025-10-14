# utils/file_utils.py

def try_encodings_for_csv(csv_path):
    """
    Пробует открыть CSV в разных кодировках, возвращает (lines, enc).
    """
    encodings_to_try = [
        "utf-16-sig",
        "utf-16",
        "utf-8-sig",
        "utf-8",
        "cp1251",
        "latin1"
    ]
    for enc in encodings_to_try:
        try:
            with open(csv_path, 'r', encoding=enc) as f:
                lines = f.readlines()
            return lines, enc
        except Exception:
            pass

    raise Exception(
        f"Невозможно открыть файл {csv_path} ни в одной из кодировок: {encodings_to_try}"
    )
