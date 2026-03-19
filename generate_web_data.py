from __future__ import annotations

import json
import re
from datetime import datetime, timezone

import openpyxl

SOURCE_FILE = "zz_structured.xlsx"
OUTPUT_FILE = "data.js"


def to_int(value: object) -> int | None:
    if value is None:
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    text = str(value).strip()
    if not text:
        return None
    return int(float(text))


def normalize_cell_text(value: object) -> str:
    if value is None:
        return ""
    raw = str(value).replace("\r\n", "\n").replace("\r", "\n").strip()
    if not raw:
        return ""
    # Treat slash as paragraph delimiter from spreadsheet cells.
    parts = [part.strip() for part in re.split(r"[\\/／]+", raw) if part.strip()]
    return "\n".join(parts).strip()


def main() -> None:
    wb = openpyxl.load_workbook(SOURCE_FILE, data_only=True)
    ws_sent = wb["small_sentences"]
    ws_anno = wb["annotations"]
    ws_interp = wb["interpretations"]

    texts: dict[int, dict] = {}
    sentence_note_keys: set[tuple[int, int]] = set()

    annotations_by_key: dict[str, list[dict]] = {}
    interpretations: list[dict] = []

    for row in ws_anno.iter_rows(min_row=2, values_only=True):
        annotation_id, text_id, sentence_id, commentator, dynasty, annotation = row[:6]
        tid = to_int(text_id)
        sid = to_int(sentence_id)
        annotation_text = normalize_cell_text(annotation)
        if tid is None or sid is None or not annotation_text:
            continue

        key = f"{tid}-{sid}"
        sentence_note_keys.add((tid, sid))
        annotations_by_key.setdefault(key, []).append(
            {
                "annotation_id": to_int(annotation_id),
                "commentator": str(commentator or "未署名"),
                "dynasty": str(dynasty or "未详"),
                "content": annotation_text,
            }
        )

    for row in ws_interp.iter_rows(min_row=2, values_only=True):
        interpretation_id, text_id, start_sid, end_sid, commentator, dynasty, content = row[:7]
        tid = to_int(text_id)
        start_id = to_int(start_sid)
        end_id = to_int(end_sid)
        content_text = normalize_cell_text(content)
        if tid is None or start_id is None or end_id is None or not content_text:
            continue

        if end_id < start_id:
            start_id, end_id = end_id, start_id

        for sid in range(start_id, end_id + 1):
            sentence_note_keys.add((tid, sid))

        interpretations.append(
            {
                "interpretation_id": to_int(interpretation_id),
                "text_id": tid,
                "start_sentence_id": start_id,
                "end_sentence_id": end_id,
                "commentator": str(commentator or "未署名"),
                "dynasty": str(dynasty or "未详"),
                "content": content_text,
            }
        )

    for row in ws_sent.iter_rows(min_row=2, values_only=True):
        text_id, text_title, sentence_id, sentence = row[:4]
        tid = to_int(text_id)
        sid = to_int(sentence_id)
        sentence_text = normalize_cell_text(sentence)
        if tid is None or sid is None or not sentence_text:
            continue

        text_data = texts.setdefault(
            tid,
            {
                "text_id": tid,
                "text_title": str(text_title or f"篇章{tid}"),
                "sentences": [],
            },
        )

        text_data["sentences"].append(
            {
                "sentence_id": sid,
                "sentence": sentence_text,
                "has_note": (tid, sid) in sentence_note_keys,
            }
        )

    text_list = []
    for _, text_data in sorted(texts.items(), key=lambda item: item[0]):
        text_data["sentences"].sort(key=lambda s: s["sentence_id"])
        text_list.append(text_data)

    payload = {
        "meta": {
            "source": SOURCE_FILE,
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "text_count": len(text_list),
            "sentence_count": sum(len(item["sentences"]) for item in text_list),
            "annotation_key_count": len(annotations_by_key),
            "interpretation_count": len(interpretations),
        },
        "texts": text_list,
        "annotations_by_key": annotations_by_key,
        "interpretations": interpretations,
    }

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write("window.ZZ_DATA = ")
        json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))
        f.write(";\n")

    print(
        "Generated",
        OUTPUT_FILE,
        f"with {payload['meta']['sentence_count']} sentences and",
        f"{payload['meta']['interpretation_count']} interpretations.",
    )


if __name__ == "__main__":
    main()
