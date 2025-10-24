from __future__ import annotations

from typing import Any, Dict, Tuple

from PySide6.QtWidgets import QMessageBox, QComboBox

from gui.language_pair import LanguagePairWidget
from logic.translation_config import tr
from logic.xml_parser_common import resolve_language_display
from logic.activity_logger import log_window_action


class LanguagePairsMixin:
    def _labels_from_entry(self, entry: Dict[str, Any]) -> Dict[str, str]:
        if not entry:
            return {"en": "", "ru": ""}

        key_value = entry.get("key", "").strip()
        text = entry.get("text", "").strip()
        en_value = entry.get("en", "").strip()
        ru_value = entry.get("ru", "").strip()

        resolved_en = ""
        resolved_ru = ""
        if key_value:
            resolved_en = resolve_language_display(key_value, locale="en") or ""
            resolved_ru = resolve_language_display(key_value, locale="ru") or ""
        if not resolved_en and text:
            resolved_en = resolve_language_display(text, locale="en") or ""
        if not resolved_ru and text:
            resolved_ru = resolve_language_display(text, locale="ru") or ""
        if not resolved_en and en_value:
            resolved_en = resolve_language_display(en_value, locale="en") or ""
        if not resolved_ru and ru_value:
            resolved_ru = resolve_language_display(ru_value, locale="ru") or ""

        en_name = resolved_en or en_value or ru_value or text or key_value
        ru_name = resolved_ru or ru_value or en_value or text or key_value

        return {
            "en": self._prepare_language_label(en_name, "en"),
            "ru": self._prepare_language_label(ru_name, "ru"),
        }

    def add_language_pair(self):
        src = self._parse_combo(self.source_lang_combo)
        tgt = self._parse_combo(self.target_lang_combo)
        if not src["text"] or not tgt["text"]:
            lang = self.gui_lang
            QMessageBox.warning(
                self, tr("Ошибка", lang), tr("Выберите/введите оба языка", lang)
            )
            return

        def key_name(obj: Dict[str, Any]) -> str:
            return obj["en"] or obj["ru"] or obj["text"]

        left_key = key_name(src)
        right_key = key_name(tgt)
        pair_key = f"{left_key} → {right_key}"

        if pair_key in self.language_pairs:
            lang = self.gui_lang
            QMessageBox.warning(
                self, tr("Ошибка", lang), tr("Такая языковая пара уже существует", lang)
            )
            return

        self._store_pair_language_inputs(pair_key, src, tgt, left_key, right_key)

        labels = self._pair_language_inputs.get(pair_key, {})
        src_labels = self._labels_from_entry(labels.get("source", {}))
        tgt_labels = self._labels_from_entry(labels.get("target", {}))
        lang_key = "ru" if self.lang_display_ru else "en"
        locale = "ru" if self.lang_display_ru else "en"
        src_value = src_labels[lang_key]
        tgt_value = tgt_labels[lang_key]
        if not src_value:
            src_value = self._prepare_language_label(src.get("text", ""), locale)
        if not tgt_value:
            tgt_value = self._prepare_language_label(tgt.get("text", ""), locale)
        display_name = f"{src_value} - {tgt_value}"
        header_value = tgt_labels[lang_key]
        if not header_value:
            locale = "ru" if self.lang_display_ru else "en"
            header_value = self._prepare_language_label(tgt.get("text", ""), locale)
        self.pair_headers[pair_key] = header_value

        widget = LanguagePairWidget(
            display_name,
            self.currency_symbol,
            self.get_current_currency_code(),
            lang="ru" if self.lang_display_ru else "en",
        )
        widget.remove_requested.connect(
            lambda w=widget: self._on_widget_remove_requested(w)
        )
        widget.subtotal_changed.connect(self.update_total)
        widget.name_changed.connect(
            lambda new_name, w=widget: self.on_pair_name_changed(w, new_name)
        )
        self.language_pairs[pair_key] = widget
        if self.only_new_repeats_mode:
            widget.set_only_new_and_repeats_mode(True)

        self.pairs_layout.insertWidget(self.pairs_layout.count() - 1, widget)

        self.update_pairs_list()
        self.update_total()

        self._update_language_variant_regions_from_pairs(self.language_pairs.keys())

        self._reset_language_pair_inputs()
        log_window_action(
            "Добавлена языковая пара",
            self,
            details={"Пара": pair_key, "Отображение": display_name},
        )

    def _extract_pair_parts(self, pair_key: str) -> Tuple[str, str]:
        for sep in (" → ", " - "):
            if sep in pair_key:
                left, right = pair_key.split(sep, 1)
                return left.strip(), right.strip()
        return pair_key.strip(), ""

    def _display_pair_name(self, pair_key: str) -> str:
        lang = "ru" if self.lang_display_ru else "en"
        entries = self._pair_language_inputs.get(pair_key)
        if entries:
            left_labels = self._labels_from_entry(entries.get("source", {}))
            right_labels = self._labels_from_entry(entries.get("target", {}))
            if right_labels.get("en") or right_labels.get("ru"):
                return f"{left_labels[lang]} - {right_labels[lang]}"
            return left_labels[lang]

        left_key, right_key = self._extract_pair_parts(pair_key)
        if not right_key:
            return self._prepare_language_label(left_key, lang)
        left = self._find_language_by_key(left_key)
        right = self._find_language_by_key(right_key)
        return f"{left[lang]} - {right[lang]}"

    def remove_language_pair(self, pair_key: str):
        widget = self.language_pairs.pop(pair_key, None)
        if widget:
            widget.setParent(None)
            self.pair_headers.pop(pair_key, None)
        self._pair_language_inputs.pop(pair_key, None)
        self.update_pairs_list()
        self.update_total()

        self._update_language_variant_regions_from_pairs(self.language_pairs.keys())
        log_window_action(
            "Удалена языковая пара",
            self,
            details={"Пара": pair_key},
        )

    def clear_language_pairs(self):
        for w in self.language_pairs.values():
            w.setParent(None)
        removed = list(self.language_pairs.keys())
        self.language_pairs.clear()
        self.pair_headers.clear()
        self._pair_language_inputs.clear()
        self.update_pairs_list()
        self.update_total()

        self._update_language_variant_regions_from_pairs([])
        log_window_action(
            "Удалены все языковые пары",
            self,
            details={"Количество": len(removed)},
        )

    def _reset_language_pair_inputs(self) -> None:
        """Return language combo boxes to their default selections."""

        self._select_language_in_combo(self.target_lang_combo, "English")

    def _select_language_in_combo(self, combo: QComboBox, en_name: str) -> None:
        """Select a language in ``combo`` by its English name.

        Falls back to the first available option when the requested language is
        missing. If the combo box is editable and empty, its edit text is
        cleared.
        """

        normalized = en_name.strip().lower()
        target_index = -1
        for idx in range(combo.count()):
            data = combo.itemData(idx)
            if (
                isinstance(data, dict)
                and str(data.get("en", "")).strip().lower() == normalized
            ):
                target_index = idx
                break

        combo.blockSignals(True)
        try:
            if target_index >= 0:
                combo.setCurrentIndex(target_index)
            elif combo.count() > 0:
                combo.setCurrentIndex(0)
            else:
                combo.setCurrentIndex(-1)
            if combo.isEditable() and combo.currentIndex() < 0:
                combo.setEditText("")
        finally:
            combo.blockSignals(False)
