"""Utility helpers for currency and total calculations."""

from __future__ import annotations

from typing import Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QInputDialog

from gui.utils import format_amount
from logic.translation_config import tr
from logic.activity_logger import log_user_action, log_window_action

CURRENCY_SYMBOLS = {"RUB": "₽", "EUR": "€", "USD": "$"}


def set_currency_code(window, code: Optional[str]) -> bool:
    """Set currency combo box to the provided code.

    Parameters
    ----------
    window: TranslationCostCalculator
        Main window instance that exposes ``currency_combo``.
    code: Optional[str]
        Currency code to select in the combo box. ``None`` or empty value resets
        the selection.
    """

    combo = window.currency_combo
    if code:
        normalized = str(code).strip().upper()
        idx = combo.findText(normalized, Qt.MatchFixedString)
        if idx < 0:
            for i in range(1, combo.count()):
                text = combo.itemText(i).strip().upper()
                if text == normalized:
                    idx = i
                    break
        if idx >= 0:
            combo.setCurrentIndex(idx)
            return True
    combo.setCurrentIndex(0)
    return False


def on_currency_changed(window, code: str):
    """Propagate currency change across widgets."""

    window.currency_symbol = CURRENCY_SYMBOLS.get(code, code)
    if getattr(window, "project_setup_widget", None):
        window.project_setup_widget.set_currency(window.currency_symbol, code)
    for widget in window.language_pairs.values():
        widget.set_currency(window.currency_symbol, code)
    if getattr(window, "additional_services_widget", None):
        window.additional_services_widget.set_currency(window.currency_symbol, code)
    if getattr(window, "convert_btn", None):
        window.convert_btn.setEnabled(code == "USD")

    log_window_action(
        "Изменена валюта проекта",
        window,
        details={"Новый код": code or ""},
    )


def convert_to_rub(window):
    """Convert all rates from USD to RUB using a provided rate."""

    if window.get_current_currency_code() != "USD":
        return

    lang = window.gui_lang
    rate, ok = QInputDialog.getDouble(
        window,
        tr("Курс USD", lang),
        tr("1 USD в рублях", lang),
        0.0,
        0.0,
        1_000_000.0,
        4,
    )
    if not ok or rate <= 0:
        log_user_action(
            "Конвертация USD→RUB отменена",
            details={"Подтверждение": ok, "Введённый курс": rate},
        )
        return

    if getattr(window, "project_setup_widget", None):
        window.project_setup_widget.convert_rates(rate)
    for widget in window.language_pairs.values():
        widget.convert_rates(rate)
    if getattr(window, "additional_services_widget", None):
        window.additional_services_widget.convert_rates(rate)

    set_currency_code(window, "RUB")
    update_total(window)

    log_window_action(
        "Все ставки пересчитаны в рубли",
        window,
        details={"Курс": rate},
    )


def update_total(window, *_):
    """Recalculate totals and update labels on the window."""

    total = 0.0
    discount_total = 0.0
    markup_total = 0.0

    if getattr(window, "project_setup_widget", None):
        total += window.project_setup_widget.get_subtotal()
        discount_total += window.project_setup_widget.get_discount_amount()
        markup_total += window.project_setup_widget.get_markup_amount()

    for widget in window.language_pairs.values():
        total += widget.get_subtotal()
        discount_total += widget.get_discount_amount()
        markup_total += widget.get_markup_amount()

    if getattr(window, "additional_services_widget", None):
        total += window.additional_services_widget.get_subtotal()
        discount_total += window.additional_services_widget.get_discount_amount()
        markup_total += window.additional_services_widget.get_markup_amount()

    lang = "ru" if window.lang_display_ru else "en"
    vat_rate = window.vat_spin.value() / 100 if window.vat_spin.isEnabled() else 0.0

    symbol_suffix = f" {window.currency_symbol}" if window.currency_symbol else ""

    if markup_total > 0:
        window.markup_total_label.setText(
            f"{tr('Сумма наценки', lang)}: {format_amount(markup_total, lang)}{symbol_suffix}"
        )
        window.markup_total_label.show()
    else:
        window.markup_total_label.hide()

    if discount_total > 0:
        window.discount_total_label.setText(
            f"{tr('Сумма скидки', lang)}: {format_amount(discount_total, lang)}{symbol_suffix}"
        )
        window.discount_total_label.show()
    else:
        window.discount_total_label.hide()

    if vat_rate > 0:
        vat_amount = total * vat_rate
        total_with_vat = total + vat_amount
        window.total_label.setText(
            f"{tr('Итого', lang)}: {format_amount(total_with_vat, lang)}{symbol_suffix} {tr('с НДС', lang)}. "
            f"{tr('НДС', lang)}: {format_amount(vat_amount, lang)}{symbol_suffix}"
        )
    else:
        window.total_label.setText(
            f"{tr('Итого', lang)}: {format_amount(total, lang)}{symbol_suffix}"
        )

    log_window_action(
        "Пересчитаны итоги",
        window,
        details={
            "Итого": round(total, 2),
            "Скидки": round(discount_total, 2),
            "Наценки": round(markup_total, 2),
            "НДС": round(vat_rate * 100, 2),
        },
    )
