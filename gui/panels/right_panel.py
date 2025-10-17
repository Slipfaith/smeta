from typing import TYPE_CHECKING

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QLabel,
    QPushButton,
    QScrollArea,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from gui.additional_services import AdditionalServicesWidget
from gui.project_setup_widget import ProjectSetupWidget
from gui.styles import (
    DROP_HINT_LABEL_STYLE,
    RIGHT_PANEL_MAIN_SPACING,
    SUMMARY_HINT_LABEL_STYLE,
    TOTAL_LABEL_STYLE,
)
from logic.translation_config import tr

if TYPE_CHECKING:  # pragma: no cover - only for type checking
    from gui.main_window import TranslationCostCalculator


def create_right_panel(window: "TranslationCostCalculator") -> QWidget:
    widget = QWidget()
    layout = QVBoxLayout()
    layout.setSpacing(RIGHT_PANEL_MAIN_SPACING)
    window.tabs = QTabWidget()
    gui_lang = window.gui_lang
    estimate_lang = "ru" if window.lang_display_ru else "en"

    window.pairs_scroll = QScrollArea()
    window.pairs_scroll.setWidgetResizable(True)
    window.pairs_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
    window.pairs_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

    window.pairs_container_widget = QWidget()
    window.pairs_layout = QVBoxLayout()
    window.pairs_layout.setSpacing(RIGHT_PANEL_MAIN_SPACING)

    window.only_new_repeats_btn = QPushButton(
        tr("Только новые слова и повторы", gui_lang)
    )
    window.only_new_repeats_btn.clicked.connect(window.toggle_only_new_repeats_mode)
    window.pairs_layout.addWidget(window.only_new_repeats_btn)

    window.delete_all_pairs_btn = QPushButton(
        tr("Удалить все языки", gui_lang)
    )
    window.delete_all_pairs_btn.setEnabled(False)
    window.delete_all_pairs_btn.clicked.connect(window.delete_all_language_pairs)
    window.pairs_layout.addWidget(window.delete_all_pairs_btn)

    window.project_setup_widget = ProjectSetupWidget(
        window.project_setup_fee_spin.value(),
        window.currency_symbol,
        window.get_current_currency_code(),
        lang=estimate_lang,
    )
    window.project_setup_widget.remove_requested.connect(
        window.remove_project_setup_widget
    )
    window.project_setup_widget.subtotal_changed.connect(window.update_total)
    window.pairs_layout.addWidget(window.project_setup_widget)
    window.project_setup_fee_spin.valueChanged.connect(
        window.update_project_setup_volume_from_spin
    )
    window.project_setup_widget.table.itemChanged.connect(
        window.on_project_setup_item_changed
    )

    window.drop_hint_label = QLabel(
        tr(
            "Перетащите XML файлы отчетов Trados или Smartcat сюда для автоматического заполнения",
            gui_lang,
        )
    )
    window.drop_hint_label.setStyleSheet(DROP_HINT_LABEL_STYLE)
    window.drop_hint_label.setAlignment(Qt.AlignCenter)
    window.pairs_layout.addWidget(window.drop_hint_label)

    window.pairs_layout.addStretch()

    window.pairs_container_widget.setLayout(window.pairs_layout)
    window.pairs_scroll.setWidget(window.pairs_container_widget)

    window.pairs_scroll.setAcceptDrops(True)
    window.setup_drag_drop()

    window.tabs.addTab(window.pairs_scroll, tr("Языковые пары", gui_lang))

    window.additional_services_widget = AdditionalServicesWidget(
        window.currency_symbol,
        window.get_current_currency_code(),
        lang=estimate_lang,
    )
    window.additional_services_widget.subtotal_changed.connect(window.update_total)
    additional_services_scroll = QScrollArea()
    additional_services_scroll.setWidget(window.additional_services_widget)
    additional_services_scroll.setWidgetResizable(True)
    window.tabs.addTab(additional_services_scroll, tr("Дополнительные услуги", gui_lang))

    layout.addWidget(window.tabs)

    window.markup_total_label.setAlignment(Qt.AlignRight)
    window.markup_total_label.setStyleSheet(SUMMARY_HINT_LABEL_STYLE)
    window.markup_total_label.hide()
    layout.addWidget(window.markup_total_label)

    window.discount_total_label.setAlignment(Qt.AlignRight)
    window.discount_total_label.setStyleSheet(SUMMARY_HINT_LABEL_STYLE)
    window.discount_total_label.hide()
    layout.addWidget(window.discount_total_label)

    window.total_label.setAlignment(Qt.AlignRight)
    window.total_label.setStyleSheet(TOTAL_LABEL_STYLE)
    layout.addWidget(window.total_label)

    widget.setLayout(layout)
    window.update_total()
    return widget
