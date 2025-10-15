from typing import TYPE_CHECKING

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDoubleSpinBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QSlider,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from gui.drop_areas import ProjectInfoDropArea
from gui.styles import (
    LEFT_PANEL_ADD_LANG_SECTION_SPACING,
    LEFT_PANEL_LANG_MODE_SLIDER_WIDTH,
    LEFT_PANEL_MAIN_SPACING,
    LEFT_PANEL_PAIRS_LIST_MAX_HEIGHT,
    LEFT_PANEL_PAIRS_SECTION_SPACING,
    LEFT_PANEL_PROJECT_SECTION_SPACING,
)
from logic.translation_config import tr

if TYPE_CHECKING:  # pragma: no cover - only for type checking
    from gui.main_window import TranslationCostCalculator


def create_left_panel(window: "TranslationCostCalculator") -> QWidget:
    container = QWidget()
    layout = QVBoxLayout()
    layout.setSpacing(LEFT_PANEL_MAIN_SPACING)

    lang = window.gui_lang

    window.project_group = ProjectInfoDropArea(
        tr("Информация о проекте", lang),
        window.handle_project_info_drop,
        lambda: window.gui_lang,
    )
    project_layout = QVBoxLayout()
    project_layout.setSpacing(LEFT_PANEL_PROJECT_SECTION_SPACING)

    window.project_name_label = QLabel(tr("Название проекта", lang) + ":")
    project_layout.addWidget(window.project_name_label)
    window.project_name_edit = QLineEdit()
    project_layout.addWidget(window.project_name_edit)

    window.client_name_label = QLabel(tr("Название клиента", lang) + ":")
    project_layout.addWidget(window.client_name_label)
    window.client_name_edit = QLineEdit()
    project_layout.addWidget(window.client_name_edit)

    window.contact_person_label = QLabel(tr("Контактное лицо", lang) + ":")
    project_layout.addWidget(window.contact_person_label)
    window.contact_person_edit = QLineEdit()
    project_layout.addWidget(window.contact_person_edit)

    window.email_label = QLabel(tr("Email", lang) + ":")
    project_layout.addWidget(window.email_label)
    window.email_edit = QLineEdit()
    project_layout.addWidget(window.email_edit)

    window.legal_entity_label = QLabel(tr("Юрлицо", lang) + ":")
    project_layout.addWidget(window.legal_entity_label)
    window.legal_entity_combo = QComboBox()
    window.legal_entity_placeholder = tr("Выберите юрлицо", lang)
    window.legal_entity_combo.addItem(window.legal_entity_placeholder)
    window.legal_entity_combo.addItems(window.legal_entities.keys())
    window.legal_entity_combo.setCurrentIndex(0)
    window.legal_entity_combo.currentTextChanged.connect(
        window.on_legal_entity_changed
    )
    project_layout.addWidget(window.legal_entity_combo)

    window.currency_label = QLabel(tr("Валюта", lang) + ":")
    project_layout.addWidget(window.currency_label)
    window.currency_combo = QComboBox()
    window.currency_placeholder = tr("Выберите валюту", lang)
    window.currency_combo.addItem(window.currency_placeholder)
    window.currency_combo.addItems(["RUB", "EUR", "USD"])
    window.currency_combo.setCurrentIndex(0)
    window.currency_combo.currentIndexChanged.connect(
        window.on_currency_index_changed
    )
    project_layout.addWidget(window.currency_combo)

    window.convert_btn = QPushButton(tr("Конвертировать в рубли", lang))
    window.convert_btn.clicked.connect(window.convert_to_rub)
    project_layout.addWidget(window.convert_btn)

    window.vat_label = QLabel(tr("НДС, %", lang) + ":")
    window.vat_spin = QDoubleSpinBox()
    window.vat_spin.setDecimals(2)
    window.vat_spin.setRange(0, 100)
    window.vat_spin.setValue(20.0)
    window.vat_spin.valueChanged.connect(window.update_total)
    window.vat_spin.wheelEvent = lambda event: event.ignore()

    vat_layout = QHBoxLayout()
    vat_layout.addWidget(window.vat_label)
    vat_layout.addWidget(window.vat_spin)
    project_layout.addLayout(vat_layout)

    window.project_group.setLayout(project_layout)
    layout.addWidget(window.project_group)

    window.on_legal_entity_changed("")
    window.on_currency_changed(window.get_current_currency_code())

    window.pairs_group = QGroupBox(tr("Языковые пары", lang))
    pairs_layout = QVBoxLayout()
    pairs_layout.setSpacing(LEFT_PANEL_PAIRS_SECTION_SPACING)

    mode_layout = QHBoxLayout()
    window.language_names_label = QLabel(tr("Названия языков", lang) + ":")
    mode_layout.addWidget(window.language_names_label)
    mode_layout.addStretch(1)
    mode_layout.addWidget(QLabel("EN"))
    window.lang_mode_slider = QSlider(Qt.Horizontal)
    window.lang_mode_slider.setRange(0, 1)
    window.lang_mode_slider.setValue(1)
    window.lang_mode_slider.setFixedWidth(LEFT_PANEL_LANG_MODE_SLIDER_WIDTH)
    window.lang_mode_slider.valueChanged.connect(window.on_lang_mode_changed)
    mode_layout.addWidget(window.lang_mode_slider)
    mode_layout.addWidget(QLabel("RU"))
    pairs_layout.addLayout(mode_layout)

    add_pair_layout = QHBoxLayout()
    window.source_lang_combo = window._make_lang_combo()
    window.source_lang_combo.setEditable(True)
    add_pair_layout.addWidget(window.source_lang_combo)
    add_pair_layout.addWidget(QLabel("→"))
    window.target_lang_combo = window._make_lang_combo()
    window.target_lang_combo.setEditable(True)
    add_pair_layout.addWidget(window.target_lang_combo)
    pairs_layout.addLayout(add_pair_layout)

    window.add_pair_btn = QPushButton(tr("Добавить языковую пару", lang))
    window.add_pair_btn.clicked.connect(window.add_language_pair)
    pairs_layout.addWidget(window.add_pair_btn)

    window.current_pairs_label = QLabel(tr("Текущие пары", lang) + ":")
    pairs_layout.addWidget(window.current_pairs_label)
    window.pairs_list = QTextEdit()
    window.pairs_list.setMaximumHeight(LEFT_PANEL_PAIRS_LIST_MAX_HEIGHT)
    window.pairs_list.setReadOnly(True)
    pairs_layout.addWidget(window.pairs_list)

    info_layout = QHBoxLayout()
    window.language_pairs_count_label = QLabel(
        f"{tr('Загружено языковых пар', lang)}: 0"
    )
    info_layout.addWidget(window.language_pairs_count_label)
    info_layout.addStretch()
    window.clear_pairs_btn = QPushButton(tr("Очистить", lang))
    window.clear_pairs_btn.clicked.connect(window.clear_language_pairs)
    info_layout.addWidget(window.clear_pairs_btn)
    pairs_layout.addLayout(info_layout)

    setup_layout = QHBoxLayout()
    window.project_setup_label = QLabel(
        tr("Запуск и управление проектом", lang) + ":",
    )
    setup_layout.addWidget(window.project_setup_label)
    window.project_setup_fee_spin = QDoubleSpinBox()
    window.project_setup_fee_spin.setDecimals(2)
    window.project_setup_fee_spin.setSingleStep(0.25)
    window.project_setup_fee_spin.setMinimum(0.5)
    window.project_setup_fee_spin.setValue(0.5)
    setup_layout.addWidget(window.project_setup_fee_spin)
    setup_layout.addStretch()
    pairs_layout.addLayout(setup_layout)

    window.add_lang_group = QGroupBox(tr("Добавить язык в справочник", lang))
    add_lang_layout = QVBoxLayout()
    add_lang_layout.setSpacing(LEFT_PANEL_ADD_LANG_SECTION_SPACING)

    row_ru = QHBoxLayout()
    window.lang_ru_label = QLabel(tr("Название RU", lang) + ":")
    row_ru.addWidget(window.lang_ru_label)
    window.new_lang_ru = QLineEdit()
    window.new_lang_ru.setPlaceholderText(tr("Валирийский", "ru"))
    row_ru.addWidget(window.new_lang_ru)
    add_lang_layout.addLayout(row_ru)

    row_en = QHBoxLayout()
    window.lang_en_label = QLabel(tr("Название EN", lang) + ":")
    row_en.addWidget(window.lang_en_label)
    window.new_lang_en = QLineEdit()
    window.new_lang_en.setPlaceholderText(tr("Valyrian", "en"))
    row_en.addWidget(window.new_lang_en)
    add_lang_layout.addLayout(row_en)

    window.btn_add_lang = QPushButton(tr("Добавить язык", lang))
    window.btn_add_lang.clicked.connect(window.handle_add_language)
    add_lang_layout.addWidget(window.btn_add_lang)

    window.add_lang_group.setLayout(add_lang_layout)
    pairs_layout.addWidget(window.add_lang_group)

    window.pairs_group.setLayout(pairs_layout)
    layout.addWidget(window.pairs_group)

    layout.addStretch()
    container.setLayout(layout)

    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setWidget(container)
    scroll.setMinimumWidth(280)
    return scroll
