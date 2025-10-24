"""Data structures describing a translation project."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any, Dict, Iterable, List, Mapping


def _slugify(value: str) -> str:
    """Return a file-system-friendly representation of *value*."""

    return value.replace(" ", "_")


@dataclass(frozen=True)
class ProjectData:
    """Structured snapshot of the data required to export a project."""

    project_name: str
    client_name: str
    contact_person: str
    email: str
    legal_entity: str
    currency: str
    language_pairs: List[Dict[str, Any]] = field(default_factory=list)
    additional_services: List[Dict[str, Any]] = field(default_factory=list)
    pm_name: str = ""
    pm_email: str = ""
    project_setup_fee: float = 0.0
    project_setup_enabled: bool = False
    project_setup: List[Dict[str, Any]] = field(default_factory=list)
    project_setup_discount_percent: float = 0.0
    project_setup_markup_percent: float = 0.0
    project_setup_discount_amount: float = 0.0
    project_setup_markup_amount: float = 0.0
    vat_rate: float = 0.0
    only_new_repeats_mode: bool = False
    total_discount_amount: float = 0.0
    total_markup_amount: float = 0.0

    @classmethod
    def from_window(cls, window: Any) -> "ProjectData":
        """Collect the current project data from the UI *window*."""

        language_pairs: List[Dict[str, Any]] = []
        total_discount_amount = 0.0
        total_markup_amount = 0.0
        project_setup_discount_amount = 0.0
        project_setup_markup_amount = 0.0

        project_setup = []
        project_setup_widget = getattr(window, "project_setup_widget", None)
        if project_setup_widget:
            project_setup_enabled = project_setup_widget.is_enabled()
            if project_setup_enabled:
                project_setup = project_setup_widget.get_data()
                project_setup_discount_amount = project_setup_widget.get_discount_amount()
                project_setup_markup_amount = project_setup_widget.get_markup_amount()
                total_discount_amount += project_setup_discount_amount
                total_markup_amount += project_setup_markup_amount
            project_setup_discount_percent = project_setup_widget.get_discount_percent()
            project_setup_markup_percent = project_setup_widget.get_markup_percent()
        else:
            project_setup_enabled = False
            project_setup_discount_percent = 0.0
            project_setup_markup_percent = 0.0

        for pair_key, pair_widget in window.language_pairs.items():
            pair_data = pair_widget.get_data()
            if pair_data.get("services"):
                pair_data = dict(pair_data)
                pair_data["header_title"] = window.pair_headers.get(
                    pair_key, pair_widget.pair_name
                )
                language_pairs.append(pair_data)
                total_discount_amount += pair_data.get("discount_amount", 0.0)
                total_markup_amount += pair_data.get("markup_amount", 0.0)

        def _sort_key(pair: Mapping[str, Any]) -> str:
            pair_name = pair.get("pair_name", "")
            if " - " in pair_name:
                return pair_name.split(" - ", maxsplit=1)[1]
            return pair_name

        language_pairs.sort(key=_sort_key)

        additional_services_widget = getattr(
            window, "additional_services_widget", None
        )
        if additional_services_widget is not None:
            additional_services = additional_services_widget.get_data()
            if additional_services:
                total_discount_amount += sum(
                    block.get("discount_amount", 0.0) for block in additional_services
                )
                total_markup_amount += sum(
                    block.get("markup_amount", 0.0) for block in additional_services
                )
        else:
            additional_services = []

        return cls(
            project_name=window.project_name_edit.text(),
            client_name=window.client_name_edit.text(),
            contact_person=window.contact_person_edit.text(),
            email=window.email_edit.text(),
            legal_entity=window.get_selected_legal_entity(),
            currency=window.get_current_currency_code(),
            language_pairs=language_pairs,
            additional_services=additional_services,
            pm_name=window.current_pm.get(
                "name_ru" if window.lang_display_ru else "name_en", ""
            ),
            pm_email=window.current_pm.get("email", ""),
            project_setup_fee=window.project_setup_fee_spin.value(),
            project_setup_enabled=project_setup_enabled,
            project_setup=project_setup,
            project_setup_discount_percent=project_setup_discount_percent,
            project_setup_markup_percent=project_setup_markup_percent,
            project_setup_discount_amount=project_setup_discount_amount,
            project_setup_markup_amount=project_setup_markup_amount,
            vat_rate=window.vat_spin.value() if window.vat_spin.isEnabled() else 0.0,
            only_new_repeats_mode=window.only_new_repeats_mode,
            total_discount_amount=total_discount_amount,
            total_markup_amount=total_markup_amount,
        )

    def to_mapping(self) -> Dict[str, Any]:
        """Return a serialisable representation of the project data."""

        return asdict(self)

    @property
    def project_slug(self) -> str:
        return _slugify(self.project_name)

    @property
    def client_slug(self) -> str:
        return _slugify(self.client_name)

    @property
    def legal_entity_slug(self) -> str:
        return _slugify(self.legal_entity)

    def has_any_services(self) -> bool:
        return bool(self.language_pairs or self.additional_services)

    def has_zero_setup_rates(self) -> bool:
        return _any_zero_rate(self.project_setup)


def _any_zero_rate(setup_items: Iterable[Mapping[str, Any]]) -> bool:
    return any(item.get("rate", 0) == 0 for item in setup_items)
