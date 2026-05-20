import base64
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from io import BytesIO
from math import ceil
from pathlib import Path
from zipfile import BadZipFile

import pytz
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import load_workbook
from PIL import Image as PILImage


# ============================================================================
# Configuracoes principais
# ============================================================================

APP_TITLE = "Trebeschi Comercial"
BASE_DIR = Path(__file__).resolve().parent

EXCEL_PATH = BASE_DIR / "disponibilidade.xlsx"

LOGO_PATH = BASE_DIR / "assets" / "trebeschi_logo.png"

AUTO_REFRESH_SECONDS = 120
BRASILIA_TZ = pytz.timezone("America/Sao_Paulo")

CULTURE_ICONS = {
    "Tomate": "🍅",
    "Cebola": "🧅",
    "Batata Doce": "🍠",
    "Alho": "🧄",
}

EXCEL_SHEET_ALIASES = {
    "Batata Doce Rosada": "Batata Doce",
}

TAXES = {
    "Funrural (1,5%)": 0.015,
    "Pessoa Jurídica (3,28%)": 0.0328,
    "Pernambuco (21%)": 0.21,
}

PACKAGE_COSTS = {
    "Papelão": 8.00,
    "HB": 4.50,
    "IFCO": 4.50,
    "Plástica": 1.00,
    "SC": 0.00,
}


# ============================================================================
# Catalogo de negocio
# ============================================================================

@dataclass(frozen=True)
class LoadingLocation:
    """Representa um filtro de local no app e os locais correspondentes no Excel."""

    name: str
    cultures: list[str]
    excel_locations: list[str]


class AppCatalog:
    """Centraliza a ordem de culturas e a relacao entre filtros e planilha."""

    def __init__(self) -> None:
        self.culture_order = ["Tomate", "Alho", "Cebola", "Batata Doce"]
        self.locations = {
            "Brasil": LoadingLocation(
                name="Brasil",
                cultures=self.culture_order,
                excel_locations=[],
            ),
            "Araguari": LoadingLocation(
                name="Araguari",
                cultures=["Tomate", "Alho", "Cebola"],
                excel_locations=["Araguari", "Trebeschi", "Mandaguari"],
            ),
            "Anápolis": LoadingLocation(
                name="Anápolis",
                cultures=["Tomate", "Batata Doce"],
                excel_locations=["Anápolis"],
            ),
            "Cascavel": LoadingLocation(
                name="Cascavel",
                cultures=["Tomate"],
                excel_locations=["Cascavel"],
            ),
        }

    def location_names(self) -> list[str]:
        return list(self.locations.keys())

    def cultures_for_location(self, location: str, items: list[dict]) -> list[str]:
        if location == "Brasil":
            present = {item["cultura"] for item in items}
            return [culture for culture in self.culture_order if culture in present]

        selected = self.locations.get(location)
        return selected.cultures if selected else []

    def excel_locations_for_filter(self, location: str) -> list[str]:
        selected = self.locations.get(location)
        return selected.excel_locations if selected else [location]


# ============================================================================
# Leitura da disponibilidade
# ============================================================================

class ExcelAvailabilityRepository:
    """Le a planilha disponibilidade.xlsx no formato de blocos por cultura/packing."""

    def __init__(self, excel_path: Path, culture_names: list[str]) -> None:
        self.excel_path = excel_path
        self.culture_names = set(culture_names)

    def list_items(self) -> list[dict]:
        if not self.excel_path.exists():
            st.warning(f"Planilha nao encontrada: {self.excel_path}")
            return []

        try:
            workbook = load_workbook(self.excel_path, data_only=True, read_only=True)
        except (PermissionError, OSError, BadZipFile):
            st.warning(
                "A planilha esta sendo salva pelo Excel. Aguarde a proxima atualizacao automatica."
            )
            return []

        try:
            items = []
            for sheet_name in workbook.sheetnames:
                culture_name = EXCEL_SHEET_ALIASES.get(sheet_name, sheet_name)
                if culture_name in self.culture_names:
                    items.extend(self._read_culture_sheet(workbook[sheet_name], culture_name))
            return items
        finally:
            workbook.close()

    def _read_culture_sheet(self, sheet, culture: str) -> list[dict]:
        matrix = [list(row) for row in sheet.iter_rows(values_only=True)]
        items = []
        next_id = 1

        for row_index, row in enumerate(matrix):
            for col_index, value in enumerate(row):
                if normalize_text(value) != "disponibilidade":
                    continue

                block_items = self._read_block(
                    matrix=matrix,
                    culture=culture,
                    start_row=row_index,
                    start_col=col_index,
                    first_id=next_id,
                )
                items.extend(block_items)
                next_id += len(block_items)

        return items

    def _read_block(
        self,
        matrix: list[list],
        culture: str,
        start_row: int,
        start_col: int,
        first_id: int,
    ) -> list[dict]:
        location = cell_text(matrix, start_row + 3, start_col)
        packing = cell_text(matrix, start_row + 3, start_col + 2)
        updated_at = format_date(cell_value(matrix, start_row + 1, start_col))

        if not location or not packing:
            return []

        items = []
        row_index = start_row + 4
        while row_index < len(matrix):
            product = cell_text(matrix, row_index, start_col)
            if not product or normalize_text(product) == "total":
                break

            quantity = parse_number(cell_value(matrix, row_index, start_col + 1))
            items.append(
                {
                    "id": first_id + len(items),
                    "local_carregamento": location,
                    "cultura": culture,
                    "packing": packing,
                    "produto": product,
                    "ordem_packing": start_row * 1000 + start_col,
                    "ordem_produto": row_index - (start_row + 4),
                    "quantidade": quantity,
                    "atualizado_em": updated_at,
                }
            )
            row_index += 1

        return items


# ============================================================================
# Estilo visual
# ============================================================================

class PageStyle:
    """Aplica CSS unico para as duas funcionalidades, com suporte real a dark mode."""

    @staticmethod
    def apply() -> None:
        st.set_page_config(
            page_title=APP_TITLE,
            page_icon="📦",
            layout="centered",
        )
        st.markdown(
            """
            <meta name="theme-color" content="#176b55">
            """,
            unsafe_allow_html=True,
        )

        st.markdown(
            """
            <style>

                :root {
                    color-scheme: light dark;

                    --bg: #f5f7fa;
                    --surface: #ffffff;
                    --surface-soft: #eef4ef;

                    --text: #17212b;
                    --text-soft: #4b5563;

                    --border: #d7dee7;

                    --green: #176b55;
                    --green-bright: #00a651;

                    --blue: #2f75b5;
                    --danger: #d32f2f;

                    --input-bg: #ffffff;
                    --input-text: #17212b;
                    --input-border: #cfd8e3;

                    --table-header: #176b55;
                }

                @media (prefers-color-scheme: dark) {

                    :root {

                        --bg: #0f1419;

                        --surface: #182028;

                        --surface-soft: #212b34;

                        --text: #f3f7fb;

                        --text-soft: #c5d0db;

                        --border: #344250;

                        --green: #1f8a67;

                        --green-bright: #37d487;

                        --blue: #4593d6;

                        --danger: #ff7b72;

                        --input-bg: #111827;

                        --input-text: #f3f7fb;

                        --input-border: #3b4a5a;

                        --table-header: #1f8a67;
                    }
                }

                html,
                body,
                .stApp {
                    background: var(--bg) !important;
                    color: var(--text) !important;
                }
                .stApp {
                    background: var(--bg) !important;
                    color: var(--text) !important;
                    color-scheme: dark;
                }               
                .block-container {
                    max-width: 1040px;
                    padding-top: 2.4rem;
                    padding-bottom: 3rem;
                }

                /* =========================================
                   TEXTOS
                ========================================= */

                h1, h2, h3, h4, h5, h6,
                p, span, div, label,
                .stMarkdown,
                .stText,
                .stCaption {
                    color: var(--text) !important;
                }

                small {
                    color: var(--text-soft) !important;
                }

                /* =========================================
                   HEADER
                ========================================= */

                .app-header {
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    gap: 1.1rem;
                    margin-top: 0.4rem;
                    margin-bottom: 0.8rem;
                    text-align: center;
                }

                .app-logo {
                    width: 124px;
                    max-width: 32vw;
                    height: auto;
                    object-fit: contain;
                    display: block;
                }

                .app-title {
                    font-size: 2rem;
                    line-height: 1.1;
                    font-weight: 800;
                    color: var(--green-bright) !important;
                }

                /* =========================================
                   INPUTS STREAMLIT
                ========================================= */

                .stTextInput input,
                .stNumberInput input,
                .stSelectbox div[data-baseweb="select"],
                .stTextArea textarea {

                    background: var(--input-bg) !important;
                    color: var(--input-text) !important;

                    border: 1px solid var(--input-border) !important;

                    border-radius: 8px !important;
                }
                /* =========================================
                   DROPDOWN OPTIONS - DARK MODE FIX
                ========================================= */
                
                div[role="listbox"] {
                    background: var(--surface) !important;
                    border: 1px solid var(--border) !important;
                    border-radius: 8px !important;
                }
                
                div[role="option"] {
                    background: var(--surface) !important;
                    color: var(--text) !important;
                }
                
                /* Hover da opção */
                
                div[role="option"]:hover {
                    background: var(--surface-soft) !important;
                    color: var(--text) !important;
                }
                
                /* Opção selecionada */
                
                div[aria-selected="true"] {
                    background: var(--green) !important;
                    color: #ffffff !important;
                }
                
                /* Texto interno das opções */
                
                div[role="option"] * {
                    color: inherit !important;
                }
                
                /* Corrige texto do select fechado */
                
                div[data-baseweb="select"] span {
                    color: var(--input-text) !important;
                }
                
                /* Corrige ícone/seta */
                
                div[data-baseweb="select"] svg {
                    fill: var(--input-text) !important;
                }

                .stTextInput input::placeholder,
                .stNumberInput input::placeholder,
                textarea::placeholder {

                    color: #9aa7b5 !important;
                    opacity: 1 !important;
                }

                /* SELECTBOX */

                div[data-baseweb="select"] > div {

                    background: var(--input-bg) !important;

                    color: var(--input-text) !important;

                    border: 1px solid var(--input-border) !important;
                }

                div[data-baseweb="popover"] {

                    background: var(--surface) !important;
                    color: var(--text) !important;
                }

                /* RADIO */

                .stRadio label {
                    color: var(--text) !important;
                }

                /* =========================================
              CHECKBOX - DARK MODE FIX REAL
                ========================================= */

                /* Texto */
                
                .stCheckbox label,
                .stCheckbox span {
                    color: var(--text) !important;
                    font-weight: 600;
                }
                
                /* Caixa externa */
                
                .stCheckbox div[role="checkbox"] {
                
                    background-color: var(--input-bg) !important;
                
                    border: 2px solid var(--input-border) !important;
                
                    border-radius: 6px !important;
                
                    width: 20px !important;
                
                    height: 20px !important;
                
                    transition: all 0.15s ease;
                }
                
                /* Hover */
                
                .stCheckbox div[role="checkbox"]:hover {
                
                    border-color: var(--green-bright) !important;
                
                    box-shadow: 0 0 0 1px var(--green-bright) !important;
                }
                
                /* Marcado */
                
                .stCheckbox div[role="checkbox"][aria-checked="true"] {
                
                    background-color: var(--green) !important;
                
                    border-color: var(--green) !important;
                }
                
                /* Ícone do check */
                
                .stCheckbox div[role="checkbox"] svg {
                
                    fill: white !important;
                
                    stroke: white !important;
                
                    stroke-width: 3 !important;
                
                    width: 16px !important;
                
                    height: 16px !important;
                }
                
                /* Remove fundo estranho do BaseWeb */
                
                .stCheckbox div[data-testid="stMarkdownContainer"] {
                
                    color: var(--text) !important;
                }
                
                /* Espaçamento melhor */
                
                .stCheckbox {
                
                    padding-top: 0.2rem;
                    padding-bottom: 0.2rem;
                }

                /* =========================================
                   BOTÕES
                ========================================= */

                .stButton button,
                .stDownloadButton button,
                .stFormSubmitButton button {

                    background: var(--green) !important;

                    color: white !important;

                    border: none !important;

                    border-radius: 8px !important;

                    font-weight: 700 !important;
                }

                .stButton button:hover,
                .stDownloadButton button:hover,
                .stFormSubmitButton button:hover {

                    filter: brightness(1.08);
                }

                /* =========================================
                   MÉTRICAS
                ========================================= */

                div[data-testid="stMetric"] {

                    border: 1px solid var(--border);

                    border-radius: 10px;

                    padding: 0.8rem;

                    background: var(--surface);
                }

                div[data-testid="stMetricLabel"] {

                    color: var(--text-soft) !important;
                }

                div[data-testid="stMetricValue"] {

                    color: var(--text) !important;

                    font-weight: 800;
                }

                /* =========================================
                   TABELAS
                ========================================= */

                .availability-table {

                    width: 100%;

                    border-collapse: collapse;

                    font-size: 0.92rem;

                    margin-bottom: 1rem;

                    background: var(--surface);

                    border: 1px solid var(--border);

                    border-radius: 8px;

                    overflow: hidden;
                }

                .availability-table th {

                    background: var(--table-header);

                    color: white;

                    font-weight: 800;

                    text-align: center;

                    padding: 0.45rem 0.5rem;

                    border: 1px solid var(--border);
                }

                .availability-table td {

                    color: var(--text);

                    text-align: center;

                    padding: 0.42rem 0.5rem;

                    border: 1px solid var(--border);
                }

                .availability-table tbody tr:nth-child(odd) td {

                    background: var(--surface-soft);
                }

                .availability-table tbody tr.total-row td {

                    background: var(--blue);

                    color: white;

                    font-weight: 800;
                }

                /* =========================================
                   CARDS DE COTAÇÃO
                ========================================= */

                .quote-result {

                    text-align: center;

                    padding: 0.9rem;

                    margin-top: 1rem;

                    border: 1px solid var(--border);

                    border-radius: 10px;

                    background: var(--surface);
                }

                .quote-price {

                    color: var(--danger);

                    font-size: 1.7rem;

                    font-weight: 800;

                    margin: 0.35rem 0;
                }

                .quote-title,
                .quote-section-title {

                    color: var(--green-bright) !important;
                }

                /* =========================================
                   ALERTAS
                ========================================= */

                .stAlert {
                    border-radius: 10px !important;
                }

                /* =========================================
                   FOOTER
                ========================================= */

                .app-footer {

                    margin-top: 2rem;

                    padding-top: 0.85rem;

                    border-top: 1px solid var(--border);

                    color: var(--text-soft);

                    text-align: center;

                    font-size: 0.85rem;
                }

                /* =========================================
                   FOCUS IOS/ANDROID
                ========================================= */

                input:focus,
                textarea:focus,
                select:focus {

                    outline: none !important;

                    border: 1px solid var(--green-bright) !important;

                    box-shadow: 0 0 0 1px var(--green-bright) !important;
                }

                /* =========================================
                   AUTOFILL ANDROID
                ========================================= */

                input:-webkit-autofill,
                input:-webkit-autofill:hover,
                input:-webkit-autofill:focus {

                    -webkit-text-fill-color: var(--input-text);

                    -webkit-box-shadow:
                        0 0 0px 1000px var(--input-bg) inset;

                    transition: background-color 9999s ease-in-out 0s;
                }

                /* =========================================
                   MOBILE
                ========================================= */

                @media (max-width: 700px) {

                    .app-header {
                        flex-direction: column;
                        gap: 0.6rem;
                    }

                    .app-logo {
                        width: 96px;
                    }

                    .app-title {
                        font-size: 1.65rem;
                    }

                    .quote-title {
                        font-size: 1.8rem;
                    }

                    .availability-table {
                        font-size: 0.8rem;
                    }

                    .availability-table th,
                    .availability-table td {
                        padding: 0.34rem;
                    }

                    .stButton button,
                    .stDownloadButton button,
                    .stFormSubmitButton button {

                        width: 100%;
                    }
                }

            </style>
            """,
            unsafe_allow_html=True,
        )


# ============================================================================
# Aplicacao Streamlit
# ============================================================================

class TrebeschiCommercialApp:
    """Coordena navegacao, tela de disponibilidade e tela de cotacao."""

    def __init__(self) -> None:
        self.catalog = AppCatalog()
        self.repository = ExcelAvailabilityRepository(
            excel_path=EXCEL_PATH,
            culture_names=self.catalog.culture_order,
        )

    def run(self) -> None:
        PageStyle.apply()
        self.disable_browser_translation()
        self.render_header()
        selected_page = self.render_feature_selector()

        if selected_page == "Disponibilidade":
            self.render_availability_page()
        else:
            self.render_quote_page()

        self.render_footer()

    @staticmethod
    def disable_browser_translation() -> None:
        """Marca a pagina como pt-BR para evitar traducoes automaticas incorretas."""
        components.html(
            """
            <script>
                const doc = window.parent.document;
                doc.documentElement.setAttribute("lang", "pt-BR");
                doc.documentElement.setAttribute("translate", "no");
                doc.body.setAttribute("translate", "no");
            </script>
            """,
            height=0,
        )

    def render_header(self) -> None:
        logo_html = ""
        if LOGO_PATH.exists():
            logo_data = base64.b64encode(LOGO_PATH.read_bytes()).decode("ascii")
            logo_html = (
                f'<img class="app-logo" src="data:image/png;base64,{logo_data}" '
                'alt="Trebeschi">'
            )

        st.markdown(
            f"""
            <div class="app-header">
                {logo_html}
                <div class="app-title">{APP_TITLE}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    @staticmethod
    def render_feature_selector() -> str:
        st.markdown('<h4 class="feature-title">Escolha a funcionalidade</h4>', unsafe_allow_html=True)
        selected_page = st.radio(
            "Funcionalidade",
            ["Disponibilidade", "Cota\u00e7\u00e3o"],
            horizontal=True,
            label_visibility="collapsed",
        )
        st.divider()
        return selected_page

    def render_availability_page(self) -> None:
        st.caption(
            "Consulta de produtos disponiveis por local de carregamento."
        )
        self.render_update_status()
        self.enable_auto_refresh()
        self.render_consultation(self.repository.list_items())

    def render_quote_page(self) -> None:
        self.render_quote_header()
        self.render_quote_simulator()

    def render_update_status(self) -> None:
        if not EXCEL_PATH.exists():
            st.warning(f"Planilha disponibilidade.xlsx nao encontrada em: {EXCEL_PATH}")
            return

        modified_at = datetime.fromtimestamp(EXCEL_PATH.stat().st_mtime)
        st.caption(
            f"Ultima atualizacao da planilha: {modified_at.strftime('%d/%m/%Y %H:%M:%S')} "
            "| Atualizacao automatica a cada 2 minutos"
        )
        st.caption(f"Arquivo lido: {EXCEL_PATH}")

    @staticmethod
    def enable_auto_refresh() -> None:
        components.html(
            f"""
            <script>
                window.parent.document.documentElement.setAttribute("lang", "pt-BR");
                window.parent.document.body.setAttribute("translate", "no");
                setTimeout(function() {{
                    window.parent.location.reload();
                }}, {AUTO_REFRESH_SECONDS * 1000});
            </script>
            """,
            height=0,
        )

    def render_consultation(self, items: list[dict]) -> None:
        st.subheader("Disponibilidade por local")

        if not items:
            st.info("Nenhum produto encontrado na planilha.")
            return

        selected_location = st.selectbox(
            "Local de carregamento",
            self.catalog.location_names(),
            index=0,
        )
        available_cultures = self.catalog.cultures_for_location(selected_location, items)
        selected_culture_label = st.selectbox(
            "Cultura",
            ["Todas"] + [culture_label(culture) for culture in available_cultures],
        )
        selected_culture = culture_from_label(selected_culture_label)
        search = st.text_input(
            "Buscar produto ou packing",
            placeholder="Ex: AA Misto, Classe 5, LV",
        )

        filtered = self.filter_items(items, selected_location, selected_culture, search)
        filtered = [item for item in filtered if item["quantidade"] > 0]

        col1, col2 = st.columns(2)
        col1.metric("Itens", len(filtered))
        col2.metric("Qtd. disponivel", format_quantity(sum_quantities(filtered)) or "0")

        if not filtered:
            st.warning("Nenhum item encontrado com os filtros atuais.")
            return

        cultures_to_render = (
            [selected_culture] if selected_culture != "Todas" else available_cultures
        )
        for culture in cultures_to_render:
            culture_items = [item for item in filtered if item["cultura"] == culture]
            if culture_items:
                self.render_culture_table(culture, culture_items)

    def filter_items(
        self,
        items: list[dict],
        location: str,
        culture: str,
        search: str,
    ) -> list[dict]:
        filtered = list(items)

        if location != "Brasil":
            allowed_locations = self.catalog.excel_locations_for_filter(location)
            filtered = [
                item
                for item in filtered
                if item["local_carregamento"] in allowed_locations
            ]

        if culture != "Todas":
            filtered = [item for item in filtered if item["cultura"] == culture]

        if search:
            term = normalize_text(search)
            filtered = [
                item
                for item in filtered
                if term in normalize_text(item["produto"])
                or term in normalize_text(item["packing"])
                or term in normalize_text(item["local_carregamento"])
            ]

        return filtered

    def render_culture_table(self, culture: str, items: list[dict]) -> None:
        st.markdown(
            f'<div class="culture-title">{culture_label(culture)}</div>',
            unsafe_allow_html=True,
        )

        split_packings = unique_by_order(items, "packing", "ordem_packing")
        if len(split_packings) > 1:
            columns = st.columns(len(split_packings))
            for column, packing in zip(columns, split_packings):
                packing_items = [item for item in items if item["packing"] == packing]
                with column:
                    self.render_quantity_table(packing_title(culture, packing), packing_items)
            return

        self.render_quantity_table("", items)

    @staticmethod
    def render_quantity_table(title: str, items: list[dict]) -> None:
        if not items:
            return

        if title:
            st.markdown(f'<div class="table-subtitle">{title}</div>', unsafe_allow_html=True)

        columns = unique_quantity_columns(items)
        products = unique_by_order(items, "produto", "ordem_produto")
        rows = build_availability_rows(items, products, columns)
        st.markdown(render_html_table(rows), unsafe_allow_html=True)

    def render_quote_header(self) -> None:
        logo_html = ""
        if LOGO_PATH.exists():
            logo_data = base64.b64encode(LOGO_PATH.read_bytes()).decode("ascii")
            logo_html = (
                f'<img class="quote-logo" src="data:image/png;base64,{logo_data}" '
                'alt="Trebeschi">'
            )

        st.markdown(
            f"""
            <div class="quote-hero">
                {logo_html}
                <div class="quote-title">Cotação Trebeschi 🍅 🧄 🧅 🍠</div>
                <div class="quote-subtitle">
                    Selecione as culturas e insira os dados para gerar a cotação.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.divider()

    def render_quote_simulator(self) -> None:
        st.markdown(
            '<div class="quote-section-title">🧺 Selecione as culturas para cotação:</div>',
            unsafe_allow_html=True,
        )
        col_tomate, col_alho, col_cebola, col_batata_doce = st.columns(4)
        selected_cultures = {
            "Tomate": col_tomate.checkbox("🍅 Tomate", key="quote_tomate"),
            "Alho": col_alho.checkbox("🧄 Alho", key="quote_alho"),
            "Cebola": col_cebola.checkbox("🧅 Cebola", key="quote_cebola"),
            "Batata Doce": col_batata_doce.checkbox("🍠 Batata Doce", key="quote_batata_doce"),
        }

        if not any(selected_cultures.values()):
            st.info("☝️ Selecione pelo menos uma cultura para continuar.")
            return

        st.markdown(
            '<div class="quote-section-title">📋 Dados do cliente</div>',
            unsafe_allow_html=True,
        )
        quote_inputs = self.render_quote_form(selected_cultures)
        if quote_inputs is None:
            return

        quotes = calculate_quotes(**quote_inputs)
        validity = calculate_validity(quote_inputs["validity_input"])
        self.render_quote_results(quotes, validity)

    @staticmethod
    def render_quote_form(selected_cultures: dict[str, bool]) -> dict | None:
        with st.form("quote_form"):
            costs = {}
            if selected_cultures["Tomate"]:
                costs["Tomate"] = st.number_input("Custo Tomate (R$)", min_value=0.0, step=0.01)
            if selected_cultures["Alho"]:
                costs["Alho"] = st.number_input("Custo Alho (R$)", min_value=0.0, step=0.01)
            if selected_cultures["Cebola"]:
                costs["Cebola"] = st.number_input("Custo Cebola (R$)", min_value=0.0, step=0.01)
            if selected_cultures["Batata Doce"]:
                costs["Batata Doce"] = st.number_input("Custo Batata Doce (R$)", min_value=0.0, step=0.01)

            discount = st.number_input(
                "Desconto contratual / comissão (%)", min_value=0.0, step=0.01
            )
            logistics = st.number_input("Operação logística (R$)", min_value=0.0, step=0.01)
            tax_name = st.selectbox("Tipo de imposto", list(TAXES.keys()))

            tomato_package = None
            tomato_weight = None
            if selected_cultures["Tomate"]:
                st.markdown("#### Configuração Tomate")
                tomato_package = st.selectbox(
                    "Tipo de embalagem - Tomate",
                    ["Papelão", "HB", "IFCO", "Plástica"],
                )
                tomato_weight = st.selectbox(
                    "Peso por caixa - Tomate (kg)",
                    [18, 19, 20, 22, 23],
                )
                st.caption("Custo de mão de obra fixo: R$ 7,00")

            potato_package = None
            potato_weight = None
            if selected_cultures["Batata Doce"]:
                st.markdown("#### Configuração Batata Doce")
                potato_package = st.selectbox(
                    "Tipo de embalagem - Batata Doce",
                    ["Papelão", "IFCO", "HB", "SC"],
                )
                if potato_package == "SC":
                    potato_weight = 20
                    st.caption("Peso padrão para SC: 20 kg")
                else:
                    potato_weight = st.selectbox(
                        "Peso por caixa - Batata Doce (kg)",
                        [18, 20, 22,],
                    )

            validity_input = st.text_input(
                "Validade da cotação (hora e minuto, ex: 1430)",
                placeholder="ex: 1430",
            )
            submitted = st.form_submit_button("Calcular cotação", type="primary")

        if not submitted:
            st.info("Preencha os dados e clique em Calcular cotação.")
            return None

        return {
            "selected_cultures": selected_cultures,
            "costs": costs,
            "discount": discount,
            "logistics": logistics,
            "tax": TAXES[tax_name],
            "tomato_package": tomato_package,
            "tomato_weight": tomato_weight,
            "potato_package": potato_package,
            "potato_weight": potato_weight,
            "validity_input": validity_input,
        }

    def render_quote_results(self, quotes: list[dict], validity: datetime) -> None:
        data_str = validity.strftime("%d/%m/%Y")
        hora_str = validity.strftime("%H:%M")

        st.markdown(
            f"""
            <div class="quote-result">
                <h4 style="color: var(--danger); margin: 0.2rem 0;">COTAÇÃO TREBESCHI</h4>
                <strong>Válido até as {hora_str} hs do dia {data_str}.</strong><br>
                <span>Disponibilidade sujeita a alteração.</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        for quote in quotes:
            st.markdown(
                f"""
                <div class="quote-result">
                    <h3>{quote["icone"]} {quote["cultura"]}</h3>
                    <div>Embalagem: {quote["embalagem"]}</div>
                    <div class="quote-price">R$ {quote["valor_cx"]:.2f}</div>
                    <strong style="color: var(--green-bright);">
                        Valor por kg: R$ {quote["valor_kg"]:.2f}
                    </strong>
                </div>
                """,
                unsafe_allow_html=True,
            )

        try:
            pdf_buffer = generate_quote_pdf(quotes, LOGO_PATH, validity)
            st.download_button(
                label="Baixar PDF da cotação",
                data=pdf_buffer,
                file_name=f"Cotacao_Trebeschi_{data_str.replace('/', '-')}.pdf",
                mime="application/pdf",
            )
        except RuntimeError as error:
            st.error(str(error))

    @staticmethod
    def render_footer() -> None:
        st.markdown(
            """
            <div class="app-footer">
                Desenvolvido por Rafael Germano.<br>
                Para uso interno da Trebeschi.
            </div>
            """,
            unsafe_allow_html=True,
        )


# ============================================================================
# Calculos da cotacao
# ============================================================================

def calculate_quotes(
    selected_cultures: dict[str, bool],
    costs: dict[str, float],
    discount: float,
    logistics: float,
    tax: float,
    tomato_package: str | None,
    tomato_weight: int | None,
    potato_package: str | None,
    potato_weight: int | None,
    validity_input: str = "",
) -> list[dict]:
    """Calcula as cotacoes por cultura usando as mesmas premissas do app original."""

    del validity_input
    packing_labor = 7.00
    quotes = []

    if selected_cultures["Tomate"]:
        weight = tomato_weight or 20
        package = tomato_package or "Papelão"
        total_cost = ((costs["Tomate"] / 20) * weight) + logistics + packing_labor + PACKAGE_COSTS[package]
        box_value, kg_value = calculate_price(total_cost, weight, discount, tax)
        quotes.append(build_quote("Tomate", "🍅", f"{package} {weight} kg", box_value, kg_value))

    if selected_cultures["Alho"]:
        weight = 10
        total_cost = costs["Alho"] + (logistics * 0.6)
        box_value, kg_value = calculate_price(total_cost, weight, discount, tax)
        quotes.append(build_quote("Alho", "🧄", "Caixa 10 kg", box_value, kg_value))

    if selected_cultures["Cebola"]:
        weight = 20
        total_cost = costs["Cebola"] + logistics
        box_value, kg_value = calculate_price(total_cost, weight, discount, tax)
        quotes.append(build_quote("Cebola", "🧅", "Saco 20 kg", box_value, kg_value))

    if selected_cultures["Batata Doce"]:
        weight = potato_weight or 20
        package = potato_package or "Papelão"
        total_cost = ((costs["Batata Doce"] / 20) * weight) + logistics + PACKAGE_COSTS[package]
        box_value, kg_value = calculate_price(total_cost, weight, discount, tax)
        quotes.append(build_quote("Batata Doce", "🍠", f"{package} {weight} kg", box_value, kg_value))

    return quotes


def build_quote(culture: str, icon: str, package: str, box_value: float, kg_value: float) -> dict:
    return {
        "cultura": culture,
        "icone": icon,
        "embalagem": package,
        "valor_cx": box_value,
        "valor_kg": kg_value,
    }


def calculate_price(total_cost: float, weight: float, discount: float, tax: float) -> tuple[float, float]:
    box_value = total_cost / (1 - (discount / 100 + tax))
    return round_up_10_cents(box_value), box_value / weight


def round_up_10_cents(value: float) -> float:
    return ceil(value * 10) / 10


def calculate_validity(value: str) -> datetime:
    now = datetime.now(BRASILIA_TZ)
    if value and value.isdigit() and len(value) in [3, 4]:
        padded = value.zfill(4)
        try:
            hour = int(padded[:2])
            minute = int(padded[2:])
            validity = BRASILIA_TZ.localize(datetime.combine(date.today(), time(hour, minute)))
            if validity < now:
                validity += timedelta(days=1)
            return validity
        except ValueError:
            pass
    return now + timedelta(hours=2)


def generate_quote_pdf(quotes: list[dict], logo_path: Path, validity: datetime) -> BytesIO:
    """Gera PDF sob demanda; reportlab e importado aqui para o app abrir mesmo sem PDF."""

    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import cm
        from reportlab.lib.utils import ImageReader
        from reportlab.pdfgen import canvas
    except ModuleNotFoundError as exc:
        raise RuntimeError(
            "A biblioteca reportlab nao esta instalada. Rode: pip install -r requirements.txt"
        ) from exc

    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    if logo_path.exists():
        try:
            logo_img = PILImage.open(logo_path).convert("RGBA")
            background = PILImage.new("RGBA", logo_img.size, (255, 255, 255, 255))
            background.paste(logo_img, mask=logo_img)
            temp_logo = BytesIO()
            background.save(temp_logo, format="PNG")
            temp_logo.seek(0)
            pdf.drawImage(
                ImageReader(temp_logo),
                (width - 6 * cm) / 2,
                height - 6 * cm,
                width=6 * cm,
                height=3 * cm,
                mask="auto",
            )
        except OSError:
            pass

    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawCentredString(width / 2, height - 7 * cm, "Cotacao Trebeschi")
    pdf.setFont("Helvetica", 12)
    pdf.drawCentredString(
        width / 2,
        height - 8 * cm,
        f"Valido ate as {validity.strftime('%H:%M')} hs do dia {validity.strftime('%d/%m/%Y')}.",
    )
    pdf.drawCentredString(width / 2, height - 9 * cm, "Disponibilidade sujeita a alteracao.")

    y = height - 10 * cm
    for quote in quotes:
        pdf.setFont("Helvetica-Bold", 12)
        pdf.drawString(3 * cm, y, quote["cultura"])
        y -= 0.6 * cm
        pdf.setFont("Helvetica", 11)
        pdf.drawString(3 * cm, y, f"Embalagem: {quote['embalagem']}")
        y -= 0.4 * cm
        pdf.drawString(3 * cm, y, f"Valor por embalagem: R$ {quote['valor_cx']:.2f}")
        y -= 0.4 * cm
        pdf.drawString(3 * cm, y, f"Valor por kg: R$ {quote['valor_kg']:.2f}")
        y -= 1.2 * cm

        if y < 5 * cm:
            pdf.showPage()
            y = height - 4 * cm

    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawCentredString(
        width / 2,
        3 * cm,
        "Gerado automaticamente pelo App Comercial Trebeschi",
    )
    pdf.save()
    buffer.seek(0)
    return buffer


# ============================================================================
# Utilitarios de tabela e texto
# ============================================================================

def build_availability_rows(items: list[dict], products: list[str], columns: list[str]) -> list[dict]:
    rows = []
    for product in products:
        row = {"Produto": product}
        row_total = 0
        for column in columns:
            quantity = sum(
                item["quantidade"]
                for item in items
                if item["produto"] == product
                and quantity_column_name(item, items) == column
            )
            row[column] = quantity
            row_total += quantity
        if row_total > 0:
            row["Total"] = row_total
            rows.append(row)

    total_row = {"Produto": "TOTAL"}
    for column in columns:
        total_row[column] = sum(
            item["quantidade"]
            for item in items
            if quantity_column_name(item, items) == column
        )
    total_row["Total"] = sum_quantities(items)
    rows.append(total_row)
    return rows


def cell_value(matrix: list[list], row: int, col: int):
    if row >= len(matrix) or col >= len(matrix[row]):
        return None
    return matrix[row][col]


def cell_text(matrix: list[list], row: int, col: int) -> str:
    value = cell_value(matrix, row, col)
    return "" if value is None else str(value).strip()


def parse_number(value) -> float:
    if value is None or value == "":
        return 0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(".", "").replace(",", "."))
    except ValueError:
        return 0


def format_date(value) -> str:
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    if value:
        return str(value)
    return datetime.now().strftime("%d/%m/%Y")


def normalize_text(value) -> str:
    text = "" if value is None else str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    return "".join(char for char in text if not unicodedata.combining(char))


def culture_label(culture: str) -> str:
    if culture == "Todas":
        return culture
    return f"{CULTURE_ICONS.get(culture, '•')} {culture}"


def culture_from_label(label: str) -> str:
    if label == "Todas":
        return "Todas"
    return label.split(" ", 1)[1] if " " in label else label


def packing_title(culture: str, packing: str) -> str:
    tomato_titles = {
        "lv": "Longa Vida (LV)",
        "it": "Italiano (IT)",
    }
    if culture == "Tomate":
        return tomato_titles.get(normalize_text(packing), packing)
    return packing


def quantity_column_name(item: dict, items: list[dict]) -> str:
    locations = {row["local_carregamento"] for row in items}
    if len(locations) == 1:
        return item["packing"]
    return f"{item['local_carregamento']} - {item['packing']}"


def format_quantity(value: float) -> str:
    if not value:
        return ""
    return f"{value:,.0f}".replace(",", ".")


def render_html_table(rows: list[dict]) -> str:
    if not rows:
        return ""

    headers = list(rows[0].keys())
    header_html = "".join(
        f'<th translate="no">{escape_html(header)}</th>' for header in headers
    )
    body_html = []

    for row in rows:
        row_class = ' class="total-row"' if row.get("Produto") == "TOTAL" else ""
        cells = []
        for header in headers:
            value = row.get(header, "")
            if header != "Produto":
                value = format_quantity(value)
            cells.append(f'<td translate="no">{escape_html(value)}</td>')
        body_html.append(f"<tr{row_class}>{''.join(cells)}</tr>")

    return (
        '<table class="availability-table" translate="no">'
        f"<thead><tr>{header_html}</tr></thead>"
        f"<tbody>{''.join(body_html)}</tbody>"
        "</table>"
    )


def escape_html(value) -> str:
    text = "" if value is None else str(value)
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&#39;")
    )


def sum_quantities(items: list[dict]) -> float:
    return sum(item["quantidade"] for item in items)


def unique_by_order(items: list[dict], value_key: str, order_key: str) -> list:
    ordered_values = {}
    for item in items:
        value = item[value_key]
        order = item.get(order_key, 0)
        if value not in ordered_values or order < ordered_values[value]:
            ordered_values[value] = order
    return [
        value
        for value, _ in sorted(
            ordered_values.items(),
            key=lambda pair: (pair[1], str(pair[0])),
        )
    ]


def unique_quantity_columns(items: list[dict]) -> list:
    ordered_columns = {}
    for item in items:
        column = quantity_column_name(item, items)
        order = item.get("ordem_packing", 0)
        if column not in ordered_columns or order < ordered_columns[column]:
            ordered_columns[column] = order
    return [
        column
        for column, _ in sorted(
            ordered_columns.items(),
            key=lambda pair: (pair[1], str(pair[0])),
        )
    ]


if __name__ == "__main__":
    TrebeschiCommercialApp().run()
