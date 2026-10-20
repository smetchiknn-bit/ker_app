"""
Главный интерфейс приложения для работы с базой КЕР.
Запуск:  streamlit run app.py
"""
import streamlit as st
import pandas as pd
from pathlib import Path
import io
import json
from datetime import datetime

from data_loader import load_or_process, export_processed, find_duplicate_ids


# ==================== НАСТРОЙКИ СТРАНИЦЫ ====================
st.set_page_config(
    page_title="Поиск КЕР в ССР",
    page_icon="🏗",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ==================== ИНИЦИАЛИЗАЦИЯ SESSION STATE ====================
if "df" not in st.session_state:
    st.session_state.df = None
if "source_sheet" not in st.session_state:
    st.session_state.source_sheet = None
if "vor_df" not in st.session_state:
    st.session_state.vor_df = pd.DataFrame(columns=[
        "ИД_КЕР", "Наименование_КЕР", "ЕдИзм КЕР", "Кол-во", "Тип", "Комментарий"
    ])
if "display_mode" not in st.session_state:
    st.session_state.display_mode = "hide"


# ==================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ====================
def is_service_row(name: str) -> bool:
    """Определяет, является ли расценка служебной."""
    if pd.isna(name):
        return False
    name_upper = str(name).upper()
    return ("НЕ ИСПОЛЬЗОВАТЬ" in name_upper) or name_upper.startswith("XXX")


def vor_to_json() -> str:
    """Экспортирует текущую ВОР в JSON-строку."""
    return st.session_state.vor_df.to_json(orient="records", force_ascii=False)


def json_to_vor(json_str: str) -> pd.DataFrame:
    """Импортирует ВОР из JSON-строки."""
    data = json.loads(json_str)
    if not data:
        return pd.DataFrame(columns=[
            "ИД_КЕР", "Наименование_КЕР", "ЕдИзм КЕР", "Кол-во", "Тип", "Комментарий"
        ])
    return pd.DataFrame(data)


# ==================== SIDEBAR ====================
with st.sidebar:
    st.title("⚙️ Настройки")

    # --- Источник данных ---
    st.subheader("📂 Источник данных")
    uploaded_file = st.file_uploader(
        "Загрузите Excel-файл",
        type=["xlsx", "xls", "xlsm"],
        help="Файл должен содержать лист 'Выгрузка' или 'КЕР в ССР'"
    )

    if uploaded_file is not None:
        tmp_path = Path("tmp_upload.xlsx")
        tmp_path.write_bytes(uploaded_file.read())

        try:
            with st.spinner("Загрузка данных..."):
                df, source_sheet = load_or_process(str(tmp_path))
            st.session_state.df = df
            st.session_state.source_sheet = source_sheet
            st.success(f"✅ Загружено: {len(df)} строк\n(лист: {source_sheet})")
        except Exception as e:
            st.error(f"Ошибка загрузки: {e}")

    st.divider()

    # --- Отображение служебных расценок ---
    st.subheader("👁 Служебные расценки")
    st.caption("Расценки с пометкой «НЕ ИСПОЛЬЗОВАТЬ В СР» или «XXX»")
    display_mode = st.radio(
        "Выберите вариант отображения:",
        options=["hide", "mark", "show"],
        format_func=lambda x: {
            "hide": "🚫 Скрывать полностью",
            "mark": "⚠️ Показывать с пометкой",
            "show": "✅ Показывать как есть"
        }[x],
        index=["hide", "mark", "show"].index(st.session_state.display_mode),
        key="display_mode_radio"
    )
    st.session_state.display_mode = display_mode

    # --- Проверка дубликатов ---
    if st.session_state.df is not None:
        dupes = find_duplicate_ids(st.session_state.df)
        if not dupes.empty:
            st.divider()
            st.subheader("⚠️ Дубликаты ИД_КЕР")
            st.warning(f"Найдено {len(dupes)} ИД_КЕР с несколькими вариантами:")
            st.dataframe(dupes, use_container_width=True, hide_index=True)

    # --- Экспорт обработанной таблицы ---
    if st.session_state.df is not None:
        st.divider()
        st.subheader("💾 Экспорт базы")
        buffer = io.BytesIO()
        export_processed(st.session_state.df, "tmp_export.xlsx")
        with open("tmp_export.xlsx", "rb") as f:
            buffer.write(f.read())
        st.download_button(
            "📥 Скачать обработанную таблицу",
            data=buffer.getvalue(),
            file_name="Выгрузка_обработанная.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # --- Сохранение/загрузка ВОР ---
    st.divider()
    st.subheader("💾 ВОР — сохранение")
    
    if not st.session_state.vor_df.empty:
        json_data = vor_to_json()
        st.download_button(
            "📤 Сохранить ВОР (.json)",
            data=json_data,
            file_name=f"ВОР_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
            mime="application/json"
        )

    uploaded_vor = st.file_uploader(
        "📥 Загрузить ВОР из файла",
        type=["json"],
        help="Загрузите ранее сохранённый JSON-файл ВОР"
    )
    if uploaded_vor is not None:
        try:
            json_str = uploaded_vor.read().decode("utf-8")
            st.session_state.vor_df = json_to_vor(json_str)
            st.success(f"✅ ВОР загружена: {len(st.session_state.vor_df)} строк")
            st.rerun()
        except Exception as e:
            st.error(f"Ошибка загрузки ВОР: {e}")


# ==================== ОСНОВНАЯ ЧАСТЬ ====================
st.title("🏗 Поиск расценок КЕР в ССР")

if st.session_state.df is None:
    st.info("👈 Загрузите Excel-файл в панели слева, чтобы начать работу.")
    st.stop()

df = st.session_state.df
display_mode = st.session_state.display_mode


# ---- Фильтрация служебных расценок ----
df_view = df.copy()
if display_mode == "hide":
    df_view = df_view[~df_view["Наименование_КЕР"].apply(is_service_row)].copy()


# ==================== ВВОДНАЯ ПАНЕЛЬ ====================
col_l3, col_search = st.columns([1, 2])

with col_l3:
    l3_options = ["📋 Весь список"] + sorted(
        df_view["Л3 Наименование"].dropna().unique().tolist()
    )
    selected_l3 = st.selectbox(
        "🔹 Уровень Л3 (раздел работ):",
        options=l3_options,
        index=0
    )

# Фильтруем по Л3
df_filtered = df_view.copy()
if selected_l3 != "📋 Весь список":
    df_filtered = df_filtered[df_filtered["Л3 Наименование"] == selected_l3]

with col_search:
    keyword = st.text_input(
        " Ключевое слово в наименовании:",
        placeholder="например: катками",
        help="Введите слово — список ниже отфильтруется"
    )

if keyword:
    df_filtered = df_filtered[
        df_filtered["Наименование_КЕР"].str.contains(keyword, case=False, na=False)
    ]


# ==================== ВЫБОР РАСЦЕНКИ ====================
if df_filtered.empty:
    st.warning("По вашему запросу ничего не найдено. Попробуйте изменить ключевое слово или уровень Л3.")
    st.stop()

df_filtered = df_filtered.reset_index(drop=True)
options_labels = [
    f"{row['ИД_КЕР']} | {row['Наименование_КЕР']} | {row['ЕдИзм КЕР']}"
    for _, row in df_filtered.iterrows()
]

selected_label = st.selectbox(
    "📌 Выберите расценку (можно использовать стрелки ↓ ↑ на клавиатуре):",
    options=options_labels,
    index=0
)

selected_idx = options_labels.index(selected_label)
selected_row = df_filtered.iloc[selected_idx]


# ==================== КАРТОЧКА РАСЦЕНКИ ====================
st.divider()

# Визуальная пометка для служебных
if display_mode == "mark" and is_service_row(selected_row["Наименование_КЕР"]):
    st.warning("⚠️ Это служебная расценка (пометка «НЕ ИСПОЛЬЗОВАТЬ В СР» или «XXX»). Используйте с осторожностью.")

col_card1, col_card2 = st.columns([1, 2])

with col_card1:
    st.markdown("###  Основная информация")
    st.markdown(f"**ИД КЕР:** `{selected_row['ИД_КЕР']}`")
    st.markdown(f"**Ед. измерения:** `{selected_row['ЕдИзм КЕР']}`")
    fer_val = selected_row.get("ФЕР")
    if pd.notna(fer_val) and str(fer_val).strip():
        st.markdown(f"**ФЕР:** `{fer_val}`")

with col_card2:
    st.markdown("### 🗂 Иерархия")
    st.text(selected_row["Иерархия"])

st.markdown("### 📝 Лист состава работ КЕР")
work_content = selected_row.get("Лист состава работ КЕР", "")
if pd.notna(work_content) and str(work_content).strip():
    st.text_area(
        "",
        value=str(work_content),
        height=200,
        disabled=True,
        label_visibility="collapsed"
    )
else:
    st.info("Состав работ не указан.")


# ==================== КНОПКИ ДОБАВЛЕНИЯ В ВОР ====================
st.divider()
st.subheader("📊 Ведомость объёмов работ (ВОР)")

btn_col1, btn_col2 = st.columns([1, 1])

with btn_col1:
    if st.button("➕ Добавить расценку в ВОР", type="primary", use_container_width=True):
        new_row = {
            "ИД_КЕР": selected_row["ИД_КЕР"],
            "Наименование_КЕР": selected_row["Наименование_КЕР"],
            "ЕдИзм КЕР": selected_row["ЕдИзм КЕР"],
            "Кол-во": 1.0,
            "Тип": "Расценка",
            "Комментарий": ""
        }
        st.session_state.vor_df = pd.concat(
            [st.session_state.vor_df, pd.DataFrame([new_row])],
            ignore_index=True
        )
        st.rerun()

with btn_col2:
    if st.button(" Добавить строку-комментарий", use_container_width=True):
        new_row = {
            "ИД_КЕР": "",
            "Наименование_КЕР": "",
            "ЕдИзм КЕР": "",
            "Кол-во": "",
            "Тип": "Комментарий",
            "Комментарий": "Введите текст заголовка или примечания..."
        }
        st.session_state.vor_df = pd.concat(
            [st.session_state.vor_df, pd.DataFrame([new_row])],
            ignore_index=True
        )
        st.rerun()


# ==================== ТАБЛИЦА ВОР (РЕДАКТИРУЕМАЯ) ====================
if st.session_state.vor_df.empty:
    st.info("ВОР пуста. Добавьте расценки или комментарии кнопками выше.")
else:
    st.markdown(f"**В ВОР добавлено позиций:** {len(st.session_state.vor_df)}")

    column_config = {
        "ИД_КЕР": st.column_config.TextColumn("ИД КЕР", width="small", disabled=True),
        "Наименование_КЕР": st.column_config.TextColumn("Наименование работ", width="large", disabled=True),
        "ЕдИзм КЕР": st.column_config.TextColumn("Ед. изм.", width="small", disabled=True),
        "Кол-во": st.column_config.NumberColumn("Кол-во", width="small", min_value=0.0, step=0.1, format="%.2f"),
        "Тип": st.column_config.SelectboxColumn("Тип", options=["Расценка", "Комментарий"], width="small"),
        "Комментарий": st.column_config.TextColumn("Комментарий / Заголовок", width="medium")
    }

    edited_df = st.data_editor(
        st.session_state.vor_df,
        column_config=column_config,
        hide_index=True,
        use_container_width=True,
        num_rows="dynamic",
        disabled=["ИД_КЕР", "Наименование_КЕР", "ЕдИзм КЕР"]
    )

    if not edited_df.equals(st.session_state.vor_df):
        st.session_state.vor_df = edited_df

    # ==================== ЭКСПОРТ И ОЧИСТКА ВОР ====================
    st.divider()
    exp_col1, exp_col2 = st.columns([1, 1])

    with exp_col1:
        export_df = edited_df.copy()
        export_df = export_df[["ИД_КЕР", "Наименование_КЕР", "ЕдИзм КЕР", "Кол-во", "Тип", "Комментарий"]]

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            export_df.to_excel(writer, sheet_name="ВОР", index=False)
            workbook = writer.book
            worksheet = writer.sheets["ВОР"]
            worksheet.set_column("A:A", 12)
            worksheet.set_column("B:B", 60)
            worksheet.set_column("C:C", 10)
            worksheet.set_column("D:D", 10)
            worksheet.set_column("E:E", 12)
            worksheet.set_column("F:F", 40)

        st.download_button(
            "📥 Скачать ВОР.xlsx",
            data=buffer.getvalue(),
            file_name=f"ВОР_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with exp_col2:
        if st.button("🗑 Очистить всю ВОР", type="secondary", use_container_width=True):
            st.session_state.vor_df = pd.DataFrame(columns=[
                "ИД_КЕР", "Наименование_КЕР", "ЕдИзм КЕР", "Кол-во", "Тип", "Комментарий"
            ])
            st.rerun()
