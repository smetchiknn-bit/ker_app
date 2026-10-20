"""
Главный интерфейс приложения для работы с базой КЕР.
Запуск:  streamlit run app.py
"""
import streamlit as st
import pandas as pd
from pathlib import Path
import io

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
if "vor" not in st.session_state:
    # ВОР — список словарей. Пустая строка = {"empty": True}
    st.session_state.vor = []
if "display_mode" not in st.session_state:
    # "hide" / "mark" / "show"
    st.session_state.display_mode = "hide"


# ==================== SIDEBAR ====================
with st.sidebar:
    st.title("⚙️ Настройки")

    st.subheader("📂 Источник данных")
    uploaded_file = st.file_uploader(
        "Загрузите Excel-файл",
        type=["xlsx", "xls", "xlsm"],
        help="Файл должен содержать лист 'Выгрузка' или 'КЕР в ССР'"
    )

    if uploaded_file is not None:
        # Сохраняем во временный файл
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

    # Проверка дублей
    if st.session_state.df is not None:
        dupes = find_duplicate_ids(st.session_state.df)
        if not dupes.empty:
            st.divider()
            st.subheader("⚠️ Дубликаты ИД_КЕР")
            st.warning(f"Найдено {len(dupes)} ИД_КЕР с несколькими вариантами:")
            st.dataframe(dupes, use_container_width=True, hide_index=True)

    # Экспорт обработанной таблицы
    if st.session_state.df is not None:
        st.divider()
        st.subheader("💾 Экспорт")
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


# ==================== ОСНОВНАЯ ЧАСТЬ ====================
st.title("🏗 Поиск расценок КЕР в ССР")

if st.session_state.df is None:
    st.info("👈 Загрузите Excel-файл в панели слева, чтобы начать работу.")
    st.stop()

df = st.session_state.df
display_mode = st.session_state.display_mode


# ---- Фильтрация служебных расценок ----
def is_service_row(name: str) -> bool:
    """Определяет, является ли расценка служебной."""
    if pd.isna(name):
        return False
    name_upper = str(name).upper()
    return ("НЕ ИСПОЛЬЗОВАТЬ" in name_upper) or name_upper.startswith("XXX")


df_view = df.copy()
if display_mode == "hide":
    df_view = df_view[~df_view["Наименование_КЕР"].apply(is_service_row)].copy()
# "mark" и "show" — показываем всё, но в "mark" добавим визуальную пометку ниже


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
        "🔍 Ключевое слово в наименовании:",
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

# Формируем список для selectbox: "ИД_КЕР | Наименование_КЕР | ЕдИзм"
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
    st.markdown("### 📋 Основная информация")
    st.markdown(f"**ИД КЕР:** `{selected_row['ИД_КЕР']}`")
    st.markdown(f"**Ед. измерения:** `{selected_row['ЕдИзм КЕР']}`")
    if pd.notna(selected_row.get("ФЕР")) and str(selected_row.get("ФЕР")).strip():
        st.markdown(f"**ФЕР:** `{selected_row['ФЕР']}`")

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


# ==================== КНОПКИ ВОР ====================
st.divider()
st.subheader("📊 Ведомость объёмов работ (ВОР)")

btn_col1, btn_col2, btn_col3 = st.columns([1, 1, 2])

with btn_col1:
    if st.button("➕ Вставить в ВОР", type="primary", use_container_width=True):
        st.session_state.vor.append({
            "ИД_КЕР": selected_row["ИД_КЕР"],
            "Наименование_КЕР": selected_row["Наименование_КЕР"],
            "ЕдИзм КЕР": selected_row["ЕдИзм КЕР"],
            "empty": False
        })
        st.rerun()

with btn_col2:
    if st.button("📝 Пустая строка (комментарий)", use_container_width=True):
        st.session_state.vor.append({"empty": True})
        st.rerun()


# ==================== ТАБЛИЦА ВОР ====================
if not st.session_state.vor:
    st.info("ВОР пуста. Добавьте расценки кнопкой выше.")
else:
    st.markdown(f"**В ВОР добавлено позиций:** {len(st.session_state.vor)}")

    # Рендерим таблицу построчно с кнопками удаления
    vor_rows = st.session_state.vor
    rows_to_delete = []

    for i, row in enumerate(vor_rows):
        col_num, col_content, col_action = st.columns([0.5, 8, 0.5])
        with col_num:
            st.markdown(f"**{i + 1}.**")
        with col_content:
            if row.get("empty"):
                st.markdown("*— пустая строка для комментария —*")
            else:
                st.markdown(
                    f"**{row['ИД_КЕР']}** | {row['Наименование_КЕР']} | *{row['ЕдИзм КЕР']}*"
                )
        with col_action:
            if st.button("❌", key=f"del_{i}", help="Удалить строку"):
                rows_to_delete.append(i)

    # Удаляем помеченные строки (с конца, чтобы индексы не сдвинулись)
    if rows_to_delete:
        for idx in sorted(rows_to_delete, reverse=True):
            st.session_state.vor.pop(idx)
        st.rerun()

    # ==================== ЭКСПОРТ ВОР ====================
    st.divider()

    # Формируем DataFrame для экспорта
    vor_data = []
    for row in st.session_state.vor:
        if row.get("empty"):
            vor_data.append({"ИД_КЕР": "", "Наименование_КЕР": "", "ЕдИзм КЕР": ""})
        else:
            vor_data.append({
                "ИД_КЕР": row["ИД_КЕР"],
                "Наименование_КЕР": row["Наименование_КЕР"],
                "ЕдИзм КЕР": row["ЕдИзм КЕР"]
            })

    df_vor = pd.DataFrame(vor_data)

    # Экспорт в Excel
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df_vor.to_excel(writer, sheet_name="ВОР", index=False)
        workbook = writer.book
        worksheet = writer.sheets["ВОР"]
        worksheet.set_column("A:A", 12)
        worksheet.set_column("B:B", 60)
        worksheet.set_column("C:C", 15)

    st.download_button(
        "📥 Скачать ВОР.xlsx",
        data=buffer.getvalue(),
        file_name="ВОР.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Кнопка очистки
    if st.button("🗑 Очистить всю ВОР"):
        st.session_state.vor = []
        st.rerun()
