"""Streamlit-интерфейс для расчётного пакета hydraulic."""

from __future__ import annotations

import io
import tempfile
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

import hydraulic.config as config
from hydraulic import profile


st.set_page_config(page_title="Hydraulic — расчёт", page_icon="🌊", layout="wide")


def read_sheet_names(file_bytes: bytes) -> list[str]:
    workbook = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    try:
        return workbook.sheetnames
    finally:
        workbook.close()


def read_sheet_preview(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None, engine="openpyxl")


def render_profile(data: pd.DataFrame) -> None:
    """Рисует координаты сечения, которые расчётное ядро читает из A:B начиная со 2 строки."""
    if data.shape[1] < 2:
        st.info("На листе не найдены две первые колонки координат профиля.")
        return

    points = data.iloc[1:, :2].copy()
    points.columns = ["Пикетаж", "Отметка"]
    points["Пикетаж"] = pd.to_numeric(points["Пикетаж"], errors="coerce")
    points["Отметка"] = pd.to_numeric(points["Отметка"], errors="coerce")
    points = points.dropna()
    if len(points) < 2:
        st.info("Не удалось определить точки профиля в колонках A и B.")
        return

    fig, ax = plt.subplots(figsize=(10, 4.2))
    ax.fill_between(points["Пикетаж"], points["Отметка"], points["Отметка"].min(), color="#b9e5f5", alpha=0.55)
    ax.plot(points["Пикетаж"], points["Отметка"], color="#126782", linewidth=2)
    ax.scatter(points["Пикетаж"], points["Отметка"], color="#126782", s=15, zorder=3)
    ax.set_xlabel("Пикетаж, м")
    ax.set_ylabel("Отметка, м")
    ax.grid(True, color="#dce5e9", linewidth=0.7)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)


st.title("🌊 Hydraulic")
st.caption("Гидравлический расчёт поперечных профилей по исходным данным Excel")

uploaded = st.file_uploader("Исходный Excel-файл", type=["xlsx", "xlsm"], help="Формат листа должен соответствовать шаблону hydraulic.")

if uploaded is None:
    st.info("Загрузите Excel-файл, чтобы просмотреть поперечный профиль и настроить расчёт.")
    st.stop()

file_bytes = uploaded.getvalue()
try:
    sheet_names = read_sheet_names(file_bytes)
except Exception as exc:
    st.error(f"Не удалось открыть Excel-файл: {exc}")
    st.stop()

if not sheet_names:
    st.error("В книге нет листов для расчёта.")
    st.stop()

left, right = st.columns([1, 2])
with left:
    sheet_name = st.selectbox("Расчётный створ", sheet_names)
    st.subheader("Параметры расчёта")
    calc_type_label = st.selectbox(
        "Тип потока",
        ["Обычная вода", "Водокаменный селевой поток", "Грязекаменный селевой поток"],
    )
    shezi_type = st.selectbox("Формула коэффициента Шези", config.SHEZI_TYPE_DEFAULTS, index=config.SHEZI_TYPE_DEFAULTS.index(config.SHEZI_TYPE))
    overflow = st.checkbox("Учитывать перелив через бровку", value=config.OVERFLOW)
    use_depth = st.checkbox("Использовать среднюю глубину вместо гидравлического радиуса", value=config.USE_H_INSTEAD_R)
    smooth_curve = st.checkbox("Сглаживать расчётные кривые", value=config.HYDRAULIC_CURVE_SMOOTHED)
    report_name = st.text_input("Имя отчёта", value=f"{Path(uploaded.name).stem}_расчёт.docx")

with right:
    try:
        preview = read_sheet_preview(file_bytes, sheet_name)
    except Exception as exc:
        st.error(f"Не удалось прочитать лист «{sheet_name}»: {exc}")
        st.stop()
    st.subheader("Поперечный профиль")
    render_profile(preview)
    with st.expander("Предварительный просмотр исходных данных"):
        st.dataframe(preview, use_container_width=True, hide_index=True)

st.divider()
if st.button("Выполнить расчёт", type="primary", use_container_width=True):
    if not report_name.lower().endswith(".docx"):
        report_name += ".docx"

    with st.spinner("Выполняется расчёт и формируется Word-отчёт…"):
        try:
            with tempfile.TemporaryDirectory(prefix="hydraulic_") as temp_dir:
                input_path = Path(temp_dir) / Path(uploaded.name).name
                output_path = Path(temp_dir) / Path(report_name).name
                input_path.write_bytes(file_bytes)

                config.CALC_TYPE = [1, 2, 3][["Обычная вода", "Водокаменный селевой поток", "Грязекаменный селевой поток"].index(calc_type_label)]
                config.SHEZI_TYPE = shezi_type
                config.OVERFLOW = overflow
                config.USE_H_INSTEAD_R = use_depth
                config.HYDRAULIC_CURVE_SMOOTHED = smooth_curve
                config.SAVE_REPORT = True
                config.PROFILE_SAVE_PICTURES = False
                config.CURVE_SAVE_PICTURES = False

                sheet_index = sheet_names.index(sheet_name)
                profile.xls_calculate_hydraulic(str(input_path), str(output_path), sheet_index)
                report_bytes = output_path.read_bytes()

            st.success("Расчёт завершён. Отчёт готов к скачиванию.")
            st.download_button(
                "Скачать Word-отчёт",
                data=report_bytes,
                file_name=Path(report_name).name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        except Exception as exc:
            st.error(f"Не удалось выполнить расчёт: {exc}")
