"""Streamlit-интерфейс для расчётного пакета hydraulic."""

from __future__ import annotations

import io
import hashlib
import tempfile
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

import hydraulic.config as config
from hydraulic import profile
from hydraulic.report import generate_morfostvor_report


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

with st.sidebar:
    st.header("Исходные данные")
    uploaded = st.file_uploader(
        "Excel-файл",
        type=["xlsx", "xlsm"],
        help="Формат листа должен соответствовать шаблону hydraulic.",
    )

if uploaded is None:
    st.info("Загрузите Excel-файл через боковую панель, чтобы просмотреть поперечный профиль и настроить расчёт.")
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

with st.sidebar:
    st.header("Расчёт")
    sheet_name = st.selectbox("Расчётный створ", sheet_names)
    calc_type_label = st.selectbox(
        "Тип потока",
        ["Вода", "Водокаменный селевой поток", "Грязекаменный селевой поток"],
    )
    shezi_type = st.selectbox(
        "Формула коэффициента Шези",
        config.SHEZI_TYPE_DEFAULTS,
        index=config.SHEZI_TYPE_DEFAULTS.index(config.SHEZI_TYPE),
    )
    overflow = st.checkbox("Учитывать перелив через бровку", value=config.OVERFLOW)
    use_depth = st.checkbox("Использовать среднюю глубину вместо гидравлического радиуса", value=config.USE_H_INSTEAD_R)
    smooth_curve = st.checkbox("Сглаживать расчётные кривые", value=config.HYDRAULIC_CURVE_SMOOTHED)
    with st.expander("Параметры сглаживания", expanded=False):
        smoothing_tolerance = st.slider(
            "Порог обнаружения скачка",
            min_value=0.0,
            max_value=0.1,
            value=float(config.HYDRAULIC_CURVE_SMOOTHING_TOLERANCE),
            step=0.001,
            help="Минимальное отклонение от локального тренда как доля диапазона кривой.",
        )
        smoothing_points = st.slider(
            "Точек отрисовки на интервал",
            min_value=1,
            max_value=20,
            value=int(config.HYDRAULIC_CURVE_SMOOTHING_POINTS_PER_INTERVAL),
        )
        smoothing_preserve = st.slider(
            "Закрепить начальные точки",
            min_value=0,
            max_value=10,
            value=int(config.HYDRAULIC_CURVE_SMOOTHING_PRESERVE_INITIAL_POINTS),
        )
        smoothing_window = st.slider(
            "Ширина окна вокруг скачка",
            min_value=0,
            max_value=10,
            value=int(config.HYDRAULIC_CURVE_SMOOTHING_WINDOW),
        )
        smoothing_strength = st.slider(
            "Сила сглаживания",
            min_value=0.0,
            max_value=1.0,
            value=float(config.HYDRAULIC_CURVE_SMOOTHING_STRENGTH),
            step=0.05,
        )
    report_name = st.text_input("Имя отчёта", value=f"{Path(uploaded.name).stem}_расчёт.docx")

calc_type = ["Вода", "Водокаменный селевой поток", "Грязекаменный селевой поток"].index(calc_type_label) + 1
calculation_key = hashlib.sha256(
    file_bytes
    + (
        f"{sheet_name}|{calc_type}|{shezi_type}|{overflow}|{use_depth}|{smooth_curve}|"
        f"{smoothing_tolerance}|{smoothing_points}|{smoothing_preserve}|"
        f"{smoothing_window}|{smoothing_strength}"
    ).encode()
).hexdigest()
# Последний расчёт сохраняем, чтобы показать его как устаревший до пересчёта.
source_key = hashlib.sha256(file_bytes + sheet_name.encode()).hexdigest()
if st.session_state.pop("profile_calc_completed_notice", False):
    st.toast("Расчёт текущего профиля выполнен.", icon="✅")
if st.session_state.get("single_report_calculation_key") not in (None, calculation_key):
    st.session_state.pop("single_profile_report", None)
    st.session_state.pop("single_profile_report_name", None)
    st.session_state.pop("single_report_calculation_key", None)
if st.session_state.get("all_profiles_calculation_key") not in (None, calculation_key):
    st.session_state.pop("all_profiles_report", None)
    st.session_state.pop("all_profiles_report_name", None)
    st.session_state.pop("all_profiles_calculation_key", None)

try:
    preview = read_sheet_preview(file_bytes, sheet_name)
except Exception as exc:
    st.error(f"Не удалось прочитать лист «{sheet_name}»: {exc}")
    st.stop()
st.subheader("Поперечный профиль")
calculated_profile = st.session_state.get("calculated_profile")
if calculated_profile is not None and st.session_state.get("calculated_source_key") == source_key:
    if st.session_state.get("calculation_key") != calculation_key:
        stale_notice_key = f"{st.session_state['calculation_key']}:{calculation_key}"
        if st.session_state.get("stale_notice_key") != stale_notice_key:
            st.toast(
                "Параметры изменены — на графике предыдущий расчёт. Пересчитайте профиль, чтобы обновить его.",
                icon="⚠️",
            )
            st.session_state["stale_notice_key"] = stale_notice_key
    st.pyplot(calculated_profile.fig_profile.fig, use_container_width=True)
    st.caption("Расчётные уровни по обеспеченности и исходный урез воды нанесены поверх профиля.")
else:
    render_profile(preview)
with st.expander("Предварительный просмотр исходных данных"):
    st.dataframe(preview, use_container_width=True, hide_index=True)

with st.sidebar:
    st.divider()
    calculate_clicked = st.button("Рассчитать текущий профиль", type="primary", use_container_width=True)
    report_clicked = st.button(
        "Создать отчёт по текущему профилю",
        use_container_width=True,
        disabled=st.session_state.get("calculation_key") != calculation_key,
    )
    all_profiles_clicked = st.button("Рассчитать все профили и создать отчёт", use_container_width=True)
    st.divider()

if calculate_clicked:
    with st.spinner("Выполняется расчёт текущего профиля…"):
        try:
            with tempfile.TemporaryDirectory(prefix="hydraulic_") as temp_dir:
                input_path = Path(temp_dir) / Path(uploaded.name).name
                input_path.write_bytes(file_bytes)

                config.CALC_TYPE = calc_type
                config.SHEZI_TYPE = shezi_type
                config.OVERFLOW = overflow
                config.USE_H_INSTEAD_R = use_depth
                config.HYDRAULIC_CURVE_SMOOTHED = smooth_curve
                config.HYDRAULIC_CURVE_SMOOTHING_TOLERANCE = smoothing_tolerance
                config.HYDRAULIC_CURVE_SMOOTHING_POINTS_PER_INTERVAL = smoothing_points
                config.HYDRAULIC_CURVE_SMOOTHING_PRESERVE_INITIAL_POINTS = smoothing_preserve
                config.HYDRAULIC_CURVE_SMOOTHING_WINDOW = smoothing_window
                config.HYDRAULIC_CURVE_SMOOTHING_STRENGTH = smoothing_strength
                config.SAVE_REPORT = False
                config.PROFILE_SAVE_PICTURES = False
                config.CURVE_SAVE_PICTURES = False

                calculated = profile.Morfostvor()
                calculated.read_xls(str(input_path), sheet_names.index(sheet_name))
                calculated.calculate()

            st.session_state["calculated_profile"] = calculated
            st.session_state["calculation_key"] = calculation_key
            st.session_state["calculated_source_key"] = source_key
            st.session_state["profile_calc_completed_notice"] = True
            st.rerun()
        except Exception as exc:
            st.error(f"Не удалось выполнить расчёт: {exc}")

if report_clicked:
    report_name = report_name if report_name.lower().endswith(".docx") else f"{report_name}.docx"
    calculated = st.session_state["calculated_profile"]
    try:
        with tempfile.TemporaryDirectory(prefix="hydraulic_report_") as temp_dir:
            output_path = Path(temp_dir) / Path(report_name).name
            generate_morfostvor_report(calculated, str(output_path))
            report_bytes = output_path.read_bytes()
        st.session_state["single_profile_report"] = report_bytes
        st.session_state["single_profile_report_name"] = Path(report_name).name
        st.session_state["single_report_calculation_key"] = calculation_key
    except Exception as exc:
        st.error(f"Не удалось создать отчёт: {exc}")

if st.session_state.get("single_profile_report"):
    with st.sidebar:
        st.download_button(
            "Скачать отчёт по текущему профилю",
            data=st.session_state["single_profile_report"],
            file_name=st.session_state["single_profile_report_name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="download_single_profile_report",
        )

if all_profiles_clicked:
    all_report_name = f"{Path(uploaded.name).stem}_все_профили.docx"
    start_notice = st.toast("Расчёт всех профилей запущен, пожалуйста подождите…", icon="⏳", duration="infinite")
    with st.spinner("Выполняется расчёт всех профилей и формируется общий отчёт, пожалуйста подождите…"):
        try:
            with tempfile.TemporaryDirectory(prefix="hydraulic_all_profiles_") as temp_dir:
                input_path = Path(temp_dir) / Path(uploaded.name).name
                output_path = Path(temp_dir) / all_report_name
                input_path.write_bytes(file_bytes)

                config.CALC_TYPE = calc_type
                config.SHEZI_TYPE = shezi_type
                config.OVERFLOW = overflow
                config.USE_H_INSTEAD_R = use_depth
                config.HYDRAULIC_CURVE_SMOOTHED = smooth_curve
                config.HYDRAULIC_CURVE_SMOOTHING_TOLERANCE = smoothing_tolerance
                config.HYDRAULIC_CURVE_SMOOTHING_POINTS_PER_INTERVAL = smoothing_points
                config.HYDRAULIC_CURVE_SMOOTHING_PRESERVE_INITIAL_POINTS = smoothing_preserve
                config.HYDRAULIC_CURVE_SMOOTHING_WINDOW = smoothing_window
                config.HYDRAULIC_CURVE_SMOOTHING_STRENGTH = smoothing_strength
                config.SAVE_REPORT = True
                config.PROFILE_SAVE_PICTURES = False
                config.CURVE_SAVE_PICTURES = False

                profile.xls_calculate_hydraulic(str(input_path), str(output_path))
                all_report_bytes = output_path.read_bytes()

            st.session_state["all_profiles_report"] = all_report_bytes
            st.session_state["all_profiles_report_name"] = all_report_name
            st.session_state["all_profiles_calculation_key"] = calculation_key
            start_notice.toast("Расчёт всех профилей завершён. Отчёт готов.", icon="✅", duration="short")
        except Exception as exc:
            start_notice.toast(f"Не удалось рассчитать все профили: {exc}", icon="❌", duration="long")
            st.error(f"Не удалось рассчитать все профили: {exc}")

if st.session_state.get("all_profiles_report"):
    with st.sidebar:
        st.download_button(
            "Скачать отчёт по всем профилям",
            data=st.session_state["all_profiles_report"],
            file_name=st.session_state["all_profiles_report_name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="download_all_profiles_report",
        )

calculated = st.session_state.get("calculated_profile")
if calculated is not None and st.session_state.get("calculation_key") == calculation_key:
    st.subheader("Расчётные уровни по обеспеченности")
    st.dataframe(calculated.levels_result, use_container_width=True, hide_index=True)

    st.subheader("Параметры расчётных участков")
    sector_results = calculated.sectors_result.rename(
        columns={
            "name": "Участок",
            "slope": "Уклон, ‰",
            "roughness": "Шероховатость n",
            "consumption": "Расход Q, м³/с",
            "depth": "Средняя глубина, м",
            "speed": "Скорость V, м/с",
            "width": "Ширина, м",
            "area": "Площадь, м²",
        }
    )
    st.dataframe(sector_results, use_container_width=True, hide_index=True)

    with st.expander("Полная таблица расчёта по уровням", expanded=False):
        st.dataframe(calculated.hydraulic_table.reset_index(), use_container_width=True, hide_index=True)

    st.subheader("Графики расчёта")
    chart_rows = [st.columns(2), st.columns(2)]
    for col, graph, caption in [
        (chart_rows[0][0], calculated.fig_profile, "Поперечный профиль и расчётные уровни"),
        (chart_rows[0][1], calculated.fig_qh, "Гидравлическая кривая Q(H)"),
        (chart_rows[1][0], calculated.fig_qv, "Расход и скорость Q(V)"),
        (chart_rows[1][1], calculated.fig_qwvh, "Расход, площадь и скорость от уровня"),
    ]:
        with col:
            st.pyplot(graph.fig, use_container_width=True)
            st.caption(caption)
