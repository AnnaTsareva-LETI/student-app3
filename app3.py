import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

# ============================================================
# Настройка страницы
# ============================================================
st.set_page_config(
    page_title="Лабораторная работа: квантование сигнала",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# CSS: стиль лабораторного стенда
# ============================================================
st.markdown("""
<style>
    .stApp {
        background:
            radial-gradient(circle at top, #2f3b45 0%, #1f252b 45%, #15181c 100%);
        color: #f2f2f2;
    }

    .lab-header {
        background: linear-gradient(90deg, #0d47a1 0%, #1565c0 50%, #1e88e5 100%);
        border-radius: 20px;
        padding: 22px 26px;
        margin-bottom: 18px;
        box-shadow: 0 10px 28px rgba(0,0,0,0.35);
        border: 1px solid rgba(255,255,255,0.15);
    }

    .lab-title {
        font-size: 30px;
        font-weight: 800;
        color: #ffffff;
        margin-bottom: 6px;
    }

    .lab-subtitle {
        font-size: 15px;
        color: #d9ecff;
    }

    .panel {
        background: linear-gradient(180deg, #2d3339 0%, #22272d 100%);
        border: 1px solid #4c5964;
        border-radius: 20px;
        padding: 18px;
        box-shadow: 0 10px 24px rgba(0,0,0,0.28);
        margin-bottom: 18px;
    }

    .panel-light {
        background: linear-gradient(180deg, #f4f6f8 0%, #e9edf1 100%);
        border: 1px solid #b0bcc8;
        border-radius: 20px;
        padding: 18px;
        box-shadow: 0 10px 24px rgba(0,0,0,0.12);
        margin-bottom: 18px;
        color: #1c1f22;
    }

    .panel-title {
        font-size: 22px;
        font-weight: 800;
        margin-bottom: 8px;
        color: inherit;
    }

    .panel-note {
        font-size: 13px;
        opacity: 0.9;
        margin-bottom: 10px;
    }

    .lcd {
        background: linear-gradient(180deg, #071007 0%, #0c1a0c 100%);
        border: 4px solid #0e120e;
        border-radius: 18px;
        padding: 18px;
        box-shadow:
            inset 0 0 16px rgba(124,252,0,0.12),
            0 8px 18px rgba(0,0,0,0.3);
        color: #7CFC00;
        font-family: "Consolas", monospace;
    }

    .lcd-value {
        font-size: 46px;
        font-weight: 800;
        text-align: right;
        line-height: 1.0;
    }

    .lcd-unit {
        font-size: 20px;
        font-weight: 700;
        color: #a7ff8a;
        margin-left: 8px;
    }

    .lcd-sub {
        font-size: 12px;
        color: #9fff9f;
        margin-top: 8px;
        text-align: left;
    }

    .badge-box {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 700;
        margin-left: 8px;
        color: white;
        background: #2e7d32;
    }

    .badge-off {
        background: #c62828;
    }

    .section-title {
        font-size: 20px;
        font-weight: 800;
        margin: 16px 0 10px 0;
        color: #ffffff;
    }

    .status-box {
        background: rgba(255,255,255,0.06);
        border-left: 5px solid #42a5f5;
        padding: 12px 14px;
        border-radius: 12px;
        color: #eaf4ff;
        margin-top: 12px;
    }

    .result-ok {
        background: #e8f5e9;
        border-left: 5px solid #2e7d32;
        padding: 12px 14px;
        border-radius: 12px;
        color: #1b5e20;
        margin-top: 12px;
    }

    .result-info {
        background: #e3f2fd;
        border-left: 5px solid #1565c0;
        padding: 12px 14px;
        border-radius: 12px;
        color: #0d47a1;
        margin-top: 12px;
    }

    .stButton > button {
        border-radius: 10px;
        font-weight: 700;
        border: 1px solid #4c5964;
    }

    .stDataFrame {
        border-radius: 12px;
        overflow: hidden;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# Функции
# ============================================================
def quantize_signal(x, x_min, x_max, bits):
    levels = 2 ** bits
    q = (x_max - x_min) / (levels - 1)

    x_clipped = np.clip(x, x_min, x_max)
    code = np.round((x_clipped - x_min) / q)
    x_quantized = x_min + code * q
    error = x_quantized - x
    return x_quantized, error, q, code


def create_excel_file(df_results):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_results.to_excel(writer, sheet_name="Results", index=False)
    output.seek(0)
    return output


def fmt(v):
    return f"{v:.8f}"


# ============================================================
# Заголовок
# ============================================================
st.markdown("""
<div class="lab-header">
    <div class="lab-title">🧪 Лабораторная работа: исследование погрешности квантования</div>
    <div class="lab-subtitle">
        Моделирование АЦП, анализ ошибки квантования, сравнение разрядностей и экспорт результатов в Excel
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================================
# Боковая панель управления
# ============================================================
st.sidebar.title("⚙️ Параметры стенда")

st.sidebar.markdown("### Диапазон квантования")
x_min = st.sidebar.number_input("Нижний предел, В", value=0.0, format="%.3f")
x_max = st.sidebar.number_input("Верхний предел, В", value=5.0, format="%.3f")

st.sidebar.markdown("### Исходный сигнал")
amplitude = st.sidebar.number_input("Амплитуда, В", value=2.0, format="%.3f")
offset = st.sidebar.number_input("Смещение, В", value=2.5, format="%.3f")
frequency = st.sidebar.number_input("Частота, Гц", value=5.0, format="%.3f")

st.sidebar.markdown("### Дискретизация")
sampling_frequency = st.sidebar.number_input("Частота дискретизации, Гц", value=1000, min_value=10, step=10)
duration = st.sidebar.number_input("Длительность, с", value=1.0, format="%.3f")

st.sidebar.markdown("### Квантование")
bits_for_plot = st.sidebar.slider("Разрядность для графика", 1, 16, 4)
bits_list = st.sidebar.multiselect(
    "Разрядности для сравнения",
    options=[2, 3, 4, 5, 6, 8, 10, 12, 14, 16],
    default=[4, 6, 8, 10, 12, 16]
)

if x_max <= x_min:
    st.error("Верхний предел должен быть больше нижнего предела диапазона.")
    st.stop()

# ============================================================
# Формирование сигнала
# ============================================================
t = np.linspace(0, duration, int(sampling_frequency * duration), endpoint=False)
x = amplitude * np.sin(2 * np.pi * frequency * t) + offset

x_q, error, q, code = quantize_signal(x, x_min, x_max, bits_for_plot)
rms_error = np.sqrt(np.mean(error ** 2))
max_error = np.max(np.abs(error))

results = []
for bits in bits_list:
    x_q_i, error_i, q_i, code_i = quantize_signal(x, x_min, x_max, bits)
    results.append({
        "bits": bits,
        "levels": 2 ** bits,
        "q": q_i,
        "max_error": np.max(np.abs(error_i)),
        "rms_error": np.sqrt(np.mean(error_i ** 2)),
        "theoretical_max": q_i / 2,
        "theoretical_rms": q_i / np.sqrt(12)
    })

df = pd.DataFrame(results)

# ============================================================
# Верхние индикаторы
# ============================================================
st.markdown("## 📟 Панель контроля")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Разрядность", f"{bits_for_plot} бит")
c2.metric("Шаг квантования q", fmt(q))
c3.metric("Макс. ошибка |Δq|", fmt(max_error))
c4.metric("СКО ошибки", fmt(rms_error))

# ============================================================
# Схема установки
# ============================================================
st.markdown("## 🧩 Лабораторный стенд")

left, right = st.columns([1.1, 0.9])

with left:
    st.markdown("""
    <div class="panel">
        <div class="panel-title">Генератор сигнала</div>
        <div class="panel-note">Синусоидальный сигнал на входе АЦП</div>
    """, unsafe_allow_html=True)

    st.markdown(
        f"""
        <div class="status-box">
            <b>Диапазон:</b> {x_min:.3f} ... {x_max:.3f} В<br>
            <b>Амплитуда:</b> {amplitude:.3f} В<br>
            <b>Смещение:</b> {offset:.3f} В<br>
            <b>Частота:</b> {frequency:.3f} Гц<br>
            <b>Длительность:</b> {duration:.3f} с
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("</div>", unsafe_allow_html=True)

with right:
    st.markdown("""
    <div class="panel-light">
        <div class="panel-title">АЦП / Анализ квантования</div>
        <div class="panel-note">Текущая разрядность и параметры дискретизации</div>
    """, unsafe_allow_html=True)

    st.markdown(
        f"""
        <div class="result-info">
            <b>Разрядность для графика:</b> {bits_for_plot} бит<br>
            <b>Число уровней:</b> {2 ** bits_for_plot}<br>
            <b>Шаг квантования:</b> {q:.8f} В<br>
            <b>Теоретическая max ошибка:</b> {(q/2):.8f} В<br>
            <b>Теоретическое СКО:</b> {(q/np.sqrt(12)):.8f} В
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# Графики
# ============================================================
st.markdown("## 📈 Осциллограммы и ошибка")

fig1, ax1 = plt.subplots(figsize=(11, 5))
ax1.set_facecolor("#f7f9fb")
ax1.plot(t, x, color="#1565c0", linewidth=2.2, label="Исходный аналоговый сигнал")
ax1.step(t, x_q, where="mid", color="#2e7d32", linewidth=1.8, label=f"Квантованный сигнал, {bits_for_plot} бит")
ax1.set_xlabel("Время, с")
ax1.set_ylabel("Напряжение, В")
ax1.set_title("Исходный и квантованный сигнал")
ax1.grid(True, alpha=0.35)
ax1.legend()
st.pyplot(fig1, clear_figure=True)

fig2, ax2 = plt.subplots(figsize=(11, 4))
ax2.set_facecolor("#f7f9fb")
ax2.plot(t, error, color="#c62828", linewidth=1.7, label="Ошибка квантования")
ax2.axhline(q / 2, color="black", linestyle="--", linewidth=1.2, label="+q/2")
ax2.axhline(-q / 2, color="black", linestyle="--", linewidth=1.2, label="-q/2")
ax2.set_xlabel("Время, с")
ax2.set_ylabel("Ошибка, В")
ax2.set_title(f"Ошибка квантования при {bits_for_plot} бит")
ax2.grid(True, alpha=0.35)
ax2.legend()
st.pyplot(fig2, clear_figure=True)

# ============================================================
# Таблица результатов
# ============================================================
st.markdown("## 📋 Таблица расчётов")

show_df = df.copy()
for col in ["q", "max_error", "rms_error", "theoretical_max", "theoretical_rms"]:
    show_df[col] = show_df[col].map(lambda v: f"{v:.8f}")

st.dataframe(show_df, use_container_width=True, hide_index=True)

# ============================================================
# График зависимости ошибок от разрядности
# ============================================================
st.markdown("## 🔬 Влияние разрядности АЦП на ошибку квантования")

fig3, ax3 = plt.subplots(figsize=(10, 5))
ax3.set_facecolor("#f7f9fb")
ax3.plot(df["bits"], df["q"], "o-", linewidth=2, label="Шаг квантования q")
ax3.plot(df["bits"], df["max_error"], "s-", linewidth=2, label="Макс. |Δq|")
ax3.plot(df["bits"], df["rms_error"], "^-", linewidth=2, label="СКО ошибки")
ax3.plot(df["bits"], df["theoretical_rms"], "x--", linewidth=2, label="q/sqrt(12)")
ax3.set_xlabel("Разрядность АЦП, бит")
ax3.set_ylabel("Значение, В")
ax3.set_title("Влияние разрядности на ошибку квантования")
ax3.grid(True, alpha=0.35)
ax3.set_yscale("log")
ax3.legend()
st.pyplot(fig3, clear_figure=True)

# ============================================================
# Выводы
# ============================================================
st.markdown("## 🧠 Анализ результатов")

st.markdown(
    f"""
    <div class="result-ok">
        <b>Для выбранной разрядности {bits_for_plot} бит:</b><br>
        Шаг квантования q = {q:.8f} В<br>
        Максимальная ошибка |Δq| = {max_error:.8f} В<br>
        СКО ошибки = {rms_error:.8f} В<br>
        Теоретическое значение q/2 = {(q/2):.8f} В
    </div>
    """,
    unsafe_allow_html=True
)

# ============================================================
# Экспорт в Excel
# ============================================================
st.markdown("## 💾 Экспорт результатов")

excel_file = create_excel_file(df)

st.download_button(
    label="Скачать таблицу результатов Excel",
    data=excel_file,
    file_name="quantization_results.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ============================================================
# Справка
# ============================================================
with st.expander("Справка"):
    st.write(
        """
        **Назначение программы**  
        Моделирование квантования аналогового сигнала в АЦП и исследование погрешности квантования.

        **Формулы**
        - Шаг квантования: `q = (x_max - x_min) / (2^bits - 1)`
        - Теоретическая максимальная ошибка: `q/2`
        - Теоретическое СКО ошибки: `q/sqrt(12)`

        **Что можно менять**
        - диапазон квантования;
        - амплитуду и смещение сигнала;
        - частоту сигнала;
        - длительность;
        - частоту дискретизации;
        - разрядность АЦП.
        """
    )