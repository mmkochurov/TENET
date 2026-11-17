import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit_authenticator as stauth

st.set_page_config(page_title="Воронка продаж Chery", layout="wide", page_icon="📊")

# ============ АВТОРИЗАЦИЯ ============
try:
    # Получаем секреты
    raw_config = st.secrets["auth"]
    
    # Ключевое изменение: преобразуем AttrDict в обычный dict
    # для credentials и всех вложенных словарей пользователей
    credentials = {
        "usernames": {
            username: dict(user_data)
            for username, user_data in raw_config["credentials"]["usernames"].items()
        }
    }
    
    # Собираем итоговый конфиг в виде обычного словаря
    config = {
        "credentials": credentials,
        "cookie": dict(raw_config["cookie"]),
        "preauthorized": dict(raw_config.get("preauthorized", {})),
    }

except Exception:
    # Локальный запуск из auth.yaml (там уже обычный dict, конвертация не нужна)
    with open("auth.yaml", encoding="utf-8") as f:
        config = yaml.safe_load(f)

# Теперь config — это обычный словарь, и Authenticate сможет в него писать
authenticator = stauth.Authenticate(
    config["credentials"],
    config["cookie"]["name"],
    config["cookie"]["key"],
    config["cookie"]["expiry_days"],
    config.get("preauthorized"),
)

name, auth_status, username = authenticator.login("Вход в дашборд", "main")

if auth_status is False:
    st.error("Неверный логин или пароль")
    st.stop()
elif auth_status is None:
    st.info("Введите логин и пароль")
    st.stop()

# ============ ПАРСЕР 1С ============
STAGE_COLS = {
    "leads": 5,             # Количество РЛОП
    "phone": 6,             # Мобильный телефон
    "test_drive": 7,        # Тест-драйв
    "order": 8,             # Заказ на автомобиль
    "sale": 9,              # Реализация автомобилей
    "tradein_lead": 10,     # РЛ Trade-IN
    "tradein_received": 11, # Поступление а/м Trade-IN
}
CONTACT_TYPES = {
    "Личная встреча", "Телефонный звонок", "Online консультант",
    "Электронное письмо", "Входящий звонок", "Исходящий звонок",
}
SKIP_LABELS = {
    "Подразделение", "Менеджер РЛОП", "Тип первичного контакта",
    'Модель "реализация"', "Мобильный телефон", "Наличие реализации",
    "Наличие РЛ Trade-IN", "Статус", "РЛОП", "VINПотребность", "Итого",
}
PHONE_RE = re.compile(r"^\+?\d[\d\s\-\(\)]{8,15}$")
MANAGER_RE = re.compile(r"^[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+$")


def _num(v) -> int:
    try:
        return int(float(v)) if pd.notna(v) and str(v).strip() != "" else 0
    except (ValueError, TypeError):
        return 0


@st.cache_data(ttl=600, show_spinner="Парсим Excel 1С...")
def parse_1c_funnel(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(file_bytes, header=None)
    records = []
    cur_manager, cur_contact, cur_model = None, None, None

    for i in range(len(df)):
        label = str(df.iloc[i, 0]).strip() if pd.notna(df.iloc[i, 0]) else ""
        if not label:
            continue
        if any(label.startswith(p) for p in
               ("Реестр", "Параметры", "Дата ", "Отбор", "Автоматич", "Отдел ")):
            continue
        if label in SKIP_LABELS:
            continue

        # --- Лид (телефон) ---
        if PHONE_RE.fullmatch(label):
            rec = {
                "phone": re.sub(r"\D", "", label),
                "manager": cur_manager or "—",
                "contact_type": cur_contact or "—",
                "model": cur_model or "—",
            }
            for k, col in STAGE_COLS.items():
                rec[k] = _num(df.iloc[i, col])
            rec.update(has_sale=None, has_tradein=None, status=None)

            # Атрибуты под лидом
            for j in range(i + 1, min(i + 10, len(df))):
                nl = str(df.iloc[j, 0]).strip() if pd.notna(df.iloc[j, 0]) else ""
                if not nl:
                    continue
                if PHONE_RE.fullmatch(nl):
                    break
                if nl in ("есть реализация", "нет реализации"):
                    rec["has_sale"] = nl == "есть реализация"
                elif nl in ("Да", "Нет"):
                    rec["has_tradein"] = nl == "Да"
                elif nl in ("Текущий", "Согласование ДЗ у РОП", "Не имеет продолжения"):
                    rec["status"] = nl
                elif nl.startswith("МТР"):
                    break
            records.append(rec)
            continue

        if label in CONTACT_TYPES:
            cur_contact = label
            continue
        if MANAGER_RE.match(label):
            cur_manager, cur_contact, cur_model = label, None, None
            continue
        if re.match(r"^[A-Z][A-Z0-9 ]{1,20}$", label):
            cur_model = label
            continue

    return pd.DataFrame(records)


# ============ ЗАГРУЗКА ДАННЫХ ============
authenticator.logout("Выйти", "sidebar")
st.sidebar.markdown(f"👤 **{name}**")

uploaded = st.sidebar.file_uploader(
    "Загрузить новый Excel из 1С", type=["xlsx", "xls"]
)

DATA_PATH = Path("data/funnel.xlsx")
if uploaded is not None:
    df = parse_1c_funnel(uploaded.getvalue())
    st.sidebar.success(f"Загружено лидов: {len(df)}")
elif DATA_PATH.exists():
    df = parse_1c_funnel(DATA_PATH.read_bytes())
else:
    st.warning("Загрузите Excel-файл в сайдбаре для начала работы")
    st.stop()

if df.empty:
    st.error("Не удалось распарсить файл. Проверьте, что это выгрузка «Реестр рабочих листов» из 1С.")
    st.stop()

# ============ ФИЛЬТРЫ ============
st.sidebar.header("Фильтры")
managers = st.sidebar.multiselect(
    "Менеджер", sorted(df["manager"].unique()), default=sorted(df["manager"].unique())
)
contacts = st.sidebar.multiselect(
    "Тип контакта", sorted(df["contact_type"].unique()),
    default=sorted(df["contact_type"].unique()),
)
models = st.sidebar.multiselect(
    "Модель", sorted(df["model"].unique()), default=sorted(df["model"].unique())
)

f = df[df["manager"].isin(managers)
       & df["contact_type"].isin(contacts)
       & df["model"].isin(models)]

# ============ KPI + ВОРОНКА ============
st.title("📊 Воронка продаж — Отдел Chery")

stages = [
    ("Лиды (РЛОП)",       "leads"),
    ("Контакт (телефон)", "phone"),
    ("Тест-драйв",        "test_drive"),
    ("Заказ на авто",     "order"),
    ("Реализация",        "sale"),
]
counts = [(label, int(f[col].sum())) for label, col in stages]

cols = st.columns(len(counts))
for c, (label, val) in zip(cols, counts):
    prev = counts[0][1]
    delta = f"{(val / prev * 100):.1f}% от лидов" if prev else "—"
    c.metric(label, val, delta)

st.subheader("Воронка продаж")
fig = go.Figure(go.Funnel(
    y=[n for n, _ in counts],
    x=[v for _, v in counts],
    textinfo="value+percent initial",
    marker=dict(color="#2E86DE"),
))
fig.update_layout(height=420, margin=dict(l=10, r=10, t=10, b=10))
st.plotly_chart(fig, use_container_width=True)

# ============ КОНВЕРСИИ ============
st.subheader("Конверсии по шагам")
conv = []
for i in range(len(counts) - 1):
    a, b = counts[i][1], counts[i + 1][1]
    conv.append({
        "Переход": f"{counts[i][0]} → {counts[i + 1][0]}",
        "Было": a, "Стало": b,
        "Конверсия %": round(b / a * 100, 1) if a else 0,
    })
st.dataframe(pd.DataFrame(conv), use_container_width=True, hide_index=True)

# ============ ПО МЕНЕДЖЕРАМ ============
st.subheader("По менеджерам")
by_mgr = f.groupby("manager").agg(
    Лиды=("leads", "sum"),
    ТД=("test_drive", "sum"),
    Заказы=("order", "sum"),
    Продажи=("sale", "sum"),
    Trade_IN_заявки=("tradein_lead", "sum"),
    Trade_IN_поступления=("tradein_received", "sum"),
).reset_index()
by_mgr["CR %"] = (by_mgr["Продажи"] / by_mgr["Лиды"] * 100).round(1)
st.dataframe(
    by_mgr.sort_values("Продажи", ascending=False),
    use_container_width=True, hide_index=True,
)

fig2 = px.bar(
    by_mgr, x="manager", y=["Лиды", "ТД", "Заказы", "Продажи"],
    barmode="group", title="Активность менеджеров по этапам",
)
fig2.update_layout(height=380, xaxis_title="", yaxis_title="Количество")
st.plotly_chart(fig2, use_container_width=True)

# ============ ПО КАНАЛАМ / МОДЕЛЯМ ============
c1, c2 = st.columns(2)

with c1:
    st.subheader("По типу контакта")
    by_ct = f.groupby("contact_type").agg(
        Лиды=("leads", "sum"), Продажи=("sale", "sum")
    ).reset_index()
    by_ct["CR %"] = (by_ct["Продажи"] / by_ct["Лиды"] * 100).round(1)
    st.dataframe(by_ct.sort_values("Лиды", ascending=False),
                 use_container_width=True, hide_index=True)

with c2:
    st.subheader("По моделям")
    by_model = f[f["model"] != "—"].groupby("model").agg(
        Лиды=("leads", "sum"), Продажи=("sale", "sum")
    ).reset_index()
    if not by_model.empty:
        by_model["CR %"] = (by_model["Продажи"] / by_model["Лиды"] * 100).round(1)
        st.dataframe(by_model.sort_values("Продажи", ascending=False),
                     use_container_width=True, hide_index=True)
    else:
        st.info("Модели не указаны в выгрузке")

# ============ TRADE-IN ============
st.subheader("Trade-IN")
c3, c4 = st.columns(2)
with c3:
    ti = pd.DataFrame({
        "Этап": ["Заявки Trade-IN", "Поступление авто"],
        "Количество": [int(f["tradein_lead"].sum()), int(f["tradein_received"].sum())],
    })
    fig3 = px.funnel(ti, x="Количество", y="Этап")
    st.plotly_chart(fig3, use_container_width=True)

with c4:
    st.markdown("**Статусы лидов**")
    st_data = f["status"].fillna("Не указан").value_counts().reset_index()
    st_data.columns = ["Статус", "Количество"]
    st.dataframe(st_data, use_container_width=True, hide_index=True)

st.caption(f"Обновлено: {pd.Timestamp.now():%d.%m.%Y %H:%M} · лидов в выборке: {len(f)}")
