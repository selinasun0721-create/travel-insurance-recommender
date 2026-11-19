
import streamlit as st
import pandas as pd
import json
import re
from pathlib import Path
from datetime import date, timedelta

# =====================================================
# 1. 基本設定
# =====================================================

st.set_page_config(
    page_title="旅旅｜智慧旅遊平安險推薦",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded"
)

ROOT = Path(__file__).resolve().parent.parent

DATA_FILE = (
    ROOT
    / "01_Excel資料"
    / "旅遊保險推薦系統_第一階段資料蒐集模板.xlsx"
)

USER_FILE = Path(__file__).resolve().parent / "users.json"


# =====================================================
# 2. 網頁外觀
# =====================================================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Noto Sans TC', sans-serif;
}

.stApp {
    background: linear-gradient(
        180deg,
        #fff9f5 0%,
        #f8f5ff 55%,
        #f2faff 100%
    );
}

[data-testid="stSidebar"] {
    background: linear-gradient(
        180deg,
        #fff0e9,
        #f1edff
    );
}

.block-container {
    padding-top: 1.5rem;
    padding-bottom: 3rem;
    max-width: 1250px;
}

.hero {
    background: linear-gradient(
        115deg,
        #ffc5b5,
        #ffd9cb,
        #d4ceff
    );
    padding: 32px;
    border-radius: 26px;
    color: #59465d;
    margin-bottom: 24px;
    box-shadow: 0 8px 25px rgba(140, 105, 150, .12);
}

.hero h1 {
    color: #59465d;
    font-weight: 800;
}

.hero p {
    color: #715d75;
    font-size: 1rem;
}

.pill {
    display: inline-block;
    background: rgba(255,255,255,.75);
    padding: 6px 13px;
    border-radius: 30px;
    color: #806782;
    font-size: .85rem;
    font-weight: 600;
    margin-bottom: 10px;
}

.soft-card {
    background: rgba(255,255,255,.9);
    border: 1px solid #f0e4ef;
    border-radius: 20px;
    padding: 20px;
    box-shadow: 0 5px 18px rgba(120,90,130,.06);
}

div[data-testid="stMetric"] {
    background: rgba(255,255,255,.9);
    border: 1px solid #f0e4ef;
    padding: 15px;
    border-radius: 18px;
}

div[data-testid="stForm"] {
    background: rgba(255,255,255,.92);
    border: 1px solid #f0e4ef;
    padding: 22px;
    border-radius: 20px;
}

.stButton > button,
.stFormSubmitButton > button {
    border-radius: 14px;
    font-weight: 700;
    min-height: 2.7rem;
}

div[data-testid="stChatMessage"] {
    border-radius: 18px;
    border: 1px solid #f0e7ef;
    background: rgba(255,255,255,.85);
}

div[data-testid="stExpander"] {
    border-radius: 16px;
    border-color: #eee2ed;
}

[data-testid="stTabs"] button {
    font-weight: 700;
}

</style>
""", unsafe_allow_html=True)


st.markdown("""
<div class="hero">
    <div class="pill">🌷 讓每趟旅行，都多一份安心</div>
    <h1>✈️ 旅旅・智慧旅遊平安險小幫手</h1>
    <p>
        認識你的旅遊習慣，幫你整理適合的旅遊平安險，
        讓每一次出發都更加安心。
    </p>
</div>
""", unsafe_allow_html=True)


# =====================================================
# 3. 使用者資料管理
# =====================================================

def load_users():

    if USER_FILE.exists():

        try:
            return json.loads(
                USER_FILE.read_text(encoding="utf-8")
            )

        except Exception:
            return {}

    return {}


def save_user(profile):

    users = load_users()

    # 健康資料不寫入本機使用者資料檔
    safe_profile = {
        key: value
        for key, value in profile.items()
        if key not in ["健康狀況", "健康補充"]
    }

    users[profile["姓名"]] = safe_profile

    USER_FILE.write_text(
        json.dumps(
            users,
            ensure_ascii=False,
            indent=4
        ),
        encoding="utf-8"
    )


if "profile" not in st.session_state:
    st.session_state.profile = None

if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

if "chat_active" not in st.session_state:
    st.session_state.chat_active = False

if "trip_requirements" not in st.session_state:
    st.session_state.trip_requirements = None

if "chat_recommendations" not in st.session_state:
    st.session_state.chat_recommendations = None


# =====================================================
# 4. Excel 資料讀取
# =====================================================

def normalize_columns(df):

    df = df.copy()

    df.columns = [
        re.sub(r"\s+", "", str(column)).replace("／", "/")
        for column in df.columns
    ]

    aliases = {
        "公司": ["公司", "保險公司"],
        "旅遊範圍": ["旅遊範圍", "旅遊地區"],
        "商品/方案": [
            "商品/方案",
            "商品/方案名稱",
            "方案名稱"
        ],
        "標準化保障項目": [
            "標準化保障項目",
            "保障項目"
        ],
        "保障額度/文字": [
            "保障額度/文字",
            "保障額度",
            "保額/文字"
        ],
        "系統分類": ["系統分類"],
        "給付方式": ["給付方式"],
        "重要限制或備註": [
            "重要限制或備註",
            "備註"
        ]
    }

    rename = {}

    for target, candidates in aliases.items():

        for candidate in candidates:

            if candidate in df.columns:
                rename[candidate] = target
                break

    return df.rename(columns=rename)


@st.cache_data(show_spinner=False)
def load_insurance_data(file_path):

    if not Path(file_path).exists():
        return pd.DataFrame(), "找不到 Excel 資料檔"

    try:

        excel = pd.ExcelFile(file_path)

        if "保障明細_整理版" in excel.sheet_names:
            sheet = "保障明細_整理版"

        elif "保障明細" in excel.sheet_names:
            sheet = "保障明細"

        else:
            return pd.DataFrame(), "找不到保障明細工作表"

        df = pd.read_excel(
            file_path,
            sheet_name=sheet,
            header=3
        )

        df = normalize_columns(df)

        df = df.dropna(how="all")

        return df, None

    except Exception as error:

        return pd.DataFrame(), str(error)


if DATA_FILE.exists():

    df, data_error = load_insurance_data(
        str(DATA_FILE)
    )

else:

    df = pd.DataFrame()
    data_error = "找不到預設 Excel 資料檔"


# =====================================================
# 5. 文字處理
# =====================================================

def clean_text(value):

    if pd.isna(value):
        return ""

    return str(value).strip()


# =====================================================
# 6. 建立保險方案
# =====================================================

def get_plans(data):

    if data.empty:
        return []

    required = [
        "公司",
        "旅遊範圍",
        "商品/方案",
        "標準化保障項目",
        "保障額度/文字"
    ]

    if not all(
        column in data.columns
        for column in required
    ):
        return []

    working = data.copy()

    optional_columns = [
        "系統分類",
        "重要限制或備註",
        "給付方式"
    ]

    for column in optional_columns:

        if column not in working.columns:
            working[column] = ""

    plans = []

    for (company, scope, product), group in working.groupby(
        ["公司", "旅遊範圍", "商品/方案"],
        dropna=False
    ):

        company = clean_text(company)
        scope = clean_text(scope)
        product = clean_text(product)

        if not product:
            continue

        benefits = {}
        notes = []

        for _, row in group.iterrows():

            item = clean_text(
                row["標準化保障項目"]
            )

            amount = clean_text(
                row["保障額度/文字"]
            )

            category = clean_text(
                row["系統分類"]
            )

            note = clean_text(
                row["重要限制或備註"]
            )

            # 排除旅遊不便險
            excluded_words = [
                "班機延誤",
                "行李延誤",
                "旅程取消",
                "旅程更改",
                "行李損失"
            ]

            if category == "排除":
                continue

            if any(word in item for word in excluded_words):
                continue

            if not item:
                continue

            if (
                "意外身故" in item
                or "失能" in item
                or "身故及失能" in item
            ):
                key = "意外身故／失能"

            elif "傷害醫療" in item:
                key = "傷害醫療"

            elif "突發疾病" in item:
                key = "海外突發疾病醫療"

            elif "緊急救援" in item:
                key = "緊急救援"

            else:
                continue

            benefits[key] = amount if amount else "資料待補"

            if note and note not in notes:
                notes.append(note)

        if benefits:

            plans.append({
                "公司": company,
                "旅遊範圍": scope,
                "方案": product,
                "保障": benefits,
                "備註": notes
            })

    return plans


plans = get_plans(df)


# =====================================================
# 7. 推薦演算法
# =====================================================

WEIGHTS = {
    "意外身故／失能": 45,
    "傷害醫療": 35,
    "海外突發疾病醫療": 20
}


def recommend(all_plans, scope_choice="全部", needs=None):

    needs = needs or []

    results = []

    for plan in all_plans:

        scope = plan["旅遊範圍"]

        if scope_choice == "海外" and "海外" not in scope:
            continue

        if scope_choice == "國內" and "國內" not in scope:
            continue

        benefits = plan["保障"]

        score = 0
        matched = []
        missing = []

        active_needs = [
            n for n in needs
            if not (
                n == "海外突發疾病醫療"
                and "國內" in scope
            )
        ]

        if active_needs:

            for need in active_needs:

                # 緊急救援列為補充資訊，不計入核心分數
                if need == "緊急救援":
                    continue

                if need in benefits:
                    score += WEIGHTS.get(need, 0)
                    matched.append(need)

                else:
                    missing.append(need)

        else:

            score = sum(
                WEIGHTS.get(key, 0)
                for key in benefits
                if key in WEIGHTS
            )

        results.append({

            "公司": plan["公司"],
            "方案": plan["方案"],
            "旅遊範圍": plan["旅遊範圍"],
            "分數": score,
            "保障": benefits,
            "符合項目": matched,
            "資料待確認": missing,
            "備註": plan["備註"]

        })

    results.sort(
        key=lambda item: (
            item["分數"],
            len(item["保障"])
        ),
        reverse=True
    )

    return results


# =====================================================
# 8. 推薦方案卡片
# =====================================================

def render_plan_card(item, rank, compact=False):

    with st.container(border=True):

        col1, col2 = st.columns([4, 1])

        with col1:

            st.markdown(
                f"#### {'🌟' if rank == 1 else '🧳'} 第 {rank} 名｜{item['公司']}"
            )

            st.markdown(
                f"**{item['方案']}**"
            )

            st.caption(
                f"適用範圍：{item['旅遊範圍']}"
            )

        with col2:

            st.metric(
                "資料符合分數",
                f"{item['分數']} 分"
            )

        if item["符合項目"]:

            st.success(
                "符合保障：" +
                "、".join(item["符合項目"])
            )

        if item["資料待確認"]:

            st.info(
                "資料尚待確認：" +
                "、".join(item["資料待確認"])
            )

        if not compact:

            st.markdown("**保障內容**")

            columns = st.columns(2)

            for index, (key, value) in enumerate(
                item["保障"].items()
            ):

                with columns[index % 2]:

                    st.write(f"**{key}**")
                    st.write(value)

            if item["備註"]:

                with st.expander("📌 查看方案備註"):

                    for note in item["備註"]:
                        st.write("• " + note)


# =====================================================
# 9. 初步推薦範圍
# =====================================================

def initial_scope(profile):

    choice = profile.get("旅遊類型", "國內外皆有")

    if choice == "國內為主":
        return "國內"

    if choice == "海外為主":
        return "海外"

    return "全部"


# =====================================================
# 10. 尚未登入：建立個人資料
# =====================================================

if st.session_state.profile is None:

    st.markdown("## 🌼 歡迎加入旅旅！")

    st.write(
        "先讓我們認識你一點點，"
        "之後就能依照你的旅遊習慣幫你推薦旅平險。"
    )

    st.markdown(
        '<div class="pill">👤 基本資料</div>'
        '<div class="pill">🌏 旅遊足跡</div>'
        '<div class="pill">💗 健康狀況</div>',
        unsafe_allow_html=True
    )

    with st.form("profile_form"):

        st.markdown("### ① 認識你")

        col1, col2, col3 = st.columns(3)

        with col1:

            name = st.text_input(
                "姓名 *",
                placeholder="怎麼稱呼你？"
            )

            gender = st.selectbox(
                "性別",
                ["不透露", "女性", "男性", "其他"]
            )

        with col2:

            age = st.number_input(
                "年齡",
                min_value=1,
                max_value=100,
                value=25
            )

            occupation = st.selectbox(
                "職業",
                [
                    "學生",
                    "一般上班族",
                    "服務業",
                    "自營工作者",
                    "自由工作者",
                    "退休",
                    "其他"
                ]
            )

        with col3:

            family_status = st.selectbox(
                "家庭狀況",
                [
                    "不透露",
                    "單身",
                    "已婚",
                    "有子女",
                    "與家人同住"
                ]
            )

            contact = st.text_input(
                "聯絡方式（選填）",
                placeholder="可不填"
            )

        st.divider()

        st.markdown("### ② 你的旅遊足跡")

        col1, col2 = st.columns(2)

        with col1:

            visited = st.text_input(
                "曾經去過哪些國家／地區？",
                placeholder="例如：日本、韓國、泰國"
            )

            frequent = st.text_input(
                "最常去或特別喜歡的國家？",
                placeholder="例如：日本、台灣"
            )

        with col2:

            frequency = st.selectbox(
                "平常旅遊頻率",
                [
                    "很少旅遊",
                    "一年 1～2 次",
                    "一年 3～5 次",
                    "一年 6 次以上"
                ]
            )

            travel_type = st.selectbox(
                "平常以哪種旅遊為主？",
                [
                    "國內外皆有",
                    "海外為主",
                    "國內為主"
                ]
            )

        st.divider()

        st.markdown("### ③ 健康狀況")

        st.caption(
            "健康資料僅在本次工作階段暫存，"
            "不會寫入本機使用者資料檔。"
        )

        health = st.multiselect(
            "是否有以下疾病或健康情形？（可複選）",
            [
                "無特殊疾病",
                "高血壓",
                "糖尿病",
                "高血脂",
                "心臟疾病",
                "氣喘／慢性呼吸道疾病",
                "癌症病史／治療中",
                "腎臟疾病",
                "肝臟疾病",
                "腦中風病史",
                "癲癇",
                "其他疾病或健康情形",
                "不方便透露"
            ]
        )

        health_note = st.text_input(
            "健康狀況補充（選填）",
            placeholder="不想補充可以留白"
        )

        submitted = st.form_submit_button(
            "🌷 完成資料，看看推薦",
            type="primary",
            use_container_width=True
        )

        if submitted:

            if not name.strip():

                st.error("請先填寫姓名喔！")

            else:

                if "不方便透露" in health:
                    health = ["不方便透露"]

                elif "無特殊疾病" in health and len(health) > 1:
                    health = [
                        item for item in health
                        if item != "無特殊疾病"
                    ]

                profile = {

                    "姓名": name.strip(),
                    "性別": gender,
                    "年齡": int(age),
                    "職業": occupation,
                    "家庭狀況": family_status,
                    "聯絡方式": contact,
                    "曾去過的國家": visited,
                    "常去的國家": frequent,
                    "旅遊頻率": frequency,
                    "旅遊類型": travel_type,
                    "健康狀況": health,
                    "健康補充": health_note

                }

                st.session_state.profile = profile

                save_user(profile)

                st.rerun()

    st.stop()


# =====================================================
# 11. 登入後側邊欄
# =====================================================

profile = st.session_state.profile

with st.sidebar:

    st.markdown("## 🐰 旅旅小屋")

    st.markdown(
        f"### 嗨，{profile['姓名']}！"
    )

    st.caption("今天也要平安出發喔 🌷")

    st.divider()

    st.markdown("**你的旅遊小檔案**")

    st.write(f"🎂 {profile['年齡']} 歲")
    st.write(f"🧳 {profile['旅遊頻率']}")
    st.write(f"🌏 {profile['旅遊類型']}")

    st.divider()

    if st.button(
        "✏️ 重新建立個人資料",
        use_container_width=True
    ):

        st.session_state.profile = None
        st.session_state.chat_history = []
        st.session_state.chat_active = False
        st.session_state.trip_requirements = None
        st.session_state.chat_recommendations = None

        st.rerun()

    if st.button(
        "登出",
        use_container_width=True
    ):

        st.session_state.profile = None
        st.session_state.chat_history = []
        st.session_state.chat_active = False

        st.rerun()


# =====================================================
# 12. 首頁
# =====================================================

st.markdown(
    f"## 嗨，{profile['姓名']}！今天想去哪裡玩呢？ 🌈"
)

st.caption(
    "這裡先依照你的個人資料與旅遊習慣，"
    "整理出初步可比較的旅平險方案。"
)

m1, m2, m3, m4 = st.columns(4)

m1.metric(
    "旅遊習慣",
    profile["旅遊類型"]
)

m2.metric(
    "旅遊頻率",
    profile["旅遊頻率"]
)

m3.metric(
    "曾去過國家",
    "已填寫" if profile["曾去過的國家"] else "待補充"
)

m4.metric(
    "可比較方案",
    len(plans)
)

tab_home, tab_chat, tab_profile, tab_data = st.tabs([
    "🏡 我的推薦首頁",
    "💬 旅遊小幫手",
    "👤 個人資料",
    "📚 方案資料庫"
])


# =====================================================
# 13. 主動推薦首頁
# =====================================================

with tab_home:

    st.markdown("### ✨ 為你準備的旅平險推薦")

    st.write(
        "依照你的旅遊類型與資料庫中的保障內容排序。"
        "還沒決定目的地也沒關係，先看看有哪些方案可以比較！"
    )

    scope = initial_scope(profile)

    if scope == "全部":

        st.info(
            "你平常國內外都會旅遊，"
            "因此先同時列出國內與海外方案。"
        )

    else:

        st.info(
            f"依照你的旅遊習慣，"
            f"先優先顯示「{scope}」旅遊方案。"
        )

    initial_results = recommend(
        plans,
        scope_choice=scope
    )

    if data_error:

        st.error(f"資料讀取問題：{data_error}")

        st.caption(f"預期 Excel 位置：{DATA_FILE}")

    elif not initial_results:

        st.warning(
            "目前資料庫沒有讀取到可推薦方案，"
            "請確認 Excel 的保障明細工作表與資料內容。"
        )

    else:

        st.markdown(
            f"#### 🎀 找到 {len(initial_results)} 個可以先比較的方案"
        )

        st.caption(
            "推薦分數只反映資料庫中可辨識的核心保障項目，"
            "不代表保險公司評級、保費高低或一定承保。"
        )

        for index, item in enumerate(
            initial_results[:6],
            1
        ):

            render_plan_card(
                item,
                index,
                compact=(index > 3)
            )

        if len(initial_results) > 6:

            with st.expander(
                f"再看看其他 {len(initial_results)-6} 個方案"
            ):

                for index, item in enumerate(
                    initial_results[6:],
                    7
                ):

                    render_plan_card(
                        item,
                        index,
                        compact=True
                    )

    st.markdown("""
    <div class="soft-card">
        <b>💡 小提醒</b><br>
        這是依個人檔案產生的初步推薦。
        當你有確定的旅遊地點、天數、活動與預算時，
        可以到「旅遊小幫手」補充需求，
        系統再依旅遊情境縮小比較範圍。
    </div>
    """, unsafe_allow_html=True)


# =====================================================
# 14. 對話式需求小幫手
# =====================================================

with tab_chat:

    st.markdown("### 💬 跟旅旅說說你的旅行計畫")

    st.caption(
        "先聊聊你想去哪裡；需要精準推薦時，"
        "再填寫下方的旅遊需求。"
    )

    st.info(
        "目前為課堂展示用的規則式對話示範，"
        "尚未連接真正的 AI API。"
    )

    # 顯示對話紀錄
    for message in st.session_state.chat_history:

        with st.chat_message(message["role"]):

            st.write(message["content"])

    question = st.chat_input(
        "例如：我下個月想去日本玩五天，想找醫療保障完整的方案……"
    )

    if question:

        st.session_state.chat_history.append({
            "role": "user",
            "content": question
        })

        q = question.lower()

        if any(
            word in q
            for word in [
                "日本", "韓國", "泰國",
                "美國", "海外", "出國",
                "東京", "大阪"
            ]
        ):

            answer = (
                "聽起來是一趟海外旅行呢！🌏 "
                "海外旅平險除了意外身故／失能與傷害醫療，"
                "也可以特別比較海外突發疾病醫療。"
                "你可以在下方補充日期、天數和活動，"
                "我再幫你整理方案。"
            )

        elif any(
            word in q
            for word in [
                "國內", "台灣", "花蓮",
                "台東", "墾丁", "宜蘭"
            ]
        ):

            answer = (
                "國內小旅行也要做好保障準備！🧳 "
                "可以在下方填寫旅遊天數、活動，"
                "以及想優先比較的保障項目。"
            )

        elif any(
            word in q
            for word in [
                "疾病", "高血壓",
                "糖尿病", "氣喘", "心臟"
            ]
        ):

            answer = (
                "健康狀況可能涉及保險公司的健康告知、"
                "既往症及承保規定。請依實際要保書誠實告知，"
                "並向保險公司確認；系統不會自行判定能否投保。"
            )

        elif any(
            word in q
            for word in [
                "保障", "醫療", "意外", "救援"
            ]
        ):

            answer = (
                "收到！可以告訴我你最在意意外身故／失能、"
                "傷害醫療，還是海外突發疾病醫療？"
                "下方也能勾選多項保障偏好。"
            )

        else:

            answer = (
                "收到你的旅行想法囉！🌷 "
                "如果想進一步比較旅平險，"
                "可以展開下方的「填寫本次旅遊需求」，"
                "我會依照你提供的情境整理方案。"
            )

        st.session_state.chat_history.append({
            "role": "assistant",
            "content": answer
        })

        st.session_state.chat_active = True

        st.rerun()

    # 有開始對話才顯示需求表單
    if st.session_state.chat_history:
        st.session_state.chat_active = True

    if st.session_state.chat_active:

        with st.expander(
            "📝 填寫本次旅遊需求，取得更精準推薦",
            expanded=True
        ):

            with st.form("trip_form"):

                col1, col2, col3 = st.columns(3)

                with col1:

                    trip_scope = st.selectbox(
                        "旅遊地區",
                        ["海外旅遊", "國內旅遊"]
                    )

                    destination = st.text_input(
                        "目的地／國家",
                        placeholder="例如：日本東京、花蓮"
                    )

                with col2:

                    departure = st.date_input(
                        "預計出發日期",
                        value=date.today() + timedelta(days=30)
                    )

                    days = st.number_input(
                        "旅遊天數",
                        min_value=1,
                        max_value=365,
                        value=5
                    )

                with col3:

                    people = st.number_input(
                        "同行人數（含自己）",
                        min_value=1,
                        max_value=30,
                        value=1
                    )

                    purpose = st.selectbox(
                        "旅遊目的",
                        [
                            "觀光休閒",
                            "探親訪友",
                            "商務出差",
                            "遊學／交換",
                            "其他"
                        ]
                    )

                activity = st.multiselect(
                    "這趟旅程可能從事的活動",
                    [
                        "一般觀光",
                        "自駕／租車",
                        "登山健行",
                        "滑雪",
                        "潛水／水上活動",
                        "自行車",
                        "露營",
                        "其他"
                    ],
                    default=["一般觀光"]
                )

                budget = st.selectbox(
                    "每人保費預算",
                    [
                        "尚未決定",
                        "500 元以下",
                        "500～1,000 元",
                        "1,001～2,000 元",
                        "2,000 元以上"
                    ]
                )

                needs = st.multiselect(
                    "想優先比較的保障",
                    [
                        "意外身故／失能",
                        "傷害醫療",
                        "海外突發疾病醫療",
                        "緊急救援"
                    ],
                    default=[
                        "意外身故／失能",
                        "傷害醫療"
                    ]
                )

                extra = st.text_area(
                    "其他希望小幫手留意的事情（選填）",
                    placeholder="例如：和長輩同行、希望醫療保障高一點……"
                )

                submitted = st.form_submit_button(
                    "✨ 根據這趟旅行幫我推薦",
                    type="primary",
                    use_container_width=True
                )

                if submitted:

                    target = (
                        "海外"
                        if trip_scope == "海外旅遊"
                        else "國內"
                    )

                    result = recommend(
                        plans,
                        scope_choice=target,
                        needs=needs
                    )

                    st.session_state.trip_requirements = {

                        "旅遊地區": trip_scope,
                        "目的地": destination,
                        "出發日期": str(departure),
                        "天數": int(days),
                        "同行人數": int(people),
                        "旅遊目的": purpose,
                        "活動": activity,
                        "預算": budget,
                        "保障需求": needs,
                        "其他需求": extra

                    }

                    st.session_state.chat_recommendations = result

                    st.rerun()

    # 顯示需求摘要與推薦
    if st.session_state.trip_requirements:

        req = st.session_state.trip_requirements

        st.markdown("### 🧡 這趟旅行的需求摘要")

        st.write(
            f"📍 **目的地：** {req['目的地'] or req['旅遊地區']} "
            f"｜ 🗓️ **出發：** {req['出發日期']} "
            f"｜ **天數：** {req['天數']} 天"
        )

        st.write(
            f"👥 **同行：** {req['同行人數']} 人 "
            f"｜ 🎒 **活動：** {'、'.join(req['活動']) or '未指定'}"
        )

        st.write(
            f"🛡️ **保障偏好：** {'、'.join(req['保障需求']) or '未指定'}"
        )

        st.caption(
            f"預算：{req['預算']}。"
            "目前資料庫未完整收錄各方案保費，"
            "因此預算僅作需求紀錄，尚未進行價格篩選。"
        )

        results = st.session_state.chat_recommendations or []

        st.markdown(
            f"### 🎁 依本次旅程整理的方案（{len(results)} 個）"
        )

        if results:

            for index, item in enumerate(results[:8], 1):

                render_plan_card(
                    item,
                    index,
                    compact=(index > 4)
                )

        else:

            st.warning(
                "目前資料庫沒有符合此旅遊範圍的方案資料，"
                "請確認 Excel 保障明細。"
            )

        st.caption(
            "提醒：推薦結果僅供課堂展示與初步比較，"
            "實際承保、保費及保障內容以保險公司最新條款為準。"
        )

    if st.button("🧹 清除對話紀錄"):

        st.session_state.chat_history = []
        st.session_state.chat_active = False
        st.session_state.trip_requirements = None
        st.session_state.chat_recommendations = None

        st.rerun()


# =====================================================
# 15. 個人資料
# =====================================================

with tab_profile:

    st.markdown("### 👤 我的個人旅遊檔案")

    left, right = st.columns(2)

    with left:

        with st.container(border=True):

            st.markdown("#### 🌷 基本資料")

            st.write(f"**姓名：** {profile['姓名']}")
            st.write(f"**性別：** {profile['性別']}")
            st.write(f"**年齡：** {profile['年齡']}")
            st.write(f"**職業：** {profile['職業']}")
            st.write(f"**家庭狀況：** {profile['家庭狀況']}")
            st.write(f"**聯絡方式：** {profile['聯絡方式'] or '未填寫'}")

    with right:

        with st.container(border=True):

            st.markdown("#### 🌏 旅遊足跡")

            st.write(
                f"**曾去過：** {profile['曾去過的國家'] or '未填寫'}"
            )

            st.write(
                f"**常去國家：** {profile['常去的國家'] or '未填寫'}"
            )

            st.write(
                f"**旅遊頻率：** {profile['旅遊頻率']}"
            )

            st.write(
                f"**旅遊類型：** {profile['旅遊類型']}"
            )

    with st.expander("💗 查看健康狀況"):

        st.write(
            "、".join(
                profile.get("健康狀況", [])
            ) or "未填寫"
        )

        if profile.get("健康補充"):
            st.write(profile["健康補充"])

        st.caption(
            "健康資料只在目前工作階段暫存，"
            "不會寫入 users.json。"
            "實際投保仍須依保險公司健康告知及條款辦理。"
        )


# =====================================================
# 16. 保險方案資料庫
# =====================================================

with tab_data:

    st.markdown("### 📚 旅遊平安險方案資料庫")

    if data_error:

        st.error(f"資料讀取失敗：{data_error}")

        st.caption(f"預期 Excel 位置：{DATA_FILE}")

    elif df.empty:

        st.warning("目前 Excel 尚無保障明細。")

    else:

        st.success(
            f"已讀取 {len(df)} 筆保障資料，"
            f"整理出 {len(plans)} 個可比較方案。"
        )

        companies = (
            ["全部"] +
            sorted(
                df["公司"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
        ) if "公司" in df.columns else ["全部"]

        col1, col2 = st.columns(2)

        with col1:

            selected_company = st.selectbox(
                "保險公司",
                companies
            )

        with col2:

            scope_filter = st.selectbox(
                "旅遊範圍",
                ["全部", "海外", "國內"]
            )

        display_df = df.copy()

        if selected_company != "全部" and "公司" in display_df.columns:

            display_df = display_df[
                display_df["公司"].astype(str)
                == selected_company
            ]

        if scope_filter != "全部" and "旅遊範圍" in display_df.columns:

            display_df = display_df[
                display_df["旅遊範圍"]
                .astype(str)
                .str.contains(scope_filter, na=False)
            ]

        if "系統分類" in display_df.columns:

            display_df = display_df[
                display_df["系統分類"].astype(str)
                != "排除"
            ]

        show_columns = [

            column

            for column in [

                "公司",
                "旅遊範圍",
                "商品/方案",
                "系統分類",
                "標準化保障項目",
                "保障額度/文字",
                "給付方式",
                "重要限制或備註",
                "核對狀態"

            ]

            if column in display_df.columns

        ]

        st.dataframe(
            display_df[show_columns],
            use_container_width=True,
            hide_index=True
        )

        with st.expander("📖 資料庫與推薦方式說明"):

            st.write(
                "1. 優先讀取保障明細_整理版，"
                "若不存在才讀取保障明細。"
            )

            st.write(
                "2. 系統會統一 Excel 全形與半形斜線。"
            )

            st.write(
                "3. 只比較旅遊平安險相關保障，"
                "不納入旅遊不便險。"
            )

            st.write(
                "4. 推薦分數依資料庫中可辨識的核心保障項目計算，"
                "不代表官方商品評級。"
            )


# =====================================================
# 17. 頁尾
# =====================================================

st.divider()

st.markdown("""
<div style="text-align:center; color:#9a8b9b; padding:10px;">
    🌷 旅旅・智慧旅遊平安險推薦系統 🌷<br>
    <span style="font-size:.8rem;">
        課堂專題展示版本｜本系統僅供資訊比較，不構成正式投保建議。
        實際承保、保費、健康告知及理賠仍以保險公司最新要保書與保單條款為準。
    </span>
</div>
""", unsafe_allow_html=True)
