import pandas as pd
import random
import re
from datetime import datetime, timedelta
import os
import sys
import requests

try:
    import certifi  # type: ignore
except Exception:
    certifi = None

WORD_FILE = "日本語単語.xlsx"
PROGRESS_FILE = "progress.xlsx"
WRONG_FILE = "wrong_words.xlsx"

# DeepSeek 配置（默认按 DeepSeek OpenAI 兼容接口来调用）
LLM_API_KEY = (
    os.getenv("LLM_API_KEY")
    or os.getenv("DEEPSEEK_API_KEY")
    or os.getenv("OPENAI_API_KEY")
)

LLM_API_URL = os.getenv(
    "LLM_API_URL",
    # DeepSeek / OpenAI 兼容接口一般是 /v1/chat/completions
    "https://api.deepseek.com/v1/chat/completions",
)
LLM_MODEL = os.getenv("LLM_MODEL", "deepseek-chat")

# SSL/TLS 证书相关配置：
# - LLM_SSL_VERIFY=0 可临时关闭证书校验（不推荐，只用于排障）
# - LLM_CA_BUNDLE=证书路径 可指定自定义 CA（公司代理/抓包证书场景常用）
LLM_SSL_VERIFY = os.getenv("LLM_SSL_VERIFY", "1").strip().lower() not in ("0", "false", "no")
LLM_CA_BUNDLE = os.getenv("LLM_CA_BUNDLE") or os.getenv("REQUESTS_CA_BUNDLE")


def configure_console_utf8():
    if os.name == "nt":
        try:
            os.system("chcp 65001 >nul")
        except Exception:
            pass

    # 确保控制台输入输出使用 UTF-8，避免日文乱码
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def get_memory_tip_from_llm(word: dict) -> str | None:
    """
    调用大语言模型，为错误的单词生成记忆方法。

    环境变量配置：
      - LLM_API_KEY / DEEPSEEK_API_KEY / OPENAI_API_KEY：API 密钥（必填其一，否则不调用）
      - LLM_API_URL：接口地址（默认 DeepSeek 的 chat/completions）
      - LLM_MODEL：模型名（默认 deepseek-chat，可自行修改）
    """
    if not LLM_API_KEY:
        print("[LLM] 未设置 DeepSeek/OpenAI API Key，已跳过记忆方法。")
        return None

    jp = str(word.get("jp", ""))
    cn = str(word.get("cn", ""))
    example = str(word.get("example", "") or "")

    system_prompt = (
        "你是一名日语单词记忆教练，请用简体中文、结合联想记忆、词源、构词法、谐音等方式，"
        "为学习者设计该日语单词的记忆方法，语言要简洁、具体、便于快速记住。"
    )

    user_prompt = (
        f"日语单词：{jp}\n"
        f"中文含义：{cn}\n"
        f"例句（可选）：{example or '无'}\n\n"
        "请给出 1~2 条记忆方法，可以包括：\n"
        "1. 简短的谐音联想\n"
        "2. 与汉字含义或词源的联系\n"
        "3. 在具体场景中的画面化记忆\n"
        "回答时只给出记忆方法本身，不要解释你是 AI。"
    )

    headers = {
        "Authorization": f"Bearer {LLM_API_KEY}",
        "Content-Type": "application/json",
    }

    body = {
        "model": LLM_MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.7,
    }

    verify_opt: bool | str = True
    if not LLM_SSL_VERIFY:
        verify_opt = False
    elif LLM_CA_BUNDLE:
        verify_opt = LLM_CA_BUNDLE
    elif certifi is not None:
        # 在部分 Windows 环境里，显式指定 certifi 更稳定
        verify_opt = certifi.where()

    if verify_opt is False:
        print("[LLM] 警告：当前关闭了 HTTPS 证书校验（LLM_SSL_VERIFY=0），不安全且会触发 InsecureRequestWarning。")

    try:
        resp = requests.post(
            LLM_API_URL,
            headers=headers,
            json=body,
            timeout=20,
            verify=verify_opt,
        )

        # 针对常见状态码给出可操作的提示
        if resp.status_code == 401:
            print("[LLM] 401 未授权：请检查 LLM_API_KEY/DEEPSEEK_API_KEY 是否正确。")
            return None
        if resp.status_code == 402:
            print("[LLM] 402 Payment Required：账号余额/套餐不足或未开通，请在 DeepSeek 控制台充值/开通后再试。")
            return None
        if resp.status_code == 429:
            print("[LLM] 429 请求过多：触发限流了，稍后再试。")
            return None

        resp.raise_for_status()

        data = resp.json()

        # 兼容 OpenAI / DeepSeek chat-completions 风格的响应结构
        choices = data.get("choices")
        if not choices:
            print("[LLM] 响应中没有 choices 字段。")
            return None

        message = choices[0].get("message") or {}
        content = message.get("content")
        if isinstance(content, str):
            return content.strip()

        # 兜底：部分兼容接口可能直接返回 text
        text = choices[0].get("text")
        if isinstance(text, str):
            return text.strip()

        print("[LLM] 未能从返回结果中解析出文本内容。")

    except requests.exceptions.SSLError as e:
        print(f"[LLM] SSL 证书校验失败：{e}")
        if LLM_CA_BUNDLE:
            print(f"[LLM] 当前使用的 CA Bundle：{LLM_CA_BUNDLE}")
        else:
            print("[LLM] 可尝试：升级 certifi / 配置 REQUESTS_CA_BUNDLE 或 LLM_CA_BUNDLE（公司代理场景）。")
    except requests.exceptions.RequestException as e:
        print(f"[LLM] 网络请求失败：{e}")
    except ValueError as e:
        # JSON 解析失败等
        print(f"[LLM] 响应解析失败：{e}")
    except Exception as e:
        print(f"[LLM] 调用大模型失败：{e}")

    return None


# ========= 解析单词 =========
def load_words():
    df = pd.read_excel(WORD_FILE, header=None)

    col = df[0].fillna("").astype(str)
    words = []
    current_pos = ""

    i = 0
    n = len(col)

    while i < n:
        text = col.iloc[i].strip()

        if not text:
            i += 1
            continue

        # 识别「第一部分：名词」这类分段，记录当前词性
        m = re.match(r"^第.+部分[:：]\s*(.+)$", text)
        if m:
            current_pos = m.group(1).strip()
            i += 1
            continue

        # 识别真正的单词行：形如「1. 日语｜中文」
        word_line = re.sub(r"^\d+\.\s*", "", text)

        if "｜" in word_line:
            jp, cn = word_line.split("｜", 1)

            # 例句在下一行，如果那一行不是新的分段、也不是下一条单词，则视为例句
            example = ""
            if i + 1 < n:
                next_text = col.iloc[i + 1].strip()
                if (
                    next_text
                    and "｜" not in next_text
                    and not re.match(r"^第.+部分", next_text)
                ):
                    example = next_text
                    i += 1  # 额外跳过例句行

            words.append({
                "jp": jp.strip(),
                "cn": cn.strip(),
                "example": example,
                "key": jp.strip(),
                "pos": current_pos,
            })

        i += 1

    return words


# ========= 加载学习进度 =========
def load_progress(words):
    is_first_run = not os.path.exists(PROGRESS_FILE)

    if not is_first_run:
        progress = pd.read_excel(PROGRESS_FILE)
        progress["next_date"] = pd.to_datetime(progress["next_date"]).dt.date
    else:
        progress = pd.DataFrame(columns=["key", "interval", "next_date", "correct", "wrong"])

    return progress, is_first_run


# ========= 学习进度统计 =========
def show_statistics(words, progress):
    total_words = len(words)

    if progress.empty:
        learned = 0
        due_today = 0
    else:
        learned = progress["key"].nunique()
        today = datetime.now().date()
        due_today = (progress["next_date"] <= today).sum()

    new_words = total_words - learned
    percent = (learned / total_words * 100) if total_words else 0

    print("\n📊 学习进度统计")
    print(f"总单词数：{total_words}")
    print(f"已学习：{learned}")
    print(f"未学习：{new_words}")
    print(f"今日待复习：{due_today}")
    print(f"学习进度：{percent:.2f}%")
    print("-" * 30)
        
# ========= 保存进度 =========
def save_progress(progress):
    progress.to_excel(PROGRESS_FILE, index=False)


# ========= 错题本 =========
def save_wrong(word):
    if os.path.exists(WRONG_FILE):
        wrong_df = pd.read_excel(WRONG_FILE)
    else:
        wrong_df = pd.DataFrame(columns=["日语", "中文", "例句", "时间", "错误次数"])

    # 确保有“错误次数”这一列
    if "错误次数" not in wrong_df.columns:
        wrong_df["错误次数"] = 0

    jp = word["jp"]
    cn = word["cn"]
    example = word.get("example", "")
    now = datetime.now()

    # 查找是否已有同一单词（按 日语+中文 识别重复）
    mask = (wrong_df["日语"] == jp) & (wrong_df["中文"] == cn)

    if mask.any():
        # 已存在：错误次数 +1，更新时间
        idx = wrong_df.index[mask][0]
        wrong_df.at[idx, "错误次数"] = (
            pd.to_numeric(wrong_df.at[idx, "错误次数"], errors="coerce") if "错误次数" in wrong_df.columns else 0
        )
        wrong_df.at[idx, "错误次数"] = (wrong_df.at[idx, "错误次数"] or 0) + 1
        wrong_df.at[idx, "时间"] = now
        # 如需更新例句，可取消下一行注释
        # wrong_df.at[idx, "例句"] = example
    else:
        # 新单词：从 1 次开始记录
        wrong_df.loc[len(wrong_df)] = [jp, cn, example, now, 1]

    wrong_df.to_excel(WRONG_FILE, index=False)


# ========= 更新艾宾浩斯间隔 =========
def update_interval(row, correct):
    intervals = [1, 3, 7, 15, 30]

    if correct:
        idx = int(row["interval"])
        if idx < len(intervals):
            row["next_date"] = datetime.now().date() + timedelta(days=intervals[idx])
            row["interval"] = idx + 1
        row["correct"] += 1
    else:
        row["interval"] = 0
        row["next_date"] = datetime.now().date()
        row["wrong"] += 1

    return row


# ========= 获取今日要复习的词 =========
def get_today_words(words, progress):
    today = datetime.now().date()

    due_keys = progress[progress["next_date"] <= today]["key"].tolist()

    return [w for w in words if w["key"] in due_keys]


# ========= 选择新学习单词 =========
def get_new_words(words, progress, limit=50):
    learned_keys = set(progress["key"].astype(str).tolist()) if not progress.empty else set()
    new_words = [w for w in words if w["key"] not in learned_keys]

    # 优先选择「常句」部分的内容
    priority = [w for w in new_words if "常句" in str(w.get("pos", ""))]
    others = [w for w in new_words if w not in priority]

    random.shuffle(priority)
    random.shuffle(others)

    selected = priority[:limit]
    if len(selected) < limit:
        need = limit - len(selected)
        selected.extend(others[:need])

    return selected


# ========= 确保进度行存在（新学习会用到） =========
def ensure_progress_row(progress, key: str):
    idxs = progress.index[progress["key"] == key].tolist() if not progress.empty else []
    if idxs:
        return idxs[0], progress

    progress.loc[len(progress)] = [key, 0, datetime.now().date(), 0, 0]
    return progress.index[-1], progress


# ========= 出题 =========
def quiz(words, progress, direction_mode, session_mode, is_first_run):
    word_by_key = {w["key"]: w for w in words}

    if session_mode == "1":
        today_words = get_new_words(words, progress, limit=50)
        print(f"新学习模式：本次 {len(today_words)} 个新词")
    else:
        today_words = get_today_words(words, progress)

        if not progress.empty and today_words:
            due = progress[progress["next_date"] <= datetime.now().date()].copy()
            due["wrong"] = pd.to_numeric(due.get("wrong", 0), errors="coerce").fillna(0).astype(int)
            due = due.sort_values(["next_date", "wrong"], ascending=[True, False])

            ordered = []
            seen = set()
            for k in due["key"].astype(str).tolist():
                if k in word_by_key and k not in seen:
                    ordered.append(word_by_key[k])
                    seen.add(k)

            # 在到期词中优先「常句」部分
            priority = [w for w in ordered if "常句" in str(w.get("pos", ""))]
            others = [w for w in ordered if w not in priority]
            today_words = (priority + others)[:50]
        else:
            today_words = today_words[:50]

        print(f"复习模式：本次 {len(today_words)} 个到期词")

    if not today_words:
        print("没有需要学习的单词！")
        return

    if session_mode != "1":
        random.shuffle(today_words)

    total = len(today_words)
    correct_count = 0

    def normalize(s: str) -> str:
        s = str(s)
        # 常见格式：素敵な（すてきな） / 素敵(すてき) —— 忽略括号里的读音，避免影响判定
        s = re.sub(r"[（(][^）)]*[）)]", "", s)
        s = s.strip().lower()
        s = re.sub(r"\s+", "", s)
        s = re.sub(r"[，,。．\.、;；/／\(\)（）\[\]【】「」『』《》<>]", "", s)
        return s

    def split_answers(s: str) -> list[str]:
        parts = re.split(r"[;；/／、，,]|或", str(s))
        return [normalize(p) for p in parts if normalize(p)]

    def is_na_adjective(pos_text: str, cn_text: str = "") -> bool:
        p = str(pos_text or "")
        c = str(cn_text or "")
        return (
            ("な形容" in p)
            or ("形容動詞" in p)
            or ("ナ形容" in p)
            or ("な形容" in c)
            or ("形容動詞" in c)
            or ("ナ形容" in c)
        )

    def acceptable_jp_answers(expected_norm_list: list[str], pos_text: str, cn_text: str = "") -> set[str]:
        acc = set(expected_norm_list or [])
        if is_na_adjective(pos_text, cn_text):
            expanded: set[str] = set()
            for e in list(acc):
                if not e:
                    continue
                if e.endswith("な"):
                    expanded.add(e[:-1])
                else:
                    expanded.add(e + "な")
            acc |= expanded
        return acc

    # 🔹 按词性分组，用于生成选择题
    pos_dict = {}
    for w in words:
        pos = w.get("pos", "")
        if pos not in pos_dict:
            pos_dict[pos] = []
        pos_dict[pos].append(w)

    for i, word in enumerate(today_words, start=1):
        key = word["key"]

        if session_mode == "1":
            row_idx, progress = ensure_progress_row(progress, key)
        else:
            idxs = progress.index[progress["key"] == key].tolist()
            if not idxs:
                continue
            row_idx = idxs[0]

        pos = word.get("pos") or ""
        pos_label = f" ({pos})" if pos else ""

        # =========================
        # 🎯 接続詞 / 副詞 → 选择题（中译日）
        # =========================
        if pos in ["接続詞", "接续词", "副詞", "副词"]:
            same_pos_words = pos_dict.get(pos, [])
            candidates = [w for w in same_pos_words if w["key"] != word["key"]]

            # 至少要有 3 个干扰项，才能凑满 4 选 1；否则退化为输入题
            if len(candidates) >= 3:
                options = random.sample(candidates, 3) + [word]
                random.shuffle(options)

                correct_index = options.index(word) + 1

                print(f"\n[{i}/{total}] {word['cn']}{pos_label} 对应哪个日语？")
                for idx, opt in enumerate(options, 1):
                    print(f"{idx}. {opt['jp']}")

                user_ans = input("请选择 (1-4)：").strip()

                if user_ans == str(correct_index):
                    print("正确")
                    ok = True
                    correct_count += 1
                else:
                    print("错误")
                    print(f"正确答案：{correct_index}")
                    print(f"{word['jp']} ｜ {word['cn']}")
                    if word.get("example"):
                        print(f"例句：{word['example']}")
                    save_wrong(word)
                    tip = get_memory_tip_from_llm(word)
                    if tip:
                        print("\n记忆方法（大模型建议）：")
                        print(tip)
                    ok = False
            else:
                prompt = f"[{i}/{total}] {word['cn']}{pos_label} -> "
                user_ans = input(prompt)

                expected_norm = split_answers(word["jp"])
                if not expected_norm:
                    expected_norm = [normalize(word["jp"])]
                ok = normalize(user_ans) in acceptable_jp_answers(expected_norm, pos, word.get("cn", ""))

                if ok:
                    print("正确")
                    correct_count += 1
                else:
                    print("错误")
                    print(f"正确答案：{word['jp']} ｜ {word['cn']}")
                    if word.get("example"):
                        print(f"例句：{word['example']}")
                    save_wrong(word)
                    tip = get_memory_tip_from_llm(word)
                    if tip:
                        print("\n记忆方法（大模型建议）：")
                        print(tip)

        # =========================
        # 📝 普通词 → 中译日输入
        # =========================
        else:
            prompt = f"[{i}/{total}] {word['cn']}{pos_label} -> "
            user_ans = input(prompt)

            expected_norm = split_answers(word["jp"])
            if not expected_norm:
                expected_norm = [normalize(word["jp"])]
            ok = normalize(user_ans) in acceptable_jp_answers(expected_norm, pos, word.get("cn", ""))

            if ok:
                print("正确")
                correct_count += 1
            else:
                print("错误")
                print(f"正确答案：{word['jp']} ｜ {word['cn']}")
                if word.get("example"):
                    print(f"例句：{word['example']}")
                save_wrong(word)
                tip = get_memory_tip_from_llm(word)
                if tip:
                    print("\n记忆方法（大模型建议）：")
                    print(tip)

        progress.loc[row_idx] = update_interval(progress.loc[row_idx], ok)

    save_progress(progress)
    print(f"\n完成：{correct_count}/{total} 正确")

# ========= 主程序 =========
def main():
    configure_console_utf8()

    if not os.path.exists(WORD_FILE):
        print(f"找不到单词文件：{WORD_FILE}")
        return

    words = load_words()
    progress, is_first_run = load_progress(words)

    show_statistics(words, progress)

    print("日语单词出题器（艾宾浩斯版）")

    print("学习类型：")
    print("1 新学习（从题库挑 50 个没学过的）")
    print("2 复习（挑已学习且到期的，最多 50 个）")

    session_mode = input("请选择学习类型 (1/2)：").strip()
    if session_mode not in ["1", "2"]:
        print("学习类型错误")
        return

    direction_mode = "2"  # 固定中译日

    quiz(words, progress, direction_mode, session_mode, is_first_run)


if __name__ == "__main__":
    main()