import pandas as pd
import random
import re
from datetime import datetime, timedelta
import os
import sys

WORD_FILE = "日本語単語.xlsx"
PROGRESS_FILE = "progress.xlsx"
WRONG_FILE = "wrong_words.xlsx"


def configure_console_utf8():
    if os.name == "nt":
        try:
            os.system("chcp 65001 >nul")
        except Exception:
            pass

    for stream in (sys.stdin, sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


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
    random.shuffle(new_words)
    return new_words[:limit]


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
        # 新学习：从题库中取前 50 个未学习的词
        today_words = get_new_words(words, progress, limit=50)
        print(f"新学习模式：本次 {len(today_words)} 个新词")
    else:
        # 复习：从已学习且到期（<=今天）的词中取最多 50 个
        today_words = get_today_words(words, progress)
        # 优先复习更早到期的；同一天到期时，错得多的优先
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
            today_words = ordered[:50]
        else:
            today_words = today_words[:50]
        print(f"复习模式：本次 {len(today_words)} 个到期词")

    if not today_words:
        if session_mode == "1":
            print("没有未学习的新单词了！")
        else:
            print("今天没有需要复习的单词！")
        return

    today_words = today_words[:]
    if session_mode != "1":
        random.shuffle(today_words)

    total = len(today_words)
    correct_count = 0

    def normalize(s: str) -> str:
        s = s.strip().lower()
        s = re.sub(r"\s+", "", s)
        s = re.sub(r"[，,。．\.、;；/／\(\)（）\[\]【】「」『』《》<>]", "", s)
        return s

    def split_answers(s: str) -> list[str]:
        parts = re.split(r"[;；/／、，,]|或", str(s))
        return [normalize(p) for p in parts if normalize(p)]

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

        if direction_mode == "1":
            prompt = f"[{i}/{total}] {word['jp']}{pos_label} -> "
            user_ans = input(prompt)
            expected = split_answers(word["cn"])
            ok = normalize(user_ans) in expected if expected else (normalize(user_ans) == normalize(word["cn"]))
        else:
            prompt = f"[{i}/{total}] {word['cn']}{pos_label} -> "
            user_ans = input(prompt)
            expected = split_answers(word["jp"])
            ok = normalize(user_ans) in expected if expected else (normalize(user_ans) == normalize(word["jp"]))

        if ok:
            print("正确")
            correct_count += 1
        else:
            print("错误")
            print(f"正确答案：{word['jp']} ｜ {word['cn']}")
            if word.get("example"):
                print(f"例句：{word['example']}")
            save_wrong(word)

        progress.loc[row_idx] = update_interval(progress.loc[row_idx], ok)

    save_progress(progress)
    print(f"完成：{correct_count}/{total} 正确")

# ========= 主程序 =========
def main():
    configure_console_utf8()

    if not os.path.exists(WORD_FILE):
        print(f"找不到单词文件：{WORD_FILE}")
        return

    words = load_words()
    progress, is_first_run = load_progress(words)

    print("日语单词出题器（艾宾浩斯版）")
    print("出题方向：")
    print("1 日译中")
    print("2 中译日")

    direction_mode = input("请选择出题方向 (1/2)：").strip()

    if direction_mode not in ["1", "2"]:
        print("出题方向错误")
        return

    print("学习类型：")
    print("1 新学习（从题库挑 50 个没学过的）")
    print("2 复习（挑已学习且到期的，最多 50 个）")

    session_mode = input("请选择学习类型 (1/2)：").strip()
    if session_mode not in ["1", "2"]:
        print("学习类型错误")
        return

    quiz(words, progress, direction_mode, session_mode, is_first_run)


if __name__ == "__main__":
    main()