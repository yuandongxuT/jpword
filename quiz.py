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

    for word in words:
        if word["key"] not in progress["key"].values:
            progress.loc[len(progress)] = [word["key"], 0, datetime.now().date(), 0, 0]

    return progress, is_first_run


# ========= 保存进度 =========
def save_progress(progress):
    progress.to_excel(PROGRESS_FILE, index=False)


# ========= 错题本 =========
def save_wrong(word):
    if os.path.exists(WRONG_FILE):
        wrong_df = pd.read_excel(WRONG_FILE)
    else:
        wrong_df = pd.DataFrame(columns=["日语", "中文", "例句", "时间"])

    wrong_df.loc[len(wrong_df)] = [
        word["jp"], word["cn"], word["example"], datetime.now()
    ]

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


# ========= 出题 =========
def quiz(words, progress, mode, is_first_run):

    if is_first_run:
        print("第一次运行：全量学习模式")
        today_words = words
    else:
        today_words = get_today_words(words, progress)

    if not today_words:
        print("今天没有需要复习的单词！")
        return

    word_by_key = {w["key"]: w for w in words}
    today_words = today_words[:]
    random.shuffle(today_words)

    today_words = today_words[:50]
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
        idxs = progress.index[progress["key"] == key].tolist()
        if not idxs:
            continue
        row_idx = idxs[0]

        pos = word.get("pos") or ""
        pos_label = f" ({pos})" if pos else ""

        if mode == "1":
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
    print("1 日译中")
    print("2 中译日")

    mode = input("请选择模式 (1/2)：").strip()

    if mode not in ["1", "2"]:
        print("模式错误")
        return

    quiz(words, progress, mode, is_first_run)


if __name__ == "__main__":
    main()