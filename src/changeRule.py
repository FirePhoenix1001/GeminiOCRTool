import os
from gemini import changeRule

def load_rule(path):
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        return r"""根據提供的圖片，完全依照原本內容轉成純粹文字，所有內容不考慮markdown語法，並忽略最下方標註的頁碼並遵循以下規則:
1) $...$ 的使用與原則
1.1) 僅在需要**分數**或**根號**時使用 `$...$`，分別用 `\frac{…}{…}`、`\sqrt{…}`。
1.2) **一旦進入 `$...$`，其中內容一律採 LaTeX 寫法**（如 `\times`、`\cdot` 等），不得混入純文字符號。 
1.3) 比較／關係符號（新增）**：`<, >, ≤, ≥, =` 盡量以**純文字**書寫並置於 **`$` 外**；可用多個 `$...$` 片段把符號夾出。  例：`$\overline{AB}$>$\overline{AC}$`、`$a$≤$b$`、`$x$=$y$`。在 `$...$` 中不要使用 \lt, \gt, \le, \ge 等。
2) 純文字表達 - 乘除與省略號：用 `×`、`·`、`/`、`…`（或 `...`）。
- 箭頭/因果：用 `⇒`、`∵`、`∴`。
- 對數以 `log` 書寫。- cos, sin, tan之類的函數不要包進$...$ 
3) 幾何與其他- 需用上劃線者可寫在 `$...$`（如 `$\overline{AB}$`)；其餘能純文字的皆用純文字。
4) 排版-保持原段落與行序；題號與內容格式依題面；結尾「選( )」照原樣。
最後如果這一道題目是有輔助圖片的，請幫我換行註記 *********[缺圖]*********"""

    with open(path, "r", encoding="utf-8") as f:
        return f.read().strip()

def split_text_by_blank_line(path):
    with open(path, "r", encoding="utf-8") as f:
        text = f.read().strip()
    # 依空白行分割
    parts = text.split("\n\n")
    # 如果格式固定，一定會得到兩段
    if len(parts) >= 2:
        upper = parts[0].strip()
        lower = parts[1].strip()
    else:
        # 如果使用者不小心給沒有空白行的文件
        upper = text
        lower = ""
    return upper, lower

rule_text = load_rule("src\\rule.txt")
upper, lower = split_text_by_blank_line("src\\changeRule.txt")
prompt = f"""
你是一個「規則更新器」。

你的任務：
1. 根據 AI 錯誤輸出 upper
2. 根據使用者期望的正確輸出 lower
3. 根據舊規則 rule_text
輸出「更新後的新規則」。

要求：
- 只能輸出新規則本身
- 不要解釋
- 不要加入其它內容

------------------------------------
【AI 錯誤輸出 upper】
{upper}

------------------------------------
【使用者期望輸出 lower】
{lower}

------------------------------------
【舊規則 rule_text】
{rule_text}

------------------------------------
請輸出：【更新後的新規則】
"""


response=str(changeRule(prompt))
print(response)

with open(f"src\\rule.txt", "w", encoding="utf-8") as f:
    f.write(response)
