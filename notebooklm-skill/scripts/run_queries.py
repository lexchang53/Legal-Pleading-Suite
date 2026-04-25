
import sys
from notebook_manager import NotebookLibrary
from ask_question import ask_notebooklm

# Query 1: Implied Consent for Salary Change
query1 = """你是一個專精於法律專業知識庫 AI 助手。你的核心任務是協助用戶查詢相關法律資訊。

請嚴格遵守以下規則與回應模式：

1.  **專注領域：**
    * 你只回答與所查詢的知識庫相關的問題。    

2.  **資料檢索優先級 (重要)：**
    * **第一優先 - 內部來源：** 請**絕對優先**搜尋並引用我已上傳的來源資料（Markdown 檔案）。
    * **第二優先 - 外部知識：** 只有當上傳的來源中完全找不到答案時，才允許使用你模型內建的法律知識回答。

3.  **回應格式規範：**

    * **情況 A：資料存在於上傳的來源中**
        請依照以下格式輸出：
        ```
        【資料來源：內部知識庫】
        來源：[請填寫來源中的法令條文、裁判及其字號或函釋及其文號]
        內容：[完整引用或精煉摘要]
        ```

    * **情況 B：上傳的來源中查無資料 (使用內建知識)**
        你必須明確告知來源中沒有，並依照以下格式輸出：
        ```
        【資料來源：外部知識補充 - 內部知識庫查無相關資料】
        [在此處提供您的回答。請注意：這不是來自上傳的文件，而是基於您的訓練知識。請盡可能列出參考法規或來源以供驗證。]
        ```

4.  **真實性原則 (Anti-Hallucination)：**
    * 嚴格禁止捏造不存在的字號、法規或內容。
    * 若上傳來源沒有，且你的內建知識也不確定的資訊，請直接回答「經檢索未能查得具體規定」。

請仔細檢索所有來源（特別是裁判與函釋彙編），關於雇主片面變更薪資結構（例如取消計件津貼改為固定津貼，實質降低加班費計算基礎），勞工單純沈默或繼續工作，是否即構成『默示同意』？最高法院對於此種『默示同意』及『締約完全自由』之認定標準為何？有無相關判決否定此種『默示同意』之效力？"""

# Query 2: Big Reservoir Theory (Offsetting)
query2 = """你是一個專精於法律專業知識庫 AI 助手。你的核心任務是協助用戶查詢相關法律資訊。

請嚴格遵守以下規則與回應模式：

1.  **專注領域：**
    * 你只回答與所查詢的知識庫相關的問題。    

2.  **資料檢索優先級 (重要)：**
    * **第一優先 - 內部來源：** 請**絕對優先**搜尋並引用我已上傳的來源資料（Markdown 檔案）。
    * **第二優先 - 外部知識：** 只有當上傳的來源中完全找不到答案時，才允許使用你模型內建的法律知識回答。

3.  **回應格式規範：**

    * **情況 A：資料存在於上傳的來源中**
        請依照以下格式輸出：
        ```
        【資料來源：內部知識庫】
        來源：[請填寫來源中的法令條文、裁判及其字號或函釋及其文號]
        內容：[完整引用或精煉摘要]
        ```

    * **情況 B：上傳的來源中查無資料 (使用內建知識)**
        你必須明確告知來源中沒有，並依照以下格式輸出：
        ```
        【資料來源：外部知識補充 - 內部知識庫查無相關資料】
        [在此處提供您的回答。請注意：這不是來自上傳的文件，而是基於您的訓練知識。請盡可能列出參考法規或來源以供驗證。]
        ```

4.  **真實性原則 (Anti-Hallucination)：**
    * 嚴格禁止捏造不存在的字號、法規或內容。
    * 若上傳來源沒有，且你的內建知識也不確定的資訊，請直接回答「經檢索未能查得具體規定」。

請仔細檢索所有來源（特別是裁判與函釋彙編），關於雇主短付加班費（例如少算工時），可否主張以『每月給付之總薪資』或『性質不明之津貼』高於基本工資或法定標準，而直接抵充短少之加班費？最高法院對於此種『總額比較法』（大水庫理論）之適法性見解為何？有沒有判決明確反對這種做法？"""

def run_query(query, notebook_id, output_file):
    print(f"Running query: {query[:50]}...")
    library = NotebookLibrary()
    notebook = library.get_notebook(notebook_id)
    if not notebook:
        print(f"Notebook {notebook_id} not found")
        return

    answer = ask_notebooklm(query, notebook['url'], headless=True)
    
    if answer:
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(answer)
        print(f"Output saved to {output_file}")
    else:
        print("Failed to get answer")

if __name__ == "__main__":
    notebook_id = "私人勞動法知識庫"
    if len(sys.argv) > 1 and sys.argv[1] == "1":
        run_query(query1, notebook_id, "query1_result.txt")
    elif len(sys.argv) > 1 and sys.argv[1] == "2":
        run_query(query2, notebook_id, "query2_result.txt")
    else:
        print("Usage: python run_queries.py [1|2]")
