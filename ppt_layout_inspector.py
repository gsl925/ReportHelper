from pptx import Presentation

# --- 您需要修改這裡 ---
TEMPLATE_FILE = "Weekly Report_JimChuang.pptx"  # 請換成您的範本檔案名稱
# ---------------------

try:
    prs = Presentation(TEMPLATE_FILE)
    print(f"成功讀取範本檔案：'{TEMPLATE_FILE}'")
    print("-" * 30)
    print("此範本包含以下版面配置：")
    
    for i, layout in enumerate(prs.slide_layouts):
        print(f"索引 (Index) {i}: {layout.name}")
        
    print("-" * 30)
    print("\n請在您的主程式中，使用您想要的版面配置對應的『索引編號』。")
    print("例如，如果您想用 'Title and Content'，就記下它前面的數字。")

except Exception as e:
    print(f"讀取檔案時發生錯誤：{e}")
    print("請確認您的範本檔案名稱正確，且與此腳本放在同一個資料夾中。")

