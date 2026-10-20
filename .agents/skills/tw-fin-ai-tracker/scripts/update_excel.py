import sys
import json
import openpyxl

def append_scenarios(excel_path, new_items_json):
    wb = openpyxl.load_workbook(excel_path)
    ws_main = wb["AI應用場景"]
    ws_src = wb["來源索引"]

    # 讀取現有 (業者, 場景) 避免重複寫入
    existing = set()
    for r in range(2, ws_main.max_row + 1):
        comp = str(ws_main.cell(r, 2).value or "").strip()
        scen = str(ws_main.cell(r, 3).value or "").strip()
        if comp and scen:
            existing.add((comp, scen))

    added_count = 0
    for item in new_items_json:
        sector = item["sector"].strip()      # 銀行業 / 保險業 / 證券期貨業
        company = item["company"].strip()    # 業者名稱
        scenario = item["scenario"].strip()  # 場景描述
        category = item["category"].strip()  # 14 大類別之一

        # 排除純研究論文未導入業務者，並檢查是否已存在
        if (company, scenario) not in existing:
            ws_main.append([sector, company, scenario, category])
            existing.add((company, scenario))
            added_count += 1

            # 檢查該業者是否已在「來源索引」中，若無則新增一列並帶入 COUNTIF 公式
            known_companies = [
                str(ws_src.cell(r, 2).value or "").strip()
                for r in range(2, ws_src.max_row + 1)
            ]
            if company not in known_companies:
                next_row = ws_src.max_row + 1
                ws_src.cell(next_row, 1, sector)
                ws_src.cell(next_row, 2, company)
                ws_src.cell(next_row, 3, "最新蒐集")
                ws_src.cell(next_row, 4, f'=COUNTIF(AI應用場景!B:B, B{next_row})')

    wb.save(excel_path)
    print(f"成功新增 {added_count} 筆 AI 場景至 {excel_path}")

if __name__ == "__main__":
    # 接收 Agent 傳入的 JSON 字串：python3 update_excel.py '[{"sector":"證券期貨業","company":"凱基證券",...}]'
    items = json.loads(sys.argv[1])
    append_scenarios("金融業AI場景應用.xlsx", items)