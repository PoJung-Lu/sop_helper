# SOP_Helper 專案交接筆記（v0.9.9 → 步驟C 開發前）

- 交接時間：2026-09-21
- 用途：讓新對話 session 匯入本檔案後，可以直接接續開發，不需要重新解釋背景。
- GitHub repo：`https://github.com/PoJung-Lu/sop_helper`（已連接 GitHub 連接器 `github_mcp_direct`，帳號 `PoJung-Lu`）
- 目前最新版本：**v0.9.9**（`SOP_Helper_MVP_v0_9_9.py`），已推送到 `main` 分支，功能測試通過。

---

## 0. 專案目標（不要偏離的核心）

製程工程師手動用 PowerPoint 畫圓圈、箭頭來標示部件照片上的螺絲位置、扭力值與相關資訊，屬於重複且容易出錯的工作。SOP_Helper 是一個 Windows 本機桌面工具，目標：

1. 從 Excel/YAML 規格表自動帶入螺絲/元件的扭力、O-ring、塗矽油等資訊。
2. 讓使用者直接在部件照片上點擊建立標記（螺絲位置），減少手動畫圖。
3. 產生的結果可以直接複製成 **PowerPoint 原生可編輯圖形物件**（不是死圖），貼上後圓圈、箭頭、文字都可以個別修改。
4. 不需要額外伺服器、公司權限開通、雲端上傳；完全本機執行。
5. 第一版只支援「鎖螺絲」這一種步驟類型。

---

## 1. 使用者環境

- Windows 11、Microsoft 365 桌面版 PowerPoint、VS Code（本機模式，不使用 WSL/Docker 執行本專案）、Python 3.13.15
- 專案路徑範例：`C:\Users\paes2\OneDrive\桌面\Python\Work\SOP_helper`
- 職稱：製程工程師，新北市板橋，同時準備 AI/資料科學/水冷製程相關職缺
- 使用者偏好**一步一步小範圍修改**，每次改完都要先測試、明確說「ok」或指出問題後才進行下一步

---

## 2. 版本控制協作模式（本輪新建立，非常重要）

### 2.1 分工方式

- **使用者自己負責所有 GitHub 推送操作**（commit、push、merge），AI 不再主動呼叫 `create_or_update_file`、`push_files` 等寫入類工具。
- AI 的角色：讀取使用者貼上的檔案內容（附件或貼文字）→ 修改/開發 → 產生完整檔案（用 `create_file`）→ 使用者自行推送到 GitHub。
- **原因**：本輪測試發現 AI 這邊的工具（`fetch_url`、`get_file_contents`、`search_code`）都無法完整讀取 GitHub repo 上的大型檔案文字內容（詳見第5節「已知工具限制」），所以改用「使用者貼檔案內容給AI」的方式協作，比讓AI去repo抓取更可靠。

### 2.2 Branch 策略

- 單人開發，日常小修正（bug fix、視覺微調）**直接在 main 分支上進行**。
- 風險較高、需要改資料結構的較大功能（例如接下來的步驟C），**開獨立 feature branch** 開發，測試穩定後才 merge 回 main，避免破壞可正常使用的版本。
- 這次步驟C要開的分支名稱建議：`feature/group-list`

建立分支指令：
```bash
git checkout main
git pull origin main
git checkout -b feature/group-list
git push -u origin feature/group-list
```

完成後合併回main：
```bash
git checkout main
git pull origin main
git merge feature/group-list
git push origin main
```

### 2.3 GitHub repo 現狀

- Repo：`PoJung-Lu/sop_helper`，目前為使用者自行管理是否公開/私有
- 已有的檔案：`SOP_Helper_MVP_v0_9_6.py` ~ `v0_9_9.py`、`requirements.txt`、`components_sample.xlsx`、規格書/進度文件（部分為舊版，內容以本文件與 v0.9.9 為準）、`docs/`、`figs/`、`history/` 資料夾（內容未詳細盤點，新 session 可用 `get_commit`／`list_commits` 或請使用者確認）
- 最新 commit（v0.9.9）：`9412128c18e4cb335aa013412a623e3353285bfc`

---

## 3. 目前已完成功能（v0.1 → v0.9.9 累積成果）

### 3.1 開發環境（已完成，不需重做）

- Windows 11 + Python 3.13.15、獨立虛擬環境 `.venv`
- 已安裝套件：`PySide6`、`Pillow`、`openpyxl`、`PyYAML`、`python-pptx`、`pywin32`
- PowerShell 執行原則已設定為 `RemoteSigned`

### 3.2 料件規格載入（Excel / YAML）— 與 v0.9.1 相同，未變更

- 支援 `.xlsx`（工作表名固定找 `components`）與 `.yaml`/`.yml`
- 欄位：`part_name, part_number, size, required_torque, torque_unit, applicable_model, reference_image, requires_oring, oring_spec, requires_silicone_grease, silicone_grease_note, torque_min, torque_max, notes`
- 驗證規則：`part_name`/`part_number`/`size` 必填；`part_number` 不可重複；`required_torque` 與 `torque_unit` 必須同時填或同時空

### 3.3 扭力規則 — 與 v0.9.1 相同，未變更

- 單位選單：`kgf-cm`、`N-M`、`無需扭力`、`鎖緊就好`
- 扭力數值輸入 `0` 時自動切換單位為「無需扭力」

### 3.4 標記（Annotation）資料結構

```python
@dataclass
class Annotation:
    annotation_id: str
    component_index: int
    x: float
    y: float
    sequence: int
    torque: float | None
    torque_unit: str
    marker_only: bool = False
    label_x: float = 0.0
    label_y: float = 0.0
```

**注意**：`label_x`/`label_y` 目前仍是保留欄位，尚未真正使用。

### 3.5 操作模式（畫布互動）— 步驟B已完成

- 左側料件清單「有沒有選取料件」決定畫布左鍵點擊的行為（沒選料件→選取/拖曳既有標記；選了料件→新增標記）
- **標記可拖曳移動**（`_dragging_annotation_id` 追蹤拖曳狀態，`update_annotation_position()` 即時更新座標）
- **雙擊既有標記可重新編輯**（`edit_annotation()` 開啟 `AnnotationDialog` 修改扭力/單位/順序/僅標記）
- **可刪除任一選取的標記**（不限最後一個，`delete_selected_annotation()`）
- 縮放：`Ctrl + 滑鼠滾輪`，以滑鼠游標位置為錨點縮放，範圍 25%~400%
- 平移：滑鼠中鍵拖曳
- 拖曳圖片檔案到畫布可載入為主圖片

### 3.6 畫布標記視覺（v0.9.9 修正的關鍵bug）

- **圓圈半徑改用固定螢幕像素大小換算**：`CANVAS_MARKER_SCREEN_RADIUS = 14`，透過 `self.canvas.transform().m11()`（QGraphicsView 實際變換矩陣係數）換算回場景座標，不再受載入圖片的原始解析度影響，縮放時（`fit_to_window`/`Ctrl+滾輪`）會自動觸發 `update_canvas_marks()` 重繪
- 兩位數以上標記序號自動縮小字級（`digit_count` 判斷，1位數/2位數/3位數以上各有不同字級比例）
- PNG/PPTX/畫布三處視覺比例統一用 `PPTX_MARKER_RADIUS_RATIO = 0.10 / 8.5` 常數

### 3.7 PowerPoint 輸出（`export_pptx()`）

- **圓圈**：`radius=Inches(.10)`，外框線寬 `Pt(1.5)`，數字字級依位數動態調整（`sequence_font_size()`），`word_wrap=False` 避免自動換行
- **料件名稱**：獨立 `TextBox`，`word_wrap=True`
- **扭力方框**：`fit_box_size()` 依文字長度與字級動態計算寬高，短文字（如「鎖緊就好」）方框較小，長數字方框較寬；外框線寬 `Pt(0.75)`
- **箭頭**：`add_connector`，線寬 `Pt(0.75)`（與扭力方框外框一致），用 `add_arrowhead()` 改寫 XML `<a:tailEnd>` 加上三角形箭頭頭；起點綁定在說明方框（或無扭力時的名稱文字框）左緣中點（連接點索引1），終點綁定在標記圓圈右緣中點（索引3），用 `begin_connect`/`end_connect` 實現 PowerPoint 原生連接點綁定——拖曳方框或圓圈時箭頭自動跟隨
- **「複製為 PowerPoint 物件」功能已驗證**：圓圈、文字、方框、箭頭都能個別選取、移動、修改文字；中文字型 Microsoft JhengHei 正確顯示

### 3.8 版面配置修正

- 左側「資料與料件」與中間「主圖與標記」區塊的提示文字（`mode_hint`、`component_hint_label`、`hint_label`）統一改用 `setWordWrap(True)`，依容器實際寬度動態換行，不寫死 `setMaximumWidth`
- 修正「顯示輸出預覽」後「標記清單/輸出預覽」面板無法拖曳調整大小的bug（`preview_label` 加上 `setMinimumSize(1,1)` 與 `setSizePolicy(Ignored, Ignored)`，避免 QPixmap 實際尺寸撐開 QSplitter 最小需求）

---

## 4. 已知限制 / 尚未完成（誠實列出）

1. **PowerPoint 群組化未完成**：同一元件的圓圈、箭頭、名稱、扭力框還沒自動 `Group()` 成一個 PowerPoint Group
2. **分層標記清單未完成**：右側清單目前是平面清單，還沒有分層/分群組結構、`Tab` 建群組、`Ctrl/Shift` 多選、清單拖曳排序 —— **這是接下來步驟C要做的**
3. **共用說明文字＋多箭頭未完成**：目前每個標記各自獨立顯示名稱與扭力，還沒做「同群組多個標記共用一個說明文字框、多條箭頭從同一文字框延伸」的邏輯 —— **步驟D，依賴步驟C**
4. **參考圖片（`reference_image` 欄位）尚未實際使用**於 PNG/PPTX 輸出
5. **開啟既有專案 JSON 還原畫面**尚未實作，只能存檔不能讀回
6. **尚未打包成 EXE**（PyInstaller）
7. **尚未做正式的單元測試**，目前都是手動人工測試
8. **O-ring／塗矽油資訊**只在「設定標記」對話框提示區顯示，尚未輸出到 PNG/PPTX 版面上
9. **預覽面板圖片沒有縮放/平移功能**（已歸類為小改進，優先度較低，可在步驟C之後處理）

---

## 5. 已知工具限制（AI協作注意事項）

在本輪對話中確認：

- AI 使用的 `fetch_url` 工具**完全無法讀取 GitHub 網域內容**（包含 `github.com` 網頁版、`raw.githubusercontent.com` 純文字端點），無論 repo 公開或私有，多次測試皆失敗，原因可能是該工具對 GitHub 網域有阻擋或反爬蟲機制不相容
- GitHub 連接器（`github_mcp_direct`）的 `get_file_contents` 工具，對這個環境而言**只回傳「下載成功＋SHA」的狀態摘要，不會把檔案文字內容顯示給AI**，即使是82 bytes的極小檔案也一樣，這是工具回應格式的限制
- `search_code` 對這個repo**索引搜尋不到任何結果**（即使用最寬鬆的 `extension:py` 條件也是0筆）
- **結論**：AI 沒有辦法自行從 GitHub repo 讀取檔案內容。**協作時務必由使用者主動貼上檔案內容（文字或附件上傳）給AI**，不要期待AI自己去repo抓取比對。

---

## 6. 下一步：步驟C（分層群組清單）

### 6.1 目標

新增 `group_id` 欄位到 `Annotation`，讓多個標記可以被歸入同一群組，右側清單改為分層顯示，支援多選建立群組。這是後續「共用說明文字＋多箭頭延伸」（步驟D）的資料結構基礎。

### 6.2 建議拆分成兩個小版本

**v0.10.0（資料結構＋多選＋Tab建群組骨架）**：
1. `Annotation` dataclass 新增 `group_id: str | None = None` 欄位
2. 右側標記清單（`marker_list`）改為支援 `Ctrl/Shift` 多選（`QListWidget.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)`）
3. 多選後按 `Tab` 鍵觸發建立群組（需要攔截 `keyPressEvent` 或加一個「建立群組」按鈕，兩種方式皆可，建議先做按鈕版本更直覺，之後再加鍵盤快速鍵）
4. 建立群組規則：系統自動把第一個選中項目設為顯示（`marker_only=False`），其餘自動設為 `marker_only=True`，使用者之後可個別手動覆蓋（透過既有的雙擊編輯功能）
5. 群組內只支援單層，不能再有子群組
6. 清單顯示需要能區分「有群組」與「無群組」的標記，例如群組內項目加縮排或群組標題列

**v0.10.1（共用文字框＋多箭頭延伸，即步驟D）**：
1. 每個群組找「第一個 `marker_only=False` 的項目」作為 `label_owner`，提供共用文字框
2. 群組內 `marker_only=True` 的項目，箭頭從 `label_owner` 的文字框延伸到自己的圓圈
3. 若群組內**全部**項目都是 `marker_only=True`（沒有 `label_owner`），則不畫任何箭頭，只顯示圓圈（已確認規則：選項A）
4. 說明文字框在畫布上要可拖曳，拖曳時所有連到它的箭頭終點需要重新計算
5. PNG 與 PPTX 輸出都要同步支援「一個文字框、多條箭頭」的畫法

### 6.3 開發時的注意事項

- **修改時盡量保持既有的資料結構命名**（`Component`、`Annotation`、`export_pptx`、`render_composite` 等），除非該步驟本身就是要擴充資料結構
- **每次生成新版本，用 `create_file` 工具產生完整可執行的單一 `.py` 檔案**（不要用 diff 或片段），檔名遞增版本號
- 使用者偏好**一步一步小範圍修改**，每次改完都要先讓他測試、明確說「ok」或指出問題後才進行下一步，不要一次塞多個功能
- 這份交接文件本身也應該隨版本更新；當完成v0.10.0、v0.10.1後，應更新「已完成」與「已知限制」章節

---

## 7. 給下一個 session 的操作提示

1. 使用者已有 Windows 本機開發環境，**不需要重新引導安裝 Python 或建立虛擬環境**
2. 使用者已自行在本機建立 `feature/group-list` branch（或告知新session目前在哪個branch），**開發時確認目前是在哪個分支上討論**
3. 目前最新可用版本是 **v0.9.9**（`SOP_Helper_MVP_v0_9_9.py`），這是接下來開發的起點
4. **AI 無法從 GitHub repo 直接讀取檔案內容**（見第5節），需要使用者主動貼上/上傳最新版本檔案給AI核對
5. 修改完成後，AI 用 `create_file` 產生完整檔案給使用者，**使用者自行推送到 GitHub**（AI 不再主動呼叫 GitHub 寫入類工具）
6. 這次要開發的是**步驟C：分層群組清單**，建議先做 v0.10.0（資料結構+多選+Tab建群組骨架），確認可運作後再做 v0.10.1（共用文字框+多箭頭）
