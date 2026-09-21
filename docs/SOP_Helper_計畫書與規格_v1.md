# SOP_Helper 計畫書與規格 v1

- 版本：v1.0
- 日期：2026-09-17
- 目標平台：Windows 11
- 目標 Office：Microsoft 365 桌面版 PowerPoint
- 執行模式：本機離線桌面工具
- 第一版步驟類型：鎖螺絲

---

## 1. 專案目標

SOP_Helper 是一個在 Windows 11 本機執行的桌面輔助工具，將部件原始照片轉換成可直接用於 SOP 的圖形結果。

使用者可以：

1. 載入部件原始圖片。
2. 載入 Excel 或 YAML 料件規格。
3. 在圖片預覽上點選要鎖附的螺絲或元件位置。
4. 選擇料件並自動帶入規格資料。
5. 修改本次使用的扭力與其他欄位。
6. 預覽箭頭、圓圈、文字與參考圖片。
7. 將結果複製成 PowerPoint 原生圖形物件。
8. 另存可編輯 PPTX、圖片與專案檔。

---

## 2. 已確認需求

### 2.1 執行環境

- 作業系統：Windows 11。
- Office：Microsoft 365 桌面版 PowerPoint。
- 不依賴額外伺服器。
- 不需要公司開通權限。
- 預設不需要外部網路。
- 不將公司圖片或規格上傳至外部服務。

### 2.2 第一版範圍

- 只支援一種製程步驟：鎖螺絲。
- 一次處理一張主圖。
- 支援多個螺絲或元件標記。
- 支援在預覽圖上直接新增、移動與刪除標記。
- 輸出時保留原始圖片與獨立 PowerPoint 圖形物件。
- 扭力資料可由 Excel 或 YAML 帶入。
- 使用者允許修改帶入的扭力值。

### 2.3 輸出需求

必須提供：

- 即時預覽。
- 複製為 PowerPoint 原生圖形物件。
- 另存 PPTX。
- 另存合成 PNG。
- 可選擇另存可重新編輯的專案檔。

PowerPoint 中的圓圈、箭頭、文字與參考圖片必須是獨立物件，而不是全部合併成一張圖片。

---

## 3. 使用者操作流程

```text
開啟 SOP_Helper.exe
        ↓
載入主圖片
        ↓
載入 Excel 或 YAML 規格
        ↓
在預覽圖上點選鎖附位置
        ↓
選擇料件
        ↓
自動帶入料件資訊與扭力
        ↓
必要時修改本次扭力或其他設定
        ↓
設定箭頭、圓圈與文字位置
        ↓
預覽輸出
        ↓
驗證
        ↓
複製 PowerPoint 原生物件
或另存 PPTX / PNG / 專案檔
```

---

## 4. 料件規格資料

### 4.1 必要欄位

Excel 與 YAML 欄位名稱可使用英文，但必須包含以下資訊：

| 欄位 | 英文欄位名稱 | 必填 | 預設值 | 說明 |
|---|---|---:|---|---|
| 料件名稱 | `part_name` | 是 | 無 | 顯示給使用者的名稱 |
| 料件編號 | `part_number` | 是 | 無 | 公司料號或規格編號 |
| 尺寸 | `size` | 是 | 無 | 例如 M4 x 12、3 inch |
| 所需扭力 | `required_torque` | 鎖螺絲必填 | 無 | 數值欄位 |
| 扭力單位 | `torque_unit` | 鎖螺絲必填 | 無 | 例如 kgf-cm、N*m |
| 適用機型 | `applicable_model` | 否 | `Universal` | 適用產品或機型 |
| 參考圖片 | `reference_image` | 否 | 空白 | 元件示意圖或照片 |
| 是否需搭配 O-ring | `requires_oring` | 否 | `false` | 是或否 |
| O-ring 規格 | `oring_spec` | 否 | 空白 | 若需要 O-ring，填寫規格 |
| 是否需圖矽油 | `requires_silicone_grease` | 否 | `false` | 是或否 |
| 圖矽油說明 | `silicone_grease_note` | 否 | 空白 | 塗佈或材料說明 |
| 扭力下限 | `torque_min` | 否 | 空白 | 若未填，可依規則設定或警告 |
| 扭力上限 | `torque_max` | 否 | 空白 | 若未填，可依規則設定或警告 |
| 備註 | `notes` | 否 | 空白 | 其他工程說明 |

> 第一版允許使用者修改扭力值，但修改後的值屬於本次 SOP 使用值，不會反向修改原始 Excel 或 YAML 規格檔。

### 4.2 Excel 格式

工作表名稱建議固定為：

```text
components
```

第一列為欄位名稱，每一列代表一個料件。

範例：

| part_name | part_number | size | required_torque | torque_unit | applicable_model | reference_image | requires_oring | oring_spec | requires_silicone_grease | silicone_grease_note |
|---|---|---|---:|---|---|---|---|---|---|---|
| PT100 | PT100-001 | M4 x 12 | 180 | kgf-cm | Universal | pt100.png | false |  | false |  |
| 壓力計_JPR2 | JPR2-001 | 1/4 inch | 200 | kgf-cm | Universal | jpr2.png | true | O-ring 1/4 | true | 螺紋薄塗 |

### 4.3 YAML 格式

```yaml
components:
  - part_name: PT100
    part_number: PT100-001
    size: M4 x 12
    required_torque: 180
    torque_unit: kgf-cm
    applicable_model: Universal
    reference_image: pt100.png
    requires_oring: false
    oring_spec: ""
    requires_silicone_grease: false
    silicone_grease_note: ""
    torque_min: 160
    torque_max: 200
    notes: ""

  - part_name: 壓力計_JPR2
    part_number: JPR2-001
    size: 1/4 inch
    required_torque: 200
    torque_unit: kgf-cm
    applicable_model: Universal
    reference_image: jpr2.png
    requires_oring: true
    oring_spec: O-ring 1/4
    requires_silicone_grease: true
    silicone_grease_note: 螺紋薄塗
    torque_min: 180
    torque_max: 220
    notes: ""
```

### 4.4 預設值

若 Excel 或 YAML 缺少下列欄位，程式填入：

```text
applicable_model = Universal
requires_oring = false
requires_silicone_grease = false
reference_image = 空白
notes = 空白
```

但以下欄位不可缺少：

```text
part_name
part_number
size
required_torque（鎖螺絲料件）
torque_unit（鎖螺絲料件）
```

---

## 5. 扭力修改規則

使用者選擇料件後，程式自動帶入：

- 所需扭力。
- 扭力單位。
- 扭力下限與上限，如規格存在。
- 適用機型。
- 尺寸。
- O-ring 要求。
- 圖矽油要求。

使用者可以修改本次 SOP 的：

- 所需扭力。
- 扭力單位。
- 扭力上下限。
- 文字說明。
- 標記樣式。
- 參考圖片。

修改後必須保留：

```text
原始規格值
本次使用值
是否修改
修改原因（建議欄位）
```

若 `required_torque` 修改，預覽與輸出必須明確顯示本次使用值。

---

## 6. 標記與畫布

### 6.1 第一版標記類型

- `fastener`：螺絲或需鎖附料件。
- `component`：被鎖上的元件。
- `oring`：O-ring 或墊片位置。
- `grease_area`：圖矽油塗佈位置。

第一版流程仍以 `fastener` 為主要驗證對象。

### 6.2 滑鼠操作

- 左鍵點擊：新增標記。
- 點擊標記：選取標記。
- 拖曳標記：移動標記。
- Delete：刪除標記。
- 滑鼠滾輪：放大與縮小。
- 拖曳空白區域：平移圖片。
- 雙擊標記：開啟屬性設定。

### 6.3 標記內容

每個標記包含：

- 標記 ID。
- 料件編號。
- 料件名稱。
- 圖片相對座標。
- 標記類型。
- 鎖附順序。
- 箭頭起點。
- 箭頭終點。
- 標籤位置。
- 是否顯示扭力。
- 是否顯示參考圖片。
- 本次扭力覆寫值。

座標使用 0 到 1 的相對座標，以避免圖片縮放後位置錯誤。

---

## 7. 輸出預覽

輸出預覽要接近目前人工製作的 PPT 圖片效果：

- 主圖放在中央。
- 左上或指定位置顯示圖號。
- 箭頭指向實際安裝位置。
- 元件名稱放在箭頭附近。
- 扭力值顯示在文字框中。
- 元件參考圖放在右側或指定區域。
- O-ring 與圖矽油要求可用文字或小圖顯示。

範例文字：

```text
PT100
180 kgf-cm
O-ring：無
圖矽油：無
```

或：

```text
壓力計_JPR2
200 kgf-cm
O-ring：O-ring 1/4
圖矽油：螺紋薄塗
```

### 7.1 樣式預設

- 料件標記：紅色箭頭。
- 鎖附位置：紅色圓圈或定位點。
- 鎖附順序：紅色數字。
- 扭力文字框：白底、紅色外框。
- 圖號：深藍底、白色文字。
- 一般提示：藍色箭頭。

---

## 8. PowerPoint 輸出

### 8.1 可編輯物件

輸出的 PPTX 必須包含獨立物件：

```text
原始主圖 Picture
標記圓圈 Oval
方向箭頭 Line / Arrow
鎖附順序 TextBox
料件名稱 TextBox
扭力資訊 TextBox
O-ring 資訊 TextBox
圖矽油資訊 TextBox
參考圖片 Picture
```

使用者貼入 PowerPoint 後可以：

- 移動。
- 拉伸。
- 旋轉。
- 改色。
- 修改文字。
- 刪除。
- 改變圖層順序。
- 重新排列參考圖片。

### 8.2 複製為 PowerPoint 原生物件

按下：

```text
複製 PowerPoint 圖形物件
```

程式流程：

1. 建立暫存 PPTX。
2. 用 PowerPoint COM 開啟暫存簡報。
3. 選取輸出投影片上的所有 Shape。
4. 呼叫 PowerPoint Shape Copy。
5. 將 Shape 放入 Windows 剪貼簿。
6. 提示使用者切換到目標簡報並按 Ctrl+V。

若 COM 複製失敗，必須提供明確訊息並保留其他輸出按鈕。

### 8.3 另存格式

```text
SOP_圖號_annotated.png
SOP_圖號_editable.pptx
SOP_圖號_project.json
```

PNG 是合成圖；PPTX 保留獨立物件；JSON 用來重新開啟與編輯。

---

## 9. 專案檔格式

```json
{
  "project_version": "1.0",
  "title": "圖3",
  "source_image": "source/image.jpg",
  "spec_source": "spec/components.xlsx",
  "annotations": [
    {
      "id": "F001",
      "type": "fastener",
      "part_number": "PT100-001",
      "part_name": "PT100",
      "x": 0.42,
      "y": 0.36,
      "sequence": 1,
      "torque": {
        "original": 180,
        "used": 180,
        "unit": "kgf-cm",
        "overridden": false,
        "reason": ""
      },
      "arrow": {
        "start_x": 0.32,
        "start_y": 0.18,
        "end_x": 0.42,
        "end_y": 0.36
      },
      "show_reference_image": true
    }
  ]
}
```

---

## 10. 技術架構

### 10.1 建議技術

```text
GUI：PySide6
圖片與繪圖：Pillow、QPainter
Excel：openpyxl
YAML：PyYAML
資料驗證：Pydantic
PPTX：python-pptx
PowerPoint 剪貼簿：pywin32 + PowerPoint COM
打包：PyInstaller
測試：pytest
```

### 10.2 模組

```text
sop_helper/
├── app.py
├── ui/
│   ├── main_window.py
│   ├── component_panel.py
│   ├── annotation_canvas.py
│   └── preview_panel.py
├── models/
│   ├── component.py
│   ├── annotation.py
│   └── project.py
├── loaders/
│   ├── excel_loader.py
│   └── yaml_loader.py
├── validators/
│   └── rules.py
├── renderers/
│   ├── image_renderer.py
│   └── pptx_renderer.py
├── integrations/
│   └── powerpoint_clipboard.py
├── resources/
├── tests/
└── requirements.txt
```

---

## 11. 測試計畫

### 11.1 單元測試

#### 資料載入

- 正常 Excel 可載入。
- 正常 YAML 可載入。
- 缺少 `components` 工作表時顯示錯誤。
- 缺少 `part_name` 時失敗。
- 缺少 `part_number` 時失敗。
- 缺少 `size` 時失敗。
- `applicable_model` 缺少時自動填入 `Universal`。
- `requires_oring` 缺少時自動填入 `false`。
- `requires_silicone_grease` 缺少時自動填入 `false`。

#### 扭力驗證

- `required_torque` 是數值時通過。
- 扭力不是數值時失敗。
- 鎖螺絲料件沒有扭力時失敗。
- 不支援的扭力單位時失敗或警告。
- 扭力覆寫後使用值正確更新。
- 原始扭力值仍被保留。

#### 標記

- 點擊新增標記。
- 拖曳標記後座標正確。
- Delete 刪除標記。
- 標記 ID 不重複。
- 鎖附順序不重複。
- 每個標記都對應有效料件。
- 圖片縮放後標記仍在正確位置。

### 11.2 圖片輸出測試

- 主圖可正常載入。
- 圓圈位置正確。
- 箭頭終點正確指向標記。
- 料件名稱正確。
- 扭力資訊正確。
- O-ring 資訊正確。
- 圖矽油資訊正確。
- 參考圖存在時正常顯示。
- 參考圖不存在時顯示清楚警告。
- 中文不亂碼。

### 11.3 PPTX 測試

- PPTX 可以由 Microsoft 365 PowerPoint 開啟。
- 主圖是獨立 Picture。
- 圓圈是獨立 Shape。
- 箭頭是獨立 Shape。
- 文字是可編輯 TextBox。
- 參考圖是獨立 Picture。
- 所有物件可以移動。
- 所有物件可以縮放。
- 箭頭可以旋轉。
- 文字可以修改。
- 不會全部變成單一圖片。

### 11.4 剪貼簿測試

- PowerPoint COM 可以啟動。
- 暫存投影片可以建立。
- Shape 可以複製。
- 使用者在目標 PowerPoint 按 Ctrl+V 後成功貼上。
- 貼上後物件仍可獨立編輯。
- PowerPoint 未啟動時顯示清楚提示。
- COM 失敗時可改用另存 PPTX。
- 複製失敗時可改用複製 PNG。

### 11.5 使用者驗收測試

不熟悉程式的同事應能在 10 分鐘內完成：

1. 開啟工具。
2. 載入一張圖片。
3. 載入 Excel。
4. 選擇一個料件。
5. 在圖片上標記位置。
6. 確認扭力值。
7. 看到輸出預覽。
8. 複製原生 PowerPoint 物件。
9. 貼入既有簡報。
10. 移動箭頭並修改文字。

測試期間不需要命令列、網路、管理員權限或額外伺服器。

---

## 12. 測試即通過原則

### P0：不可妥協

以下任一項失敗，版本不可交付：

- 程式無法啟動。
- Windows 11 無法執行。
- 正常 Excel 無法讀取。
- 正常 YAML 無法讀取。
- 料件名稱或料件編號錯誤。
- 鎖螺絲料件缺少扭力仍可無警告輸出。
- 扭力值輸出錯誤。
- PPTX 無法被 Microsoft 365 開啟。
- 圓圈、箭頭或文字變成不可編輯的單張圖片。
- PowerPoint 貼上後無法獨立選取物件。
- 參考圖片或原圖被意外覆寫。
- 任何公司圖片或規格被傳送到外部服務。

### P1：必須修正

- 圖片縮放後標記偏移。
- 箭頭沒有指到正確位置。
- 標記編號重複。
- 輸出內容與預覽不一致。
- 自訂扭力沒有被清楚標示。
- 參考圖片遺失但沒有警告。
- 中文顯示錯誤。
- 剪貼簿失敗時沒有備援按鈕。
- 產生的檔名或輸出資料夾不清楚。

### P2：可延後

- 自動避免文字重疊。
- 自動最佳化箭頭路徑。
- 多張主圖。
- 多種製程步驟。
- OCR。
- AI 自動辨識螺絲。
- 多人協作。
- 資料庫或 MES 串接。

---

## 13. 開發階段

### Phase 1：資料模型與規格載入

完成：

- Excel 載入。
- YAML 載入。
- 預設值補齊。
- 扭力驗證。
- 專案 JSON。

### Phase 2：圖片與標記

完成：

- 圖片預覽。
- 點擊新增標記。
- 拖曳標記。
- 刪除標記。
- 元件選擇。
- 顯示扭力。

### Phase 3：輸出預覽

完成：

- 圓圈。
- 箭頭。
- 編號。
- 元件名稱。
- 扭力框。
- O-ring 資訊。
- 圖矽油資訊。
- 參考小圖。

### Phase 4：PowerPoint

完成：

- 可編輯 PPTX。
- 獨立 Shape。
- 另存 PNG。
- 複製合成圖片。
- 複製原生 PowerPoint 物件。

### Phase 5：打包與驗收

完成：

- PyInstaller EXE。
- 無網路測試。
- 無管理員權限測試。
- Microsoft 365 測試。
- 同事使用者驗收。
- 錯誤訊息與使用說明。

---

## 14. 第一個實作步驟

第一步只建立開發環境與專案骨架，不加入 GUI 功能。

完成標準：

- 建立專案資料夾。
- 建立 Python 虛擬環境。
- 安裝必要套件。
- 建立模組資料夾。
- 建立 `requirements.txt`。
- 執行最小測試確認環境可用。

完成後等待使用者回覆 `ok`，再進入下一步。
