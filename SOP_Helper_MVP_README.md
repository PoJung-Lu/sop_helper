# SOP_Helper MVP v0.1

Windows 11 本機版原型，用於載入部件圖片、Excel/YAML 料件規格、在圖片上標記鎖螺絲位置，以及輸出 PNG 和可編輯 PPTX。

## 目前支援

- 載入 PNG/JPG/BMP 主圖片
- 載入 Excel `.xlsx` 或 YAML `.yaml/.yml`
- `components` 工作表（若不存在，使用第一張工作表）
- 選擇料件後在圖片上點擊建立標記
- 自動帶入所需扭力與單位
- 修改本次標記的扭力值
- 預覽標註結果
- 複製合成圖片到 Windows 剪貼簿
- 另存 PNG、PPTX、專案 JSON

## 安裝

在 Windows PowerShell 執行：

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
```

如果 PowerShell 不允許啟用腳本，可以直接使用：

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe SOP_Helper_MVP_v0_1.py
```

## 啟動

```powershell
.\.venv\Scripts\python.exe SOP_Helper_MVP_v0_1.py
```

## Excel 欄位

第一列使用以下英文欄位名稱：

```text
part_name
part_number
size
required_torque
torque_unit
applicable_model
reference_image
requires_oring
oring_spec
requires_silicone_grease
silicone_grease_note
torque_min
torque_max
notes
```

預設值：

```text
applicable_model = Universal
requires_oring = false
requires_silicone_grease = false
```

`required_torque` 與 `torque_unit` 是第一版鎖螺絲料件的必要欄位。

## YAML 範例

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
    requires_silicone_grease: false
```

## 操作

1. 載入主圖片。
2. 載入 Excel 或 YAML。
3. 在左側選擇料件。
4. 在圖片上左鍵點擊位置，或按「使用選取料件新增標記」。
5. 在對話框確認扭力與鎖附順序。
6. 按「更新輸出預覽」。
7. 使用「複製合成圖片」、「另存 PNG」、「另存可編輯 PPTX」或「另存專案 JSON」。

## 目前原型限制

- 尚未加入拖曳標記與完整箭頭編輯。
- 參考圖片尚未自動放入輸出版面。
- 「複製 PowerPoint 原生 Shape 到系統剪貼簿」尚未接上 PowerPoint COM；目前先輸出可編輯 PPTX。
- 暫時以固定簡單版面輸出，下一輪依實際測試調整。
