# SOP_Helper 環境、資源與應用背景記錄

- 版本：對應程式 v0.9.1
- 記錄日期：2026-09-18
- 用途：完整記錄開發環境、使用到的技術資源、應用背景與目標，供新對話 session 或交接使用時快速還原上下文。

---

## 1. 應用背景

### 1.1 使用者角色

- 職稱：製程工程師。
- 所在地：新北市板橋。
- 目前工作內容：手動撰寫作業流程 SOP，包含部件照片標註、螺絲鎖附位置與扭力規格說明。
- 職涯發展方向：同時在準備 AI／資料科學與水冷／先進製程工程相關職缺，因此這個專案除了解決實際工作痛點，也作為個人技術作品集的一部分。

### 1.2 原始問題

目前製作 SOP 的方式：

1. 用 PowerPoint 手動在部件照片上畫圓圈標示螺絲位置。
2. 用箭頭連接文字說明與圓圈。
3. 手動輸入螺絲料號、扭力值、O-ring／塗矽油等規格文字。
4. 每次修改都要重新調整圖形位置、對齊、字型。

主管希望能將 SOP 模板化：例如只要步驟是「鎖螺絲」，就必須包含扭力值等固定欄位，把整個作業流程的資料結構標準化，而不只是文件格式好看。

### 1.3 專案目標

開發一個 Windows 本機桌面工具（SOP_Helper），取代手動 PowerPoint 繪圖流程：

1. 用 Excel／YAML 作為料件規格的單一資料來源（螺絲料號、尺寸、扭力、O-ring、塗矽油需求等），避免規格分散、口耳相傳或每次重新輸入。
2. 讓使用者直接在部件照片上點擊建立標記，取代手動畫圓圈、箭頭。
3. 選擇料件後自動帶入規格值，仍允許依實際情況修改本次數值。
4. 最終產出可以直接貼進 PowerPoint 簡報，且貼上後的圓圈、箭頭、文字都是**獨立可編輯的 PowerPoint 物件**，不是死圖，方便日後其他工程師微調。
5. 不依賴公司額外開通權限、不需要架設伺服器、不上傳資料到外部服務——完全本機執行，因為公司會擋外部上傳。
6. 第一版範圍限定在「鎖螺絲」這一種製程步驟，驗證可行後再擴充其他步驟類型。

### 1.4 為什麼選擇這個技術路線（決策紀錄）

以下是討論過程中確認、且已經定案的關鍵決策，避免之後重新討論同樣的問題：

| 決策點 | 選擇 | 原因 |
|---|---|---|
| 部署方式 | Windows 本機桌面應用程式（非網頁、非 Colab） | 公司會擋外部上傳；同事不需要額外設備或權限開通；不用架設內網伺服器 |
| GUI 框架 | PySide6 | 需要互動式圖片畫布（縮放、平移、點擊標記），比原生 Tkinter 更適合 |
| 資料來源優先順序 | Excel 優先，YAML 為備用格式 | 公司已有 Excel 規格書習慣，不希望同事額外維護 YAML |
| PowerPoint 輸出方式 | 提供三層備援：複製 PowerPoint 原生物件 → 另存可編輯 PPTX → 複製合成圖片 | 第一種最理想但依賴 PowerPoint COM，可能因環境差異失敗，需要備援 |
| 扭力欄位驗證原則 | 「有鎖螺絲步驟就必須有扭力值」，但允許明確標示「無需扭力」或「鎖緊就好」 | 符合主管模板化要求，同時允許 O-ring 等非鎖螺絲元件存在 |
| 打包方式（規劃中） | PyInstaller 打包成單一 EXE | 同事不應該需要安裝 Python 或套件 |

---

## 2. 開發環境

### 2.1 作業系統與硬體

- 作業系統：Windows 11。
- 使用者帳號路徑範例：`C:\Users\paes2\`
- 專案資料夾範例：`C:\Users\paes2\OneDrive\桌面\Python\Work\SOP_helper`
- 注意事項：專案位於 OneDrive 同步資料夾內，建議將 `.venv/`、`__pycache__/`、暫存輸出檔案排除在 OneDrive 同步或至少加入 `.gitignore`，避免大量小檔案拖慢同步或造成鎖檔問題。

### 2.2 開發工具

- 編輯器：VS Code，使用**本機模式**（非 Remote-WSL、非 Dev Container）。
- 終端機：PowerShell（非 WSL Bash、非 Conda base 環境）。
- 版本控制：GitHub（已連接為外部工具，可用於後續備份與版本追蹤，目前尚未實際建立 repo 存放此專案，屬於待辦事項）。
- 背景並行環境：使用者的電腦同時在背景執行 WSL2 與 Docker Desktop（用於其他專案），SOP_Helper 開發**刻意與這兩者隔離**，全程使用 Windows 原生 Python 環境，避免互相干擾。

### 2.3 Python 環境

- Python 版本：3.13.15（從 [python.org](https://www.python.org/downloads/windows/) 官方安裝，安裝時已勾選 `Add python.exe to PATH`）。
- 虛擬環境：專案根目錄下的 `.venv`（用 `python -m venv .venv` 建立，與其他 Python 專案如量子機器學習研究環境完全隔離）。
- PowerShell 執行原則：因預設會擋 `.ps1` 腳本執行，已執行過：
  ```powershell
  Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
  ```
  此設定僅影響目前使用者帳號的 PowerShell 腳本執行權限，不影響 WSL、Docker 或其他使用者。

### 2.4 已安裝套件（`requirements.txt`）

```text
PySide6>=6.7
Pillow>=10.0
openpyxl>=3.1
PyYAML>=6.0
python-pptx>=1.0
pywin32>=306
```

各套件用途：

| 套件 | 用途 |
|---|---|
| PySide6 | 桌面 GUI 框架，提供互動式圖片畫布、對話框、清單元件 |
| Pillow | 圖片讀取、繪製標註（圓圈、文字、方框）、輸出 PNG |
| openpyxl | 讀取 Excel `.xlsx` 料件規格表 |
| PyYAML | 讀取 YAML 格式的備用規格檔 |
| python-pptx | 產生可編輯的 PowerPoint `.pptx` 檔案與其中的 Shape 物件 |
| pywin32 | 呼叫 Windows 剪貼簿 API（複製圖片）與 PowerPoint COM 自動化（複製原生 PowerPoint 物件） |

### 2.5 系統資源依賴

- 中文字型：`C:\Windows\Fonts\msjh.ttc`（微軟正黑體），程式直接讀取此路徑產生圖片文字；若此檔案不存在會直接報錯（設計上刻意不做字型備援，避免又發生中文亂碼問題）。
- PowerPoint 應用程式：Microsoft 365 桌面版，已確認使用者電腦已安裝，「複製為 PowerPoint 物件」功能透過 `pywin32` 呼叫本機 PowerPoint COM 介面實現。

---

## 3. 專案檔案清單（目前已產出的檔案）

### 3.1 程式主體（依版本遞增，目前最新為 v0.9.1）

| 檔名 | 版本 | 狀態 |
|---|---|---|
| SOP_Helper_MVP_v0_1.py | v0.1 | 已淘汰，僅供歷史參考 |
| SOP_Helper_MVP_v0_2.py | v0.2 | 已淘汰 |
| SOP_Helper_MVP_v0_3.py | v0.3 | 已淘汰 |
| SOP_Helper_MVP_v0_4.py | v0.4 | 已淘汰 |
| SOP_Helper_MVP_v0_5.py | v0.5 | 已淘汰 |
| SOP_Helper_MVP_v0_6.py | v0.6 | 已淘汰 |
| SOP_Helper_MVP_v0_7.py | v0.7 | 已淘汰 |
| SOP_Helper_MVP_v0_8.py | v0.8 | 已淘汰 |
| SOP_Helper_MVP_v0_9.py | v0.9 | 已淘汰（選取狀態有 bug） |
| **SOP_Helper_MVP_v0_9_1.py** | **v0.9.1** | **目前使用中，最新版本** |

### 3.2 支援檔案

| 檔名 | 用途 |
|---|---|
| requirements.txt | Python 套件安裝清單 |
| SOP_Helper_MVP_README.md | 早期版本安裝與操作說明（內容對應 v0.1，部分欄位說明仍適用，但操作流程需以最新版本實際介面為準） |
| components_sample.xlsx | 測試用料件規格範例（PT100、壓力計_JPR2、水冷板固定螺絲、三通接頭O-ring 四筆資料） |

### 3.3 規劃與進度文件

| 檔名 | 用途 |
|---|---|
| SOP_Helper_計畫書與規格_v1.md | 最初的完整規格書，包含資料模型、測試計畫、P0/P1/P2 測試即通過原則（測試框架仍然有效，可對照使用） |
| SOP_Helper_專案進度與後續計畫_v0_9_1.md | 目前已完成功能、已知限制、下一步開發順序（步驟 A～H）的詳細清單 |
| SOP_Helper_環境與資源清單_v0_9_1.md | 本文件，記錄環境、資源與應用背景 |

---

## 4. 資料格式規範（沿用至今，尚未變更）

### 4.1 Excel 欄位（工作表名固定為 `components`）

```text
part_name                  料件名稱（必填）
part_number                 料件編號（必填，不可重複）
size                        尺寸（必填）
required_torque              所需扭力數值（鎖螺絲類必填，需與 torque_unit 同時填或同時空）
torque_unit                 扭力單位（鎖螺絲類必填；需與 required_torque 同時填或同時空）
applicable_model             適用機型（選填，預設 Universal）
reference_image              參考圖片檔名（選填，尚未實際用於輸出）
requires_oring               是否需搭配 O-ring（選填，預設 false）
oring_spec                   O-ring 規格說明（選填）
requires_silicone_grease      是否需塗矽油（選填，預設 false；注意欄位英文名稱歷史上曾稱 silicone_grease，中文顯示已統一為「塗矽油」，非「圖矽油」）
silicone_grease_note          塗矽油說明文字（選填）
torque_min                   扭力下限（選填，若填寫會做範圍檢查）
torque_max                   扭力上限（選填，若填寫會做範圍檢查）
notes                        備註（選填）
```

### 4.2 YAML 格式（備用）

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

### 4.3 扭力單位選單固定選項

```text
kgf-cm
N-M
無需扭力
鎖緊就好
```

---

## 5. 快速還原開發狀態的操作步驟

如果需要在新的電腦或重新設定環境時還原開發狀態，依此順序操作：

```powershell
# 1. 確認 Python 已安裝並加入 PATH
python --version

# 2. 進入專案資料夾
cd C:\Users\paes2\OneDrive\桌面\Python\Work\SOP_helper

# 3. 建立虛擬環境
python -m venv .venv

# 4. 若尚未設定過執行原則
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

# 5. 啟用虛擬環境
.\.venv\Scripts\Activate.ps1

# 6. 安裝套件
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

# 7. 執行最新版本程式
python .\SOP_Helper_MVP_v0_9_1.py
```

驗證環境是否正確（非 WSL、非其他 Python 環境）：

```powershell
python -c "import sys; print(sys.executable); print(sys.platform)"
```

預期輸出應包含專案內的 `.venv\Scripts\python.exe` 路徑，且平台為 `win32`。

---

## 6. 尚未使用但已規劃的資源

- **PyInstaller**：規劃用於打包成單一 EXE，目前尚未安裝與測試。
- **GitHub**：已是可用的外部連接工具，規劃用於版本控制與備份，目前程式碼僅存在本機與對話下載檔案，尚未建立正式 repo。
- **pytest**：規劃用於補齊單元測試，目前所有測試皆為人工手動測試。

---

## 7. 與此文件搭配使用的其他文件

- 若要了解**目前程式功能細節與已知限制**，請參考：`SOP_Helper_專案進度與後續計畫_v0_9_1.md`
- 若要了解**完整原始規格與測試框架**，請參考：`SOP_Helper_計畫書與規格_v1.md`
- 本文件專注於**環境、資源與應用背景**，三份文件互補，不重複展開技術細節。
