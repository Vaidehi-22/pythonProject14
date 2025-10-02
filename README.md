# 🏢 Multi-Equipment Final Nomenclature Generator

This project is a **Streamlit web application** that generates final equipment nomenclatures based on **Planon data** and **System workbooks** (Tag Summary + Equipment sheets).  
It is designed for standardizing asset naming across multiple equipment, ensuring consistency with location, building, floor, room, and sensor tags.

---

## 🚀 Features

- **Upload Inputs**
  - Planon Excel (must contain: `Location code`, `Building code`, `Floor Code`, `Room code`)
  - System Workbook Excel (must contain: *Tag Summary* sheet and equipment sheets)

- **Dependent Dropdowns**
  - Select `Location code` → `Building code` → `Floor code` → `Room code`

- **Automatic Term–Abbreviation Extraction**
  - From *Tag Summary* sheet → extracts equipment **Terms ↔ Abbreviations**
  - From each equipment sheet → extracts **Name ↔ Abbreviation**

- **User Inputs**
  - Multi-select equipment terms
  - Enter asset numbers for each selected equipment

- **Final Nomenclature Generation**
  - Constructs nomenclature using:
    ```
    <Location_Prefix>_<Site>_<BuildingCodeAfterHyphen>_<Floor>_<EquipAbbrev+AssetNumber>_<RoomNumeric>_<SensorAbbrevNoDigits>
    ```
  - Example:
    ```
    AE_ABUS2_ABUS2-01_01_Ch_101_TInCh
    ```

- **Cleaning Rules**
  1. **Room Code** → Keeps only numeric and dot parts  
     - `"ABUS2.01"` → `"01"`  
     - `"1F 1.1"` → `"1.1"`
  2. **Sensor Name / Tag Abbreviation** → Removes numbers  
     - `"TInCh1"` → `"TInCh"`

- **Downloadable Output**
  - Generates a downloadable Excel file directly in the UI (in memory only, not stored in repo).
  - Inserts a **note row** and an **Excel comment**:
    > ⚠️ Room codes are trimmed to numeric parts only; sensor names ignore numeric values.

---

## 🖥️ How to Run

1. Clone the repo:
   ```bash
   git clone https://github.com/Vaidehi-22/pythonProject14.git
   cd pythonProject14
