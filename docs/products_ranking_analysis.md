# Project Analysis: Products Ranking Automation

This document provides a detailed analysis of the Python project `products_ranking`. The analysis covers the role of each file, detailed implementation logic, and the relationships between modules.

## 1. File Roles and Detailed Implementation Logic

### 1.1 `config.py`
**Role**: Central Configuration
**Description**: Stores global constants, file paths, and account information. It ensures the project runs correctly both as a raw script and as a PyInstaller-frozen executable.

**Key Logic**:
- **Path Resolution**: Uses `sys.executable` if frozen, else `__file__` to determine the `BASE_DIR`.
- **Constants**:
    - `EXCEL_PATH`: `TOP★점프_트래픽관리.xlsm`
    - `TOP_ADS_URL`: `https://top.re.kr/ads`
    - `MAX_PAGES`: 5 (Limits scraping depth)
    - `ACCOUNT`: Hardcoded credentials (`sstrade251016`).

### 1.2 `excel_handler.py`
**Role**: Excel Interaction Manager
**Description**: Handles reading constraints from Excel and writing back scraping results. It relies primarily on `xlwings` to interact with the active Excel application instance, preserving Macros (`.xlsm`) and external links.

**Detailed Logic**:
- **`sync_date_columns_until_today(ws)`**:
    - Iterates from `2026-01-01` to Today.
    - Checks row 5 for existing date headers.
    - **Logic**: If a date column is missing, it dynamically **inserts a new column** to the left of specific marker columns ("직전", "비고", "서식") or at the end. This ensures the sheet grows horizontally without breaking fixed formulas on the far right.
- **`get_dates_requiring_update(ws)`**:
    - **Logic Characteristic (All-or-Nothing)**: It checks each date column. A date is added to the "update list" **only if check rows (marked "순위" in Col Q) are completely empty**.
    - *Note*: If even one product has a rank recorded for a specific date, that date is skipped entirely. This is an important optimization but assumes interrupted runs are manually handled or rare.
- **`update_excel_rank_fast(ws, ...)`**:
    - Performs O(1) updates using a pre-calculated index.
    - Updates the cell at `(row_num, target_col)` with the new rank.
- **`build_row_index(ws)`**:
    - Scans columns F (Product ID) and J (Keyword).
    - Map: `{(Product_ID, Keyword) -> Row_Number}`.
    - This preprocessing step dramatically speeds up the writing phase compared to searching row-by-row for every result.

### 1.3 `web_handler.py`
**Role**: Web Automation (Selenium)
**Description**: Manages the Chrome browser lifecycle, traversing the `top.re.kr` website, and extracting ranking data.

**Detailed Logic**:
- **`create_driver`**: Initializes Chrome with `automationControlled` disabled to evade basic bot detection. Uses a persistent user profile directory.
- **`extract_normal_products(driver, target_dates)`**:
    - **Filtering Logic**: Navigates to "Normal" (정상) tab and currently active products.
    - **Date Dependency**: It calls `extract_product_results` internally. This means it only discovers "Normal" products that *also* match the `target_dates` criteria (i.e., active within the target date range).
- **`extract_product_results(driver, target_dates)`**:
    - Scrapes the 1000-row table.
    - **Date Filtering**: For each row, parses "Start Date" and "End Date".
    - **Match Logic**: Checks `Start Date <= Target Date <= End Date`. Only creates a record if the product is active on the specific target date.
    - Extracts: `Row Keyword`, `Product ID` (from URL), and `Rank`.
- **`login_success_check`**: Implements a robust login flow with up to 3 retries and validation via "Logout" button visibility.

### 1.4 `main.py`
**Role**: Orchestrator
**Description**: Ties the Excel preparation, Web scraping, and Excel updating phases together.

**Execution Flow**:
1.  **Excel Prep**: Opens Excel, syncs date columns, and identifies which dates have zero data.
    - *Condition*: If no dates need updates, the program terminates immediately.
2.  **Environment Prep**: Kills stale Chrome processes and wipes specific cache folders (`Default/Cache`, `Code Cache`, etc.) to ensure a fresh browser state.
3.  **Discovery**: Logs in and scrapes currently active "Normal" products to build a list of target `Product IDs`.
4.  **Targeted Scraping (Loop)**:
    - For each unique `Product ID` found:
        - Helper: `search_keyword(ID)` resets filters and searches for that specific ID.
        - Action: `extract_product_results` gets all ranking lines for that ID (checking against target dates).
        - **Write**: Immediately updates the in-memory Excel object (via `xlwings`) using the fast index lookup.
5.  **Commit**: Saves the Excel file (`wb.save()`) only after all loops complete.
6.  **Cleanup**: Closes Excel and Chrome driver.

## 2. Logic Flow & Relationship Diagram

The system follows a linear batch-processing model.

```mermaid
flowchart TD
    subgraph Initialization
        A[main.py Start] --> B[config.py: Load Config]
        A --> C[excel_handler: Open .xlsm]
        C --> D[excel_handler: Sync Date Columns]
        D --> E{Dates need update?}
        E -- No --> F[End Program]
    end

    subgraph Preparation
        E -- Yes --> G[excel_handler: Build Row Index]
        G --> H[web_handler: Clear Cache & Kill Chrome]
        H --> I[web_handler: Launch Chrome & Login]
    end

    subgraph Discovery
        I --> J[web_handler: Extract 'Normal' Products]
        J --> K[Identify Unique Product IDs]
    end

    subgraph Scraping_Loop [For Each Product ID]
        K --> L[web_handler: Search by ID]
        L --> M[web_handler: Extract Ranks (Date Filtered)]
        M --> N[excel_handler: Update Cell (Fast Index)]
    end

    subgraph Finalization
        N --> O_Loop{More IDs?}
        O_Loop -- Yes --> L
        O_Loop -- No --> P[excel_handler: Save File]
        P --> Q[Close Resources]
    end
```

### Key Dependencies
- **`main.py`** is the **Controller**: It imports and directs all other modules.
- **`web_handler.py`** is the **Provider**: It provides data (rankings) but doesn't know about Excel structure.
- **`excel_handler.py`** is the **Storage**: It knows the Excel structure (rows, columns, dates) but doesn't know about the web source.
- **`config.py`** is the **Shared Truth**: Both handlers use it for paths and basic settings.

### Critical Logic Observations
1.  **Partial Update Risk**: If the script crashes mid-loop, some products for a date might be updated while others aren't. Because `get_dates_requiring_update` skips a date if *any* data exists, rerunning the script might skip the partially-filled date, leaving some products un-updated for that day.
2.  **Date-Centric Filtering**: Products are only tracked if their "Advertisement Period" (Start~End) covers the target date. Expired products are automatically ignored by `extract_product_results`.
