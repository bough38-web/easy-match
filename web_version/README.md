# ExcelMatcher Web Version (Prototype)

This is a web-based version of the ExcelMatcher application using FastAPI.

## Setup

1.  **Install Requirements**:
    ```bash
    pip install -r requirements.txt
    ```
    (Note: `fastapi`, `uvicorn`, `python-multipart`, `pandas`, `openpyxl`, `python-calamine` are required)

## Run

1.  Navigate to this folder in your terminal.
2.  Run the server:
    ```bash
    uvicorn server:app --reload
    ```
    OR using python directly:
    ```bash
    python3 server.py
    ```

3.  Open your browser and go to:
    [http://localhost:8000](http://localhost:8000)

## Usage

1.  **Select Files**:
    - Choose a **Base File** and a **Target File**.
    - The system will automatically analyze the files.

2.  **Select Sheets**:
    - If the file has multiple sheets, select the one you want to match.

3.  **Select Columns**:
    - **Key Columns**: Select the columns that are common to both files for matching.
    - **Take Columns**: Select the columns you want to bring from the Target file.

4.  **Match**:
    - Click **Match**.
    - The result file will be downloaded automatically.
