import io
import os
import uuid
from datetime import date, datetime
from typing import Dict, List

import pandas as pd
import streamlit as st
from filelock import FileLock


EXCEL_FILE = "paddy_collection.xlsx"
EXCEL_SHEET = "Submissions"
LOCK_FILE = f"{EXCEL_FILE}.lock"

COLUMN_ORDER: List[str] = [
    "record_id",
    "submitted_at",
    "agent_farmer_phone",
    "agent_farmer_name",
    "truck_number",
    "truck_sheet",
    "sr_number",
    "bags_given",
    "bags_received",
    "weight",
    "mc",
    "date",
]


def clamp_moisture(value: float) -> float:
    if value is None:
        return 0.0
    if value < 0:
        value = 0
    if value > 100:
        value = 100
    return round(float(value), 2)


def ensure_workbook() -> None:
    """Ensure the Excel workbook and sheet exist with correct headers."""
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=COLUMN_ORDER)
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=EXCEL_SHEET, index=False)


def load_dataframe_safe() -> pd.DataFrame:
    """Load the Excel file into a DataFrame, recreating if corrupted/missing."""
    try:
        if not os.path.exists(EXCEL_FILE):
            ensure_workbook()
        df = pd.read_excel(
            EXCEL_FILE,
            sheet_name=EXCEL_SHEET,
            engine="openpyxl",
            dtype={
                "truck_number": str,
                "sr_number": str,
                "agent_farmer_phone": str,
                "agent_farmer_name": str,
                "truck_sheet": str,
                "date": str,
            },
        )
        # Ensure all expected columns exist, in order
        for col in COLUMN_ORDER:
            if col not in df.columns:
                df[col] = pd.Series(dtype="object")
        # Reorder and restrict to known columns
        df = df[COLUMN_ORDER]
        return df
    except Exception:
        # Recreate workbook on error
        df = pd.DataFrame(columns=COLUMN_ORDER)
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=EXCEL_SHEET, index=False)
        return df


def append_to_excel(row_dict: Dict[str, object]) -> None:
    """
    Append a single row to the Excel file in a thread/process-safe way using FileLock.
    Creates the workbook if missing. Preserves column order and types.
    """
    lock = FileLock(LOCK_FILE)
    with lock:
        df = load_dataframe_safe()

        # Coerce dtypes for string-preserved columns
        for col in ["truck_number", "sr_number", "agent_farmer_phone", "agent_farmer_name", "truck_sheet", "date"]:
            if col in row_dict and row_dict[col] is not None:
                row_dict[col] = str(row_dict[col])

        # Build a new row aligned to COLUMN_ORDER
        new_row = {col: row_dict.get(col, None) for col in COLUMN_ORDER}

        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=EXCEL_SHEET, index=False)


def get_excel_bytes() -> bytes:
    """Return the current Excel file content as bytes for download."""
    lock = FileLock(LOCK_FILE)
    with lock:
        if not os.path.exists(EXCEL_FILE):
            ensure_workbook()
        with open(EXCEL_FILE, "rb") as f:
            return f.read()


def make_record(
    agent_farmer_phone: str,
    agent_farmer_name: str,
    truck_number: str,
    truck_sheet: str,
    sr_number: str,
    bags_given: int,
    bags_received: int,
    weight: float,
    mc: float,
    date_str: str,
) -> Dict[str, object]:
    return {
        "record_id": str(uuid.uuid4()),
        "submitted_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "agent_farmer_phone": agent_farmer_phone,
        "agent_farmer_name": agent_farmer_name,
        "truck_number": truck_number,
        "truck_sheet": truck_sheet,
        "sr_number": sr_number,
        "bags_given": int(bags_given),
        "bags_received": int(bags_received),
        "weight": float(weight),
        "mc": clamp_moisture(float(mc)),
        "date": date_str,
    }


def show_help():
    with st.expander("Help"):
        st.write(
            """
            Use this form to record Paddy Collection entries. Fill all fields and submit.
            Each submission is saved to an Excel file named paddy_collection.xlsx in the
            same folder as this app. You can preview the latest 20 entries below and
            download the full Excel file with the button provided.
            """
        )


def validate_inputs(values: Dict[str, object]) -> List[str]:
    errors: List[str] = []

    phone = str(values.get("agent_farmer_phone", "")).strip()
    if not phone.isdigit() or not (7 <= len(phone) <= 15):
        errors.append("Phone must be digits only, length 7–15.")

    name = str(values.get("agent_farmer_name", "")).strip()
    if not name:
        errors.append("Name is required.")

    truck_number = str(values.get("truck_number", "")).strip()
    if not truck_number:
        errors.append("Truck Number is required.")

    truck_sheet = str(values.get("truck_sheet", "")).strip()
    if not truck_sheet:
        errors.append("Truck Sheet is required.")

    sr_number = str(values.get("sr_number", "")).strip()
    if not sr_number:
        errors.append("SR Number is required.")

    try:
        bags_given = int(values.get("bags_given", 0))
        if bags_given < 0:
            errors.append("Number of Bags Given must be ≥ 0.")
    except Exception:
        errors.append("Number of Bags Given must be an integer.")

    try:
        bags_received = int(values.get("bags_received", 0))
        if bags_received < 0:
            errors.append("Bags Received must be ≥ 0.")
    except Exception:
        errors.append("Bags Received must be an integer.")

    try:
        weight = float(values.get("weight", 0))
        if weight < 0:
            errors.append("Weight must be ≥ 0.")
    except Exception:
        errors.append("Weight must be a number.")

    try:
        mc = float(values.get("mc", 0))
        if mc < 0 or mc > 100:
            errors.append("MC must be between 0 and 100.")
    except Exception:
        errors.append("MC must be a number.")

    date_val = values.get("date")
    if not isinstance(date_val, date):
        errors.append("Date is required.")

    return errors


def main():
    st.set_page_config(page_title="Paddy Collection", layout="centered")
    st.title("Paddy Collection")

    # Ensure workbook exists early
    try:
        ensure_workbook()
    except Exception:
        st.warning("Workbook was missing or corrupted and has been reinitialized.")

    show_help()

    # Session state key to trigger form reset
    if "form_reset_token" not in st.session_state:
        st.session_state.form_reset_token = str(uuid.uuid4())

    with st.form(key=f"paddy_form_{st.session_state.form_reset_token}"):
        col1, col2 = st.columns(2, gap="small")

        with col1:
            agent_farmer_phone = st.text_input("Agent/Farmer Phone Number", placeholder="Digits only", max_chars=15)
            agent_farmer_name = st.text_input("Agent/Farmer Name")
            truck_number = st.text_input("Truck Number")
            truck_sheet = st.text_input("Truck Sheet")
            sr_number = st.text_input("SR Number")

        with col2:
            bags_given = st.number_input("Number of Bags Given", min_value=0, step=1, value=0)
            bags_received = st.number_input("Bags Received", min_value=0, step=1, value=0)
            weight = st.number_input("Weight (kg)", min_value=0.0, step=0.1, format="%0.3f")
            mc = st.number_input("MC (Moisture %)", min_value=0.0, max_value=100.0, step=0.1, format="%0.2f")
            date_val = st.date_input("Date", value=date.today())

        submitted = st.form_submit_button("Submit")

    if submitted:
        values = {
            "agent_farmer_phone": agent_farmer_phone.strip(),
            "agent_farmer_name": agent_farmer_name.strip(),
            "truck_number": truck_number.strip(),
            "truck_sheet": truck_sheet.strip(),
            "sr_number": sr_number.strip(),
            "bags_given": bags_given,
            "bags_received": bags_received,
            "weight": weight,
            "mc": mc,
            "date": date_val,
        }

        errors = validate_inputs(values)
        if errors:
            for e in errors:
                st.error(e)
        else:
            record = make_record(
                agent_farmer_phone=values["agent_farmer_phone"],
                agent_farmer_name=values["agent_farmer_name"],
                truck_number=values["truck_number"],
                truck_sheet=values["truck_sheet"],
                sr_number=values["sr_number"],
                bags_given=int(values["bags_given"]),
                bags_received=int(values["bags_received"]),
                weight=float(values["weight"]),
                mc=float(values["mc"]),
                date_str=values["date"].strftime("%Y-%m-%d"),
            )
            try:
                append_to_excel(record)
                st.success(f"Saved. Record ID: {record['record_id']}")
                # Reset form by changing the key token
                st.session_state.form_reset_token = str(uuid.uuid4())
                st.experimental_rerun()
            except Exception as exc:
                st.error(f"Failed to save record: {exc}")

    st.subheader("Recent Submissions")
    try:
        df = load_dataframe_safe()
        if not df.empty:
            # Sort by submitted_at descending and show last 20 (most recent first)
            if "submitted_at" in df.columns:
                with pd.option_context("mode.use_inf_as_na", True):
                    df_sorted = df.copy()
                    df_sorted["_submitted_at_sort"] = pd.to_datetime(df_sorted["submitted_at"], errors="coerce")
                    df_sorted = df_sorted.sort_values("_submitted_at_sort", ascending=False).drop(columns=["_submitted_at_sort"])
            else:
                df_sorted = df
            st.dataframe(df_sorted.head(20), use_container_width=True)
        else:
            st.info("No submissions yet.")
    except Exception as exc:
        st.warning(f"Could not load preview: {exc}")

    # Download button
    try:
        excel_bytes = get_excel_bytes()
        st.download_button(
            label="Download Excel",
            data=excel_bytes,
            file_name=EXCEL_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        st.warning(f"Download unavailable: {exc}")


if __name__ == "__main__":
    main()

