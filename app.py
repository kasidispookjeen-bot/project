import datetime
import pandas as pd
import streamlit as st

# ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="ระบบอ่านไฟล์นับสต็อกและบันทึกประวัติ", page_icon="📦", layout="centered"
)

st.title("📦 ระบบตรวจสอบและบันทึกประวัติยอดรวมสต็อก")
st.write(
    "อัปโหลดไฟล์ Excel เพื่อดูสรุป และระบบจะทำการบันทึกประวัติยอดรวมเก็บไว้ให้คุณโดยอัตโนมัติ"
)

# --------------------------------------------------
# ระบบฐานข้อมูลจำลอง (ประวัติยอดรวมสะสม)
# --------------------------------------------------
# สร้างตัวแปรเก็บประวัติในระบบ (จะอยู่ตลอดตราบใดที่ยังไม่ปิดหน้าเว็บหรือรีเฟรช)
if "history_log" not in st.session_state:
    st.session_state.history_log = []

# ส่วนอัปโหลดไฟล์
uploaded_file = st.file_uploader(
    "เลือกไฟล์ Excel ของคุณ (เช่น 5-25.xlsx)", type=["xlsx", "xls"]
)

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        all_data = []

        # ดึงข้อมูลจากทุกชีต
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

            required_cols = ["EMP ID", "STATION", "SN"]
            if all(col in df.columns for col in required_cols):
                df_clean = df[df["EMP ID"].notna()]
                df_clean = df_clean[df_clean["EMP ID"] != "EMP ID"]

                df_clean["EMP ID"] = df_clean["EMP ID"].astype(str).str.strip()
                df_clean["STATION"] = (
                    df_clean["STATION"].astype(str).str.strip()
                )

                all_data.append(df_clean[["EMP ID", "STATION", "SN"]])

        if all_data:
            combined_df = pd.concat(all_data, ignore_index=True)

            # คำนวณสรุปของไฟล์ปัจจุบัน
            summary_df = (
                combined_df.groupby(["EMP ID", "STATION"])["SN"]
                .count()
                .reset_index()
            )
            summary_df.columns = [
                "รหัสพนักงาน (EMP ID)",
                "สินค้า/สถานี (STATION)",
                "จำนวนที่นับได้ (ตัว)",
            ]

            # --------------------------------------------------
            # ส่วนของปุ่ม "บันทึกเข้าประวัติยอดรวม"
            # --------------------------------------------------
            st.success("✅ อ่านไฟล์สำเร็จ!")

            # แสดงตารางผลลัพธ์ของไฟล์ที่เพิ่งอัปโหลด
            st.subheader("📊 ผลการนับสต็อกจากไฟล์ปัจจุบัน")
            st.dataframe(summary_df, use_container_width=True, hide_index=True)

            # คำนวณยอดรวมของไฟล์นี้ เพื่อเตรียมบันทึก
            total_items = len(combined_df)
            unique_staff = combined_df["EMP ID"].nunique()

            st.write(
                f"💡 ไฟล์นี้มีพนักงานทำงานทั้งหมด **{unique_staff} คน** รวมยอดนับสินค้าได้ **{total_items} ตัว**"
            )

            # สร้างปุ่มกดบันทึกข้อมูลเข้าประวัติรวม
            if st.button("📥 บันทึกข้อมูลไฟล์นี้ลงประวัติยอดรวมสะสม"):
                # ตรวจสอบเพื่อไม่ให้บันทึกไฟล์ซ้ำซ้อนกัน
                file_name = uploaded_file.name
                current_time = datetime.datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )

                # ดึงข้อมูลรายคนไปบันทึกในประวัติ
                for _, row in summary_df.iterrows():
                    st.session_state.history_log.append(
                        {
                            "เวลาที่บันทึก": current_time,
                            "ชื่อไฟล์": file_name,
                            "รหัสพนักงาน (EMP ID)": row[
                                "รหัสพนักงาน (EMP ID)"
                            ],
                            "สินค้า/สถานี (STATION)": row[
                                "สินค้า/สถานี (STATION)"
                            ],
                            "จำนวนที่นับได้ (ตัว)": row[
                                "จำนวนที่นับได้ (ตัว)"
                            ],
                        }
                    )
                st.toast(f"บันทึกประวัติของไฟล์ {file_name} เรียบร้อยแล้ว!")

        else:
            st.error(
                "❌ ไม่พบคอลัมน์ 'EMP ID', 'STATION' หรือ 'SN' ในไฟล์นี้"
            )

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์: {e}")

# --------------------------------------------------
# ส่วนที่ 2: หน้าต่างประวัติยอดรวมสะสม (Historical Dashboard)
# --------------------------------------------------
st.markdown("---")
st.subheader("📜 ประวัติยอดรวมสะสม (ข้อมูลทั้งหมดที่บันทึกไว้)")

if st.session_state.history_log:
    # แปลงข้อมูลประวัติในระบบออกมาเป็นตาราง DataFrame
    history_df = pd.DataFrame(st.session_state.history_log)

    # 1. แสดงกล่องสถิติตัวเลขยอดรวม (KPI Cards)
    total_accumulated_count = history_df["จำนวนที่นับได้ (ตัว)"].sum()
    total_files_uploaded = history_df["ชื่อไฟล์"].nunique()

    col1, col2 = st.columns(2)
    with col1:
        st.metric("จำนวนไฟล์ที่บันทึกไปแล้ว", f"{total_files_uploaded} ไฟล์")
    with col2:
        st.metric(
            "ยอดนับสินค้าสะสมรวมทั้งหมด", f"{total_accumulated_count:,} ตัว"
        )

    # 2. ตารางแสดงประวัติรวมแบบละเอียด
    st.write("**ตารางประวัติการบันทึกข้อมูลย้อนหลัง:**")
    st.dataframe(history_df, use_container_width=True, hide_index=True)

    # 3. ปุ่มสำหรับล้างประวัติ (Reset)
    if st.button("🗑️ ล้างประวัติยอดรวมทั้งหมด"):
        st.session_state.history_log = []
        st.rerun()
else:
    st.info(
        "ℹ️ ยังไม่มีประวัติยอดรวมที่บันทึกไว้ กรุณาอัปโหลดไฟล์แล้วกดปุ่มบันทึกด้านบน"
    )
