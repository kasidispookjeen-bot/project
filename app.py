import pandas as pd
import streamlit as st

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="ระบบอ่านไฟล์นับสต็อก", page_icon="📦", layout="centered")

st.title("📦 ระบบตรวจสอบข้อมูลการนับสต็อกสินค้า")
st.write("อัปโหลดไฟล์ Excel เพื่อดูสรุปรายการนับสต็อกของพนักงาน")

# ส่วนอัปโหลดไฟล์
uploaded_file = st.file_uploader(
    "เลือกไฟล์ Excel ของคุณ (เช่น 5-25.xlsx)", type=["xlsx", "xls"]
)

if uploaded_file is not None:
    try:
        # อ่านไฟล์ Excel เพื่อดูชีตทั้งหมด
        xls = pd.ExcelFile(uploaded_file)
        all_data = []

        # ดึงข้อมูลจากทุกชีตที่มีโครงสร้างถูกต้อง
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

            # ตรวจสอบคอลัมน์ที่จำเป็น
            required_cols = ["EMP ID", "STATION", "SN"]
            if all(col in df.columns for col in required_cols):
                # คลีนข้อมูลแถวว่างและหัวตารางซ้ำซ้อน
                df_clean = df[df["EMP ID"].notna()]
                df_clean = df_clean[df_clean["EMP ID"] != "EMP ID"]

                # เอาช่องว่าง (space) ออกจากข้อความ
                df_clean["EMP ID"] = df_clean["EMP ID"].astype(str).str.strip()
                df_clean["STATION"] = (
                    df_clean["STATION"].astype(str).str.strip()
                )

                all_data.append(df_clean[["EMP ID", "STATION", "SN"]])

        if all_data:
            # รวมข้อมูลจากทุกชีตเข้าด้วยกัน
            combined_df = pd.concat(all_data, ignore_index=True)

            # 1. หน้าสรุปภาพรวม (Summary)
            st.subheader("📊 สรุปจำนวนการนับสต็อกแยกตามพนักงาน")
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

            # แสดงตารางสรุปบนเว็บ
            st.dataframe(summary_df, use_container_width=True, hide_index=True)

            # 2. ตัวกรองข้อมูลเพื่อค้นหาแบบละเอียด
            st.markdown("---")
            st.subheader("🔍 ค้นหารายละเอียดรายบุคคล")

            unique_emp = combined_df["EMP ID"].unique()
            selected_emp = st.selectbox("เลือกรหัสพนักงานที่ต้องการตรวจสอบ", unique_emp)

            if selected_emp:
                # กรองข้อมูลเฉพาะของพนักงานคนนั้น
                emp_details = combined_df[combined_df["EMP ID"] == selected_emp]
                st.write(
                    f"**รายการสินค้าทั้งหมดที่รหัสพนักงาน `{selected_emp}` เป็นคนนับ (รวม {len(emp_details)} ตัว):**"
                )

                # แสดงตารางรายชื่อ Serial Number ทั้งหมด
                emp_details_display = emp_details[["STATION", "SN"]].rename(
                    columns={
                        "STATION": "สินค้า/สถานี (STATION)",
                        "SN": "หมายเลขโปรดัค (SN)",
                    }
                )
                st.dataframe(
                    emp_details_display,
                    use_container_width=True,
                    hide_index=True,
                )

        else:
            st.error(
                "❌ ไม่พบคอลัมน์ 'EMP ID', 'STATION' หรือ 'SN' ในไฟล์นี้ กรุณาตรวจสอบโครงสร้างตาราง"
            )

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์: {e}")
