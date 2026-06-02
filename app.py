import datetime
import pandas as pd
import plotly.express as px
import streamlit as st

# ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="ระบบสต็อกอัจฉริยะ แบบ 2 โหมด", page_icon="📦", layout="centered"
)

st.title("📦 ระบบ ตรวจสอบสต็อกสินค้า (Multi-Mode)")
st.write(
    "อัปโหลดไฟล์ Excel เพื่อดูสรุปยอดผลิตรายวัน (ระบบตัดรายการซ้ำและรวมรหัสพนักงานที่พิมพ์ผิดให้อัตโนมัติ)"
)

# --------------------------------------------------
# ระบบฐานข้อมูลจำลอง (ประวัติยอดรวมสะสม)
# --------------------------------------------------
if "history_log" not in st.session_state:
    st.session_state.history_log = []

# ส่วนอัปโหลดไฟล์
uploaded_files = st.file_uploader(
    "เลือกหรือลากไฟล์ Excel ของคุณมาวางที่นี่ (เลือกได้หลายไฟล์พร้อมกัน)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if uploaded_files:
    all_file_data = []
    file_names_list = []

    for uploaded_file in uploaded_files:
        try:
            xls = pd.ExcelFile(uploaded_file)
            file_names_list.append(uploaded_file.name)

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

                # ตรวจสอบโครงสร้างคอลัมน์สำคัญ
                if (
                    "DATE" in df.columns
                    and "EMP ID" in df.columns
                    and "STATION" in df.columns
                    and "SN" in df.columns
                ):
                    # 1. คลีนข้อมูลเบื้องต้น ลบแถวว่างและแถวหัวตารางที่เกินมาออก
                    df_clean = df[df["EMP ID"].notna()]
                    df_clean = df_clean[df_clean["EMP ID"] != "EMP ID"]

                    # 2. เติมวันที่ในแถวที่ปล่อยเว้นว่างไว้ (ลากสูตรลงมา)
                    df_clean["DATE"] = df_clean["DATE"].ffill()

                    # 3. ตัดช่องว่าง (Space) หน้า-หลังข้อความทั้งหมด
                    df_clean["EMP ID"] = (
                        df_clean["EMP ID"].astype(str).str.strip()
                    )
                    df_clean["STATION"] = (
                        df_clean["STATION"].astype(str).str.strip()
                    )
                    df_clean["SN"] = df_clean["SN"].astype(str).str.strip()
                    df_clean["DATE_STR"] = (
                        df_clean["DATE"].astype(str).str.strip()
                    )

                    # 🚨 แก้ปัญหาพนักงานพิมพ์รหัสสลับ ตัว O กับ เลข 0
                    df_clean["EMP ID"] = df_clean["EMP ID"].str.replace(
                        "O", "0", case=False
                    )

                    all_file_data.append(
                        df_clean[["DATE_STR", "EMP ID", "STATION", "SN"]]
                    )
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาดในการอ่านไฟล์ {uploaded_file.name}: {e}")

    if all_file_data:
        # รวมข้อมูลดิบทั้งหมดจากทุกไฟล์
        raw_combined_df = pd.concat(all_file_data, ignore_index=True)

        # 🎯 ตัดรายการที่ "หมายเลขสินค้า (SN)" ซ้ำกันออก ให้เหลือชิ้นเดียว (ป้องกันยอดเบิ้ล)
        combined_df = raw_combined_df.drop_duplicates(subset=["SN"], keep="first")

        # --------------------------------------------------
        # ✨ เพิ่มระบบปุ่มกดเลือกโหมด (2 โหมด)
        # --------------------------------------------------
        st.markdown("---")
        st.subheader("⚙️ เลือกโหมดการแสดงผลข้อมูล")
        view_mode = st.radio(
            "คุณต้องการดูข้อมูลในรูปแบบใด?",
            (
                "🟢 เฉพาะสินค้าจริง (ไม่นับ MASTER)",
                "🔵 รวมยอดทั้งหมด (นับรวมสถานี MASTER)",
            ),
            horizontal=True,
        )

        # ทำการกรองข้อมูลตามโหมดที่เลือก
        if "เฉพาะสินค้าจริง" in view_mode:
            # คัดตัวที่เป็น MASTER ออกไปให้หมด
            display_df = combined_df[combined_df["STATION"].str.upper() != "MASTER"]
            mode_title = "เฉพาะสินค้าจริง (ไม่รวม MASTER)"
        else:
            # ปล่อยไว้ปกติ ให้ดึงทุกสถานีรวมถึง MASTER ด้วย
            display_df = combined_df
            mode_title = "รวมยอดทั้งหมด (รวม MASTER)"

        if not display_df.empty:
            # คำนวณสรุปยอดแยกตาม: วันที่ -> พนักงาน -> รุ่นสินค้า (STATION)
            daily_summary_df = (
                display_df.groupby(["DATE_STR", "EMP ID", "STATION"])["SN"]
                .count()
                .reset_index()
            )
            daily_summary_df.columns = [
                "วันที่",
                "รหัสพนักงาน (EMP ID)",
                "รุ่นสินค้า (STATION)",
                "จำนวนที่ผลิตได้ (ตัว)",
            ]

            # --------------------------------------------------
            # แสดงผลลัพธ์ตามโหมดที่เลือก
            # --------------------------------------------------
            st.success(f"🎉 แสดงผลข้อมูลในโหมด: **{mode_title}** สำเร็จ!")

            # แสดงลิสต์สรุปรายวันแยกบรรทัด
            for _, row in daily_summary_df.iterrows():
                st.markdown(
                    f"📅 วันที่: **{row['วันที่']}** | 👤 พนักงาน: **{row['รหัสพนักงาน (EMP ID)']}** | 🛠️ รุ่น: ` {row['รุ่นสินค้า (STATION)']} ` ➡️ จำนวน: **{row['จำนวนที่ผลิตได้ (ตัว)']:,}** ตัว"
                )

            # แสดงกราฟแท่งตามโหมดนั้นๆ
            fig = px.bar(
                daily_summary_df,
                x="วันที่",
                y="จำนวนที่ผลิตได้ (ตัว)",
                color="รุ่นสินค้า (STATION)",
                facet_col="รหัสพนักงาน (EMP ID)",
                barmode="group",
                title=f"กราฟสรุปยอดผลิตจริง ({mode_title})",
            )
            st.plotly_chart(fig, use_container_width=True)

            # ตารางข้อมูลสรุปใต้กราฟ
            with st.expander("🔍 คลิกเพื่อดูตารางสรุปข้อมูลโหมดนี้แบบทั้งหมด"):
                st.dataframe(
                    daily_summary_df, use_container_width=True, hide_index=True
                )

            # ปุ่มกดบันทึกข้อมูลเข้าคลังประวัติสะสม
            if st.button("📥 บันทึกข้อมูลในโหมดนี้ลงประวัติยอดรวมสะสม"):
                current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                files_string = ", ".join(file_names_list)

                for _, row in daily_summary_df.iterrows():
                    st.session_state.history_log.append(
                        {
                            "เวลาที่บันทึกระบบ": current_time,
                            "จากไฟล์ทั้งหมด": files_string,
                            "โหมดที่บันทึก": mode_title,
                            "วันที่ทำงาน": row["วันที่"],
                            "รหัสพนักงาน (EMP ID)": row["รหัสพนักงาน (EMP ID)"],
                            "รุ่นสินค้า (STATION)": row["รุ่นสินค้า (STATION)"],
                            "จำนวนรวมสะสม (ตัว)": row["จำนวนที่ผลิตได้ (ตัว)"],
                        }
                    )
                st.toast("บันทึกข้อมูลเข้าคลังประวัติเรียบร้อยแล้ว!")
                st.rerun()
        else:
            st.warning("⚠️ ไม่พบข้อมูลที่จะแสดงในโหมดนี้")
    else:
        st.error(
            "❌ ไม่พบโครงสร้างข้อมูลที่ถูกต้อง กรุณาตรวจสอบหัวตารางไฟล์ Excel"
        )

# --------------------------------------------------
# ส่วนที่ 2: หน้าต่างคลังประวัติยอดรวมสะสมย้อนหลัง
# --------------------------------------------------
st.markdown("---")
st.subheader("📜 คลังประวัติยอดรวมสะสมย้อนหลัง")

if st.session_state.history_log:
    history_df = pd.DataFrame(st.session_state.history_log)
    total_accumulated = history_df["จำนวนรวมสะสม (ตัว)"].sum()

    st.metric("ยอดนับสินค้าสะสมรวมในระบบทั้งหมด", f"{total_accumulated:,} ตัว")
    st.dataframe(history_df, use_container_width=True, hide_index=True)

    if st.button("🗑️ ล้างประวัติยอดรวมทั้งหมด"):
        st.session_state.history_log = []
        st.rerun()
else:
    st.info("ℹ️ ยังไม่มีประวัติยอดรวมสะสมในคลัง")
