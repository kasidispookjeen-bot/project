import datetime
import pandas as pd
import plotly.express as px
import streamlit as st

# ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="ระบบตรวจสอบสต็อกรายวัน", page_icon="📦", layout="centered"
)

st.title("📦 ระบบตรวจสอบสต็อกสินค้า (แยกสรุปยอดตามรายวัน)")
st.write(
    "อัปโหลดไฟล์ Excel เพื่อดูสรุปยอดผลิตแยกตามวัน พนักงาน และรุ่นสินค้า (ระบบตัด MASTER และตัวซ้ำให้อัตโนมัติ)"
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

                    # 2. เติมวันที่ในแถวที่ปล่อยเว้นว่างไว้ใน Excel (ลากสูตรลงมา)
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

                    # 🚨 คัดออก! ไม่เอาสถานีที่เป็น 'MASTER' หรือ 'master'
                    df_clean = df_clean[
                        df_clean["STATION"].str.upper() != "MASTER"
                    ]

                    all_file_data.append(
                        df_clean[["DATE_STR", "EMP ID", "STATION", "SN"]]
                    )
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาดในการอ่านไฟล์ {uploaded_file.name}: {e}")

    if all_file_data:
        # รวมข้อมูลดิบจากทุกไฟล์เข้าด้วยกัน
        raw_combined_df = pd.concat(all_file_data, ignore_index=True)

        # 🎯 ตัดรายการที่ "หมายเลขสินค้า (SN)" ซ้ำกันออก ให้เหลือชิ้นเดียว (ป้องกันยอดเบิ้ล)
        combined_df = raw_combined_df.drop_duplicates(subset=["SN"], keep="first")

        if not combined_df.empty:
            # คำนวณสรุปยอดแยกตาม: วันที่ -> พนักงาน -> รุ่นสินค้า (STATION)
            daily_summary_df = (
                combined_df.groupby(["DATE_STR", "EMP ID", "STATION"])["SN"]
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
            # ✨ ส่วนการแสดงผลป๊อปอัปแจ้งเตือนและกล่องเขียวรายวัน
            # --------------------------------------------------
            # 1. ทำป๊อปอัปเด้งเตือน (Toast Notification) ขวาบานล่างของจอ
            popup_message = "🔔 สรุปยอดผลิตรายวันสำเร็จ:\n"
            for _, row in daily_summary_df.iterrows():
                popup_message += f"- วันที่ {row['วันที่']} | พนักงาน {row['รหัสพนักงาน (EMP ID)']} ปรับรุ่น {row['รุ่นสินค้า (STATION)']} = {row['จำนวนที่ผลิตได้ (ตัว)']} ตัว\n"
            st.toast(popup_message, icon="📊")

            # 2. กล่องข้อความสีเขียวแจ้งสรุปผลเด่นๆ ชัดๆ ด้านบนเว็บ
            st.success(
                f"🎉 รวมข้อมูลสำเร็จทั้งหมด {len(uploaded_files)} ไฟล์! (แยกตามวันและรุ่นสินค้าเรียบร้อย)"
            )

            # แสดงลิสต์สรุปรายวันแยกบรรทัดให้อ่านง่ายกระชับ
            for _, row in daily_summary_df.iterrows():
                st.markdown(
                    f"📅 วันที่: **{row['วันที่']}** 👤 พนักงาน: **{row['รหัสพนักงาน (EMP ID)']}** 🛠️ ปรับรุ่น: ` {row['รุ่นสินค้า (STATION)']} ` ➡️ ได้ของแท้จริง **{row['จำนวนที่ผลิตได้ (ตัว)']:,}** ตัว"
                )

            # 3. แสดงกราฟแท่งเปรียบเทียบผลงานรายวันแยกตามรุ่นสินค้า
            # ช่วยให้มองเห็นเลยว่าวันไหนรุ่นไหนได้ยอดเยอะที่สุด
            fig = px.bar(
                daily_summary_df,
                x="วันที่",
                y="จำนวนที่ผลิตได้ (ตัว)",
                color="รุ่นสินค้า (STATION)",
                facet_col="รหัสพนักงาน (EMP ID)",
                barmode="group",
                title="กราฟสรุปยอดผลิตจริงแยกตามวัน พนักงาน และรุ่นสินค้า",
            )
            st.plotly_chart(fig, use_container_width=True)

            # ตารางข้อมูลดิบแบบละเอียดสำหรับกดดาวน์โหลดหรือตรวจสอบซ้ำ
            with st.expander("🔍 คลิกเพื่อดูตารางสรุปข้อมูลทั้งหมด"):
                st.dataframe(
                    daily_summary_df, use_container_width=True, hide_index=True
                )

            # ปุ่มกดบันทึกข้อมูลเข้าคลังประวัติสะสมระยะยาว
            if st.button("📥 บันทึกข้อมูลชุดนี้ลงประวัติยอดรวมสะสม"):
                current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                files_string = ", ".join(file_names_list)

                for _, row in daily_summary_df.iterrows():
                    st.session_state.history_log.append(
                        {
                            "เวลาที่บันทึกระบบ": current_time,
                            "จากไฟล์ทั้งหมด": files_string,
                            "วันที่ทำงาน": row["วันที่"],
                            "รหัสพนักงาน (EMP ID)": row["รหัสพนักงาน (EMP ID)"],
                            "รุ่นสินค้า (STATION)": row["รุ่นสินค้า (STATION)"],
                            "จำนวนรวมสะสม (ตัว)": row["จำนวนที่ผลิตได้ (ตัว)"],
                        }
                    )
                st.toast("บันทึกข้อมูลเข้าคลังประวัติเรียบร้อยแล้ว!")
                st.rerun()
        else:
            st.warning(
                "⚠️ ไม่พบข้อมูลสินค้าตัวอื่นเลยหลังจากคัดสถานี MASTER ออกแล้ว"
            )
    else:
        st.error(
            "❌ ไม่พบโครงสร้างข้อมูลที่ถูกต้อง กรุณาตรวจสอบหัวตารางไฟล์ Excel"
        )

# --------------------------------------------------
# ส่วนที่ 2: หน้าต่างคลังประวัติยอดรวมสะสมย้อนหลัง
# --------------------------------------------------
st.markdown("---")
st.subheader("📜 คลังประวัติยอดรวมสะสมย้อนหลัง (ตัวงานจริงรายวัน)")

if st.session_state.history_log:
    history_df = pd.DataFrame(st.session_state.history_log)
    total_accumulated = history_df["จำนวนรวมสะสม (ตัว)"].sum()

    st.metric(
        "ยอดผลิตสะสมรวมในระบบทั้งหมด (ไม่นับ MASTER)", f"{total_accumulated:,} ตัว"
    )
    st.dataframe(history_df, use_container_width=True, hide_index=True)

    if st.button("🗑️ ล้างประวัติยอดรวมทั้งหมด"):
        st.session_state.history_log = []
        st.rerun()
else:
    st.info("ℹ️ ยังไม่มีประวัติยอดรวมสะสมในคลัง")
