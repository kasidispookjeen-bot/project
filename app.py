import datetime
import pandas as pd
import plotly.express as px
import streamlit as st

# ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="ระบบสต็อกอัจฉริยะ (Multi-Mode)", page_icon="📦", layout="centered"
)

st.title("📦 ระบบตรวจสอบและบันทึกประวัติสต็อก")
st.write(
    "ระบบเวอร์ชันแก้ปัญหาทักษะการพิมพ์: เปลี่ยนตัวอักษร `O` ในรหัสพนักงานให้เป็นเลข `0` อัตโนมัติ, ตัดตัวซ้ำ, และเลือกสลับโหมดดู MASTER ได้"
)

# --------------------------------------------------
# ระบบฐานข้อมูลจำลอง (ประวัติยอดรวมสะสม)
# --------------------------------------------------
if "history_log" not in st.session_state:
    st.session_state.history_log = []

# ส่วนอัปโหลดไฟล์
uploaded_files = st.file_uploader(
    "เลือกหรือลากไฟล์ Excel ของคุณมาวางที่นี่ (เลือกได้หลายไฟล์)",
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
                    "EMP ID" in df.columns
                    and "STATION" in df.columns
                    and "SN" in df.columns
                ):
                    # 1. คลีนข้อมูลเบื้องต้น ลบแถวว่างและแถวหัวตารางที่เกินมาออก
                    df_clean = df[df["EMP ID"].notna()]
                    df_clean = df_clean[df_clean["EMP ID"] != "EMP ID"]

                    # 2. ตัดช่องว่าง (Space) หน้า-หลังข้อความทั้งหมด
                    df_clean["EMP ID"] = (
                        df_clean["EMP ID"].astype(str).str.strip()
                    )
                    df_clean["STATION"] = (
                        df_clean["STATION"].astype(str).str.strip()
                    )
                    df_clean["SN"] = df_clean["SN"].astype(str).str.strip()

                    # 🚨 [จุดแก้ไขสำคัญ] แก้ปัญหาพิมพ์ รหัสพนักงานสลับ เลข 0 กับ ตัว O
                    df_clean["EMP ID"] = (
                        df_clean["EMP ID"]
                        .str.replace("O", "0", case=False)
                    )

                    all_file_data.append(df_clean[["EMP ID", "STATION", "SN"]])
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาดในการอ่านไฟล์ {uploaded_file.name}: {e}")

    if all_file_data:
        # รวมข้อมูลดิบจากทุกไฟล์เข้าด้วยกัน
        raw_combined_df = pd.concat(all_file_data, ignore_index=True)

        # 🎯 ตัดรายการที่ "หมายเลขสินค้า (SN)" ซ้ำกันออก ให้เหลือชิ้นเดียว
        combined_df = raw_combined_df.drop_duplicates(subset=["SN"], keep="first")

        # --------------------------------------------------
        # ✨ เพิ่มปุ่มตัวเลือก 2 โหมด (รวม / ไม่รวม MASTER)
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

        # กรองแยกข้อมูลตามโหมดที่เลือกกด
        if "เฉพาะสินค้าจริง" in view_mode:
            display_df = combined_df[combined_df["STATION"].str.upper() != "MASTER"]
            mode_title = "เฉพาะสินค้าจริง (ไม่รวม MASTER)"
            mode_icon = "🟢"
        else:
            display_df = combined_df
            mode_title = "รวมยอดทั้งหมด (รวม MASTER)"
            mode_icon = "🔵"

        if not display_df.empty:
            # คำนวณสรุปยอดรวมสุทธิแยกรายบุคคลตามโหมดที่เลือก
            emp_total_df = (
                display_df.groupby("EMP ID")["SN"].count().reset_index()
            )
            emp_total_df.columns = ["รหัสพนักงาน (EMP ID)", "จำนวนรวมแท้จริง (ตัว)"]

            # --------------------------------------------------
            # ✨ ส่วนการแสดงผลป๊อปอัปและข้อความสรุปรายคน
            # --------------------------------------------------
            popup_message = f"🔔 สรุปผลยอดนับ ({mode_title}): \n"
            for _, row in emp_total_df.iterrows():
                popup_message += f"- พนักงาน {row['รหัสพนักงาน (EMP ID)']} ได้ {row['จำนวนรวมแท้จริง (ตัว)']} ตัว\n"
            st.toast(popup_message, icon="📊")

            st.success(
                f"🎉 รวมข้อมูลสำเร็จ! ขณะนี้กำลังแสดงผลในโหมด: **{mode_title}**"
            )

            # ลิสต์สรุปพนักงานแยกบรรทัดให้อ่านง่าย
            for _, row in emp_total_df.iterrows():
                st.markdown(
                    f"👤 รหัสพนักงาน: **{row['รหัสพนักงาน (EMP ID)']}** ➡️ ยอดรวมทั้งหมดในโหมดนี้ **{row['จำนวนรวมแท้จริง (ตัว)']:,}** ตัว"
                )

            # 3. แสดงกราฟแท่งเปรียบเทียบยอดรวมรายคนในรอบนี้ (ปรับเปลี่ยนตามโหมดอัตโนมัติ)
            fig = px.bar(
                emp_total_df,
                x="รหัสพนักงาน (EMP ID)",
                y="จำนวนรวมแท้จริง (ตัว)",
                color="รหัสพนักงาน (EMP ID)",
                title=f"กราฟเปรียบเทียบยอดรวมสินค้าของพนักงาน ({mode_title})",
            )
            st.plotly_chart(fig, use_container_width=True)

            # ตารางรายละเอียดสินค้าแยกตามสถานีจริงแบบไม่ซ้ำ
            with st.expander("🔍 คลิกเพื่อดูตารางรายละเอียดแยกตามรหัสพนักงานและสถานี"):
                detail_df = (
                    display_df.groupby(["EMP ID", "STATION"])["SN"]
                    .count()
                    .reset_index()
                )
                detail_df.columns = [
                    "รหัสพนักงาน",
                    "สินค้า/สถานี",
                    "จำนวนจริง (ตัว)",
                ]
                st.dataframe(detail_df, use_container_width=True, hide_index=True)

            # ปุ่มกดบันทึกข้อมูลเข้าคลังประวัติสะสม
            if st.button("📥 บันทึกข้อมูลชุดนี้ลงประวัติยอดรวมสะสม"):
                current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                files_string = ", ".join(file_names_list)

                for _, row in emp_total_df.iterrows():
                    st.session_state.history_log.append(
                        {
                            "เวลาที่บันทึกระบบ": current_time,
                            "จากไฟล์ทั้งหมด": files_string,
                            "โหมดที่เลือกบันทึก": mode_title,
                            "รหัสพนักงาน (EMP ID)": row["รหัสพนักงาน (EMP ID)"],
                            "จำนวนรวมสะสม (ตัว)": row["จำนวนรวมแท้จริง (ตัว)"],
                        }
                    )
                st.toast("บันทึกข้อมูลเข้าคลังประวัติเรียบร้อยแล้ว!")
                st.rerun()
        else:
            st.warning(f"⚠️ ไม่พบข้อมูลที่จะแสดงในโหมด {mode_title}")
    else:
        st.error(
            "❌ ไม่พบโครงสร้างข้อมูลที่ถูกต้อง กรุณาตรวจสอบหัวตารางไฟล์ Excel"
        )

# --------------------------------------------------
# ส่วนที่ 2: หน้าต่างประวัติยอดรวมสะสมย้อนหลัง
# --------------------------------------------------
st.markdown("---")
st.subheader("📜 คลังประวัติยอดรวมสะสมย้อนหลัง (บันทึกตามโหมด)")

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
