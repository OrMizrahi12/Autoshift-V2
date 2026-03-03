import streamlit as st
import fitz  # PyMuPDF
from PIL import Image
import io
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
import time
import os

def load_css():
    css_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "style.css")
    if os.path.exists(css_path):
        with open(css_path, "r", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

def find_and_render_page(uploaded_file, search_term):
    """מחפש דפים ב-PDF לפי מונח חיפוש ומחזיר תמונות."""
    pdf_bytes = uploaded_file.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    matches = []
    search_term_reversed = search_term[::-1]

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()

        if search_term in text or search_term_reversed in text:
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_data = pix.tobytes("png")
            image = Image.open(io.BytesIO(img_data))

            matches.append({
                "page": page_num + 1,
                "image": image,
                "image_bytes": img_data
            })

    return matches

def clean_search_value(value):
    """מנקה ערך חיפוש - מסיר .0 ממספרים, מסיר רווחים מיותרים."""
    text = str(value).strip()
    try:
        num = float(text)
        if num == int(num):
            text = str(int(num))
    except (ValueError, OverflowError):
        pass
    return text

def find_pages_by_term(pdf_bytes, search_term):
    """מחפש דפים ב-PDF לפי מונח חיפוש (מקבל bytes ישירות)."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    matches = []
    search_term_clean = clean_search_value(search_term)
    search_term_reversed = search_term_clean[::-1]
    words = search_term_clean.split()

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()

        full_match = search_term_clean in text or search_term_reversed in text
        words_match = len(words) > 1 and all(
            (w in text or w[::-1] in text) for w in words
        )

        if full_match or words_match:
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_data = pix.tobytes("png")
            image = Image.open(io.BytesIO(img_data))

            matches.append({
                "page": page_num + 1,
                "image": image,
                "image_bytes": img_data
            })

    return matches

def send_email_with_report(sender_email, sender_password, recipient_email,
                           smtp_server, smtp_port, images_data, employee_name):
    """שולח את דוח השעות כתמונות מצורפות למייל."""
    msg = MIMEMultipart("related")
    msg["From"] = sender_email
    msg["To"] = recipient_email
    msg["Subject"] = f"דוח שעות - {employee_name}"

    html_parts = [
        "<html><body dir='rtl' style='font-family: Arial, sans-serif;'>",
        f"<h2>שלום {employee_name},</h2>",
        f"<p>מצורף דוח השעות שלך.</p>",
        f"<p>נמצאו {len(images_data)} דפים רלוונטיים:</p>",
        "<hr>"
    ]

    for i, img_data in enumerate(images_data):
        cid = f"report_page_{i}"
        html_parts.append(f"<h3>📄 דף מספר {img_data['page']}</h3>")
        html_parts.append(f'<img src="cid:{cid}" style="max-width:100%; border:1px solid #ccc; margin-bottom:20px;">')
        html_parts.append("<hr>")

    html_parts.append("<p>בברכה,<br>מערכת דוחות שעות</p>")
    html_parts.append("</body></html>")

    html_body = MIMEText("\n".join(html_parts), "html", "utf-8")
    msg.attach(html_body)

    for i, img_data in enumerate(images_data):
        cid = f"report_page_{i}"
        img_attachment = MIMEImage(img_data["image_bytes"], name=f"page_{img_data['page']}.png")
        img_attachment.add_header("Content-ID", f"<{cid}>")
        img_attachment.add_header("Content-Disposition", "inline", filename=f"page_{img_data['page']}.png")
        msg.attach(img_attachment)

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())

def auto_detect_columns(df):
    """מנסה לזהות אוטומטית את העמודות הרלוונטיות בטבלה."""
    name_keywords = ["שם", "name", "שם מלא", "שם עובד", "שם פרטי", "full_name", "employee"]
    id_keywords = ["ת.ז", "ת\"ז", "תז", "תעודת זהות", "tz", "id", "מספר זהות", "id_number", "identity"]
    email_keywords = ["מייל", "אימייל", "דואל", "דוא\"ל", "email", "mail", "e-mail", "דואר"]

    detected = {"name": None, "id": None, "email": None}

    for col in df.columns:
        col_lower = str(col).strip().lower()
        for kw in name_keywords:
            if kw in col_lower:
                detected["name"] = col
                break
        for kw in id_keywords:
            if kw in col_lower:
                detected["id"] = col
                break
        for kw in email_keywords:
            if kw in col_lower:
                detected["email"] = col
                break

    return detected

def render_smtp_sidebar():
    """מציג את הגדרות SMTP בסיידבר ומחזיר את הערכים."""
    with st.sidebar:
        st.header("⚙️ הגדרות מייל")
        st.caption("הגדר את פרטי חשבון המייל השולח")

        smtp_server = st.text_input("שרת SMTP", value="smtp.gmail.com")
        smtp_port = st.number_input("פורט SMTP", value=587)
        sender_email = st.text_input("כתובת מייל שולח", placeholder="example@gmail.com")
        sender_password = st.text_input("סיסמת אפליקציה", type="password",
                                        help="עבור Gmail: צור App Password בהגדרות חשבון Google")

        st.markdown("---")
        st.markdown("""
        **💡 הוראות ל-Gmail:**
        1. הפעל [אימות דו-שלבי](https://myaccount.google.com/security)
        2. צור [App Password](https://myaccount.google.com/apppasswords)
        3. הדבק את הסיסמה שנוצרה למעלה
        """)

    return smtp_server, smtp_port, sender_email, sender_password

def tab_single_search(uploaded_pdf, smtp_server, smtp_port, sender_email, sender_password):
    """טאב חיפוש בודד - שומר על הפונקציונליות הקיימת."""
    st.markdown("### 🔎 חיפוש עובד בודד")
    st.caption("חפש לפי שם או ת.ז -> קבל תמונה -> שלח במייל או קליק ימני 'העתק' -> הדבק בוואטסאפ.")

    search_term = st.text_input("הקלד שם פרטי, שם מלא או תעודת זהות:", key="single_search")

    if uploaded_pdf and search_term:
        if st.button("חפש והצג 🔎", type="primary", key="btn_single_search"):
            with st.spinner("סורק את המסמך..."):
                uploaded_pdf.seek(0)
                results = find_and_render_page(uploaded_pdf, search_term)

            if results:
                st.session_state["single_results"] = results
                st.session_state["single_search_term"] = search_term
            else:
                st.session_state.pop("single_results", None)
                st.error(f"לא נמצא שום דף המכיל את הטקסט: '{search_term}'")

    if "single_results" in st.session_state:
        results = st.session_state["single_results"]
        current_search = st.session_state.get("single_search_term", "")

        st.success(f"נמצאו {len(results)} דפים מתאימים!")

        for res in results:
            st.markdown(f"### 📄 דף מספר {res['page']}")
            st.image(res['image'], caption="קליק ימני על התמונה -> העתק תמונה", use_container_width=True)
            st.markdown("---")

        st.markdown("## 📧 שליחה במייל")
        recipient_email = st.text_input("כתובת מייל של העובד:", placeholder="worker@example.com",
                                        key="single_recipient")

        if st.button("📤 שלח במייל", type="primary", key="btn_single_send"):
            if not recipient_email:
                st.error("❌ נא להזין כתובת מייל של העובד")
            elif not sender_email or not sender_password:
                st.error("❌ נא להגדיר את פרטי המייל השולח בסיידבר")
            elif "@" not in recipient_email:
                st.error("❌ כתובת המייל לא תקינה")
            else:
                with st.spinner("📤 שולח את הדוח במייל..."):
                    try:
                        send_email_with_report(sender_email, sender_password, recipient_email,
                                               smtp_server, smtp_port, results, current_search)
                        st.success(f"✅ הדוח נשלח בהצלחה ל-{recipient_email}!")
                        st.balloons()
                    except smtplib.SMTPAuthenticationError:
                        st.error("❌ שגיאת אימות: בדוק את כתובת המייל והסיסמה בסיידבר.")
                    except Exception as e:
                        st.error(f"❌ שגיאה: {str(e)}")

def tab_bulk_send(uploaded_pdf, smtp_server, smtp_port, sender_email, sender_password):
    """טאב שיגור מרוכז - שולח דוחות לכל העובדים מתוך קובץ Excel."""
    st.markdown("### 📨 שיגור מרוכז לכל העובדים")
    st.caption("העלה קובץ Excel עם פרטי עובדים (שם, ת.ז, מייל) ושלח לכולם בלחיצה אחת.")

    if not uploaded_pdf:
        st.warning("⚠️ נא להעלות קובץ PDF עם דוחות השעות למעלה.")
        return

    excel_file = st.file_uploader("📊 העלה קובץ Excel עם פרטי עובדים", type=["xlsx", "xls", "csv"],
                                   key="excel_upload")

    if not excel_file:
        return

    try:
        if excel_file.name.endswith(".csv"):
            df = pd.read_csv(excel_file)
        else:
            df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"❌ שגיאה בקריאת הקובץ: {str(e)}")
        return

    st.success(f"✅ הקובץ נקרא בהצלחה! {len(df)} שורות נמצאו.")

    detected = auto_detect_columns(df)

    st.markdown("#### 🔧 בחירת עמודות")
    st.caption("המערכת ניסתה לזהות את העמודות אוטומטית. ניתן לשנות במידת הצורך.")

    col_options = ["-- לא נבחר --"] + list(df.columns)

    col1, col2, col3 = st.columns(3)
    with col1:
        name_col = st.selectbox("עמודת שם 👤",
                                options=col_options,
                                index=col_options.index(detected["name"]) if detected["name"] in col_options else 0,
                                key="sel_name")
    with col2:
        id_col = st.selectbox("עמודת ת.ז 🪪",
                              options=col_options,
                              index=col_options.index(detected["id"]) if detected["id"] in col_options else 0,
                              key="sel_id")
    with col3:
        email_col = st.selectbox("עמודת דוא\"ל 📧",
                                 options=col_options,
                                 index=col_options.index(detected["email"]) if detected["email"] in col_options else 0,
                                 key="sel_email")

    if name_col == "-- לא נבחר --" or email_col == "-- לא נבחר --":
        st.warning("⚠️ נא לבחור לפחות את עמודת השם ועמודת הדוא\"ל.")
        return

    use_id_for_search = id_col != "-- לא נבחר --"

    st.markdown("#### 🔍 בחירת שיטת חיפוש בדוח")
    search_method = st.radio(
        "לפי מה לחפש את העובד בקובץ ה-PDF?",
        options=["חיפוש לפי שם", "חיפוש לפי ת.ז"] if use_id_for_search else ["חיפוש לפי שם"],
        horizontal=True,
        key="search_method"
    )

    search_col = name_col if search_method == "חיפוש לפי שם" else id_col

    if st.button("🔍 הכן לשיגור - סרוק את ה-PDF", type="primary", key="btn_prepare"):
        uploaded_pdf.seek(0)
        pdf_bytes = uploaded_pdf.read()

        preparation_results = []
        progress_bar = st.progress(0, text="סורק את ה-PDF עבור כל עובד...")

        total = len(df)
        for idx, row in df.iterrows():
            employee_name = clean_search_value(row[name_col])
            employee_email = str(row[email_col]).strip()
            search_value = clean_search_value(row[search_col])

            if not employee_name or employee_name == "nan" or not employee_email or employee_email == "nan":
                continue

            pages_found = find_pages_by_term(pdf_bytes, search_value)

            preparation_results.append({
                "name": employee_name,
                "email": employee_email,
                "search_term": search_value,
                "pages_found": len(pages_found),
                "pages_data": pages_found,
                "page_numbers": ", ".join([str(p["page"]) for p in pages_found]) if pages_found else "—"
            })

            progress_bar.progress((idx + 1) / total, text=f"סורק: {employee_name} ({idx + 1}/{total})")

        progress_bar.empty()
        st.session_state["bulk_results"] = preparation_results

    if "bulk_results" in st.session_state:
        results = st.session_state["bulk_results"]

        found_count = sum(1 for r in results if r["pages_found"] > 0)
        not_found_count = sum(1 for r in results if r["pages_found"] == 0)

        st.markdown("---")
        st.markdown("## 📋 סיכום לפני שיגור")

        stat1, stat2, stat3 = st.columns(3)
        with stat1:
            st.metric("סה\"כ עובדים", len(results))
        with stat2:
            st.metric("נמצאו דפים ✅", found_count)
        with stat3:
            st.metric("לא נמצאו ❌", not_found_count)

        preview_data = []
        for r in results:
            status = "✅ נמצא" if r["pages_found"] > 0 else "❌ לא נמצא"
            preview_data.append({
                "שם": r["name"],
                "דוא\"ל": r["email"],
                "מונח חיפוש": r["search_term"],
                "דפים שנמצאו": r["pages_found"],
                "מספרי דפים": r["page_numbers"],
                "סטטוס": status
            })

        preview_df = pd.DataFrame(preview_data)
        st.dataframe(preview_df, use_container_width=True, hide_index=True)

        if not_found_count > 0:
            st.warning(f"⚠️ שים לב: ל-{not_found_count} עובדים לא נמצאו דפים בדוח. הם לא יקבלו מייל.")

        st.markdown("---")

        if not sender_email or not sender_password:
            st.error("❌ נא להגדיר את פרטי המייל השולח בסיידבר (⚙️ הגדרות מייל) לפני השיגור.")
            return

        st.markdown(f"**📤 ישלח מ:** `{sender_email}`")
        st.markdown(f"**👥 ישלח ל:** {found_count} עובדים")

        col_confirm, col_cancel = st.columns([1, 3])
        with col_confirm:
            confirm_send = st.button("✅ אשר ושגר לכולם!", type="primary", key="btn_confirm_send")
        with col_cancel:
            if st.button("🗑️ בטל", key="btn_cancel"):
                st.session_state.pop("bulk_results", None)
                st.rerun()

        if confirm_send:
            employees_to_send = [r for r in results if r["pages_found"] > 0]

            if not employees_to_send:
                st.error("❌ אין עובדים עם דפים שנמצאו לשלוח להם.")
                return

            progress_bar = st.progress(0, text="מתחיל שיגור...")
            success_count = 0
            fail_count = 0
            errors = []

            for idx, emp in enumerate(employees_to_send):
                progress_bar.progress(
                    (idx + 1) / len(employees_to_send),
                    text=f"שולח ל-{emp['name']} ({idx + 1}/{len(employees_to_send)})..."
                )

                try:
                    send_email_with_report(
                        sender_email=sender_email,
                        sender_password=sender_password,
                        recipient_email=emp["email"],
                        smtp_server=smtp_server,
                        smtp_port=smtp_port,
                        images_data=emp["pages_data"],
                        employee_name=emp["name"]
                    )
                    success_count += 1
                    time.sleep(1)
                except Exception as e:
                    fail_count += 1
                    errors.append(f"{emp['name']} ({emp['email']}): {str(e)}")

            progress_bar.empty()

            st.markdown("---")
            st.markdown("## 📊 סיכום שיגור")

            res1, res2 = st.columns(2)
            with res1:
                st.metric("נשלחו בהצלחה ✅", success_count)
            with res2:
                st.metric("נכשלו ❌", fail_count)

            if success_count > 0:
                st.success(f"🎉 {success_count} דוחות נשלחו בהצלחה!")
                st.balloons()

            if errors:
                st.error("שגיאות בשיגור:")
                for err in errors:
                    st.markdown(f"- {err}")

def main():
    st.set_page_config(page_title="מערכת דוחות שעות", page_icon="📸", layout="centered")
    load_css()

    st.title("📸 מערכת דוחות שעות")

    smtp_server, smtp_port, sender_email, sender_password = render_smtp_sidebar()

    uploaded_pdf = st.file_uploader("📄 העלה קובץ PDF עם דוחות שעות", type="pdf", key="pdf_upload")

    tab1, tab2 = st.tabs(["🔎 חיפוש בודד", "📨 שיגור מרוכז"])

    with tab1:
        tab_single_search(uploaded_pdf, smtp_server, smtp_port, sender_email, sender_password)

    with tab2:
        tab_bulk_send(uploaded_pdf, smtp_server, smtp_port, sender_email, sender_password)

if __name__ == "__main__":
    main()
