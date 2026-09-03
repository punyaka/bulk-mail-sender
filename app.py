"""
Streamlit Bulk-Mailer (Resend API)
────────────────────────────────────────────
• Excel mail-merge — tags {Name}, {Salutation}, … in subject & body
• Quill editors for header / body / footer
• Optional header / footer images
• Per-row PDF (URL or local path)
• Sends via Resend (resend.com) using an API key entered in the sidebar
"""

import base64, pathlib, re
from io import BytesIO

import pandas as pd
import requests
import streamlit as st
from streamlit_quill import st_quill
from jinja2 import Environment, Undefined, select_autoescape

import time
import gdown


# ─── Resend.com email helpers ──────────────────────────────────────────────────
RESEND_API_URL = "https://api.resend.com/emails"

def send_via_resend(api_key, sender, to, subject, html_body, attachment=None):
    """Send one email via the Resend API. attachment = (filename, bytes) or None."""
    payload = {"from": sender, "to": [to], "subject": subject, "html": html_body}
    if attachment:
        fname, data = attachment
        payload["attachments"] = [{
            "filename": fname,
            "content": base64.b64encode(data).decode(),
        }]
    resp = requests.post(
        RESEND_API_URL,
        headers={"Authorization": f"Bearer {api_key}"},
        json=payload,
        timeout=30,
    )
    if not resp.ok:
        try:
            detail = resp.json().get("message", resp.text)
        except ValueError:
            detail = resp.text
        raise RuntimeError(f"{resp.status_code}: {detail}")
    return resp.json()


# ─── Jinja env (single-brace tags) ─────────────────────────────────────────────
class SilentUndef(Undefined):
    def _fail_with_undefined_error(self, *a, **kw):
        return ""
jinja_env = Environment(
    variable_start_string="{",
    variable_end_string="}",
    undefined=SilentUndef,
    autoescape=select_autoescape(enabled_extensions=("html",)),
)

def clean_quill(html: str) -> str:
    html = html.replace("&#123;", "{").replace("&#125;", "}")
    html = re.sub(r"<\w+[^>]*>({\w+})<\/\w+>", r"\1", html)
    html = re.sub(r"{ *([^} ]*?) *}", r"{\1}", html)
    html = re.sub(r"<p><br></p>", "", html)
    return html

def fix_inline_img_widths(html: str, width: int) -> str:
    """Force every <img ...> in Quill HTML to the chosen width."""
    return re.sub(
        r'<img([^>]*?)>',
        lambda m: (
            f'<img{m.group(1)} style="width:{width}px;max-width:100%;'
            'display:block;margin:0 auto;" />'
        ),
        html,
        flags=re.IGNORECASE,
    )

def inline_p_spacing(html: str,
                     margin: str = "0 0 0 0",
                     lh: str = "1.4") -> str:
    """Add spacing *and justify alignment* to every <p …> tag."""
    pattern = re.compile(r"<p\b([^>]*)>", flags=re.IGNORECASE)

    def repl(match):
        attrs = match.group(1)
        if "margin" in attrs or "line-height" in attrs:
            return f"<p{attrs}>"
        style = (
            f'style="margin:{margin};line-height:{lh};'
            'text-align:justify;"'
        )
        return f"<p{attrs} {style}>"

    return pattern.sub(repl, html)

def to_img_tag(file, width, br_after=False):
    if not file:
        return ""
    b64 = base64.b64encode(file.read()).decode()
    tag = (
        f'<img src="data:image/png;base64,{b64}" '
        f'style="width:{width}px;max-width:100%;display:block;margin:0 auto;" />'
    )
    return tag + ("<br>" if br_after else "")


# ─── Fun re-branding! ─────────────────────────────────────────────
APP_NAME = "Mail-MagiK ✨📧"

# Page config + fun intro
st.set_page_config(page_title=APP_NAME, layout="centered",
                   initial_sidebar_state="expanded")

st.title(APP_NAME)

st.markdown(
    """
    **Welcome, daring communicator!**

    • 🪄 Just sprinkle your Excel list into the sidebar cauldron.\n
    • 🪄 Conjure a charming subject and body with `{Tags}` for names, links, emoji—whatever.\n
    • 🪄 Pick a width, images, Google-Drive links…\n
    • 🪄 Press **Preview** to gaze into the crystal ball, then **Send bulk emails** to unleash the owls—I mean, SMTP-owls—one by one (1-second pause so Gmail stays happy).\n

    _May your inbox be ever spellbound!_
    """
)

SAMPLE_DF = pd.DataFrame([
    {"Name": "Jane Doe", "Salutation": "Ms.", "Email": "jane@example.com", "PDF Link": ""},
    {"Name": "John Smith", "Salutation": "Mr.", "Email": "john@example.com", "PDF Link": ""},
])

def sample_excel_bytes() -> bytes:
    buf = BytesIO()
    SAMPLE_DF.to_excel(buf, index=False)
    return buf.getvalue()

with st.sidebar:
    st.download_button(
        "⬇️ Download sample Excel",
        data=sample_excel_bytes(),
        file_name="sample_bulkmailer.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    xlsx = st.file_uploader("Excel file", ["xlsx"])
    resend_api_key = st.text_input("Resend API key", type="password")
    sender = st.text_input("Sender address (must be a verified Resend domain)")
    subj_tpl = st.text_input("Email Subject (tags OK)", "Hello {Name}")
    hdr_img = st.file_uploader("Header image", ["png","jpg","jpeg"])
    ftr_img = st.file_uploader("Footer image", ["png","jpg","jpeg"], key="ftr_img")
    # NEW — choose a common width
    img_width = st.number_input(
        "Width for images & text (px)",
        min_value=200, max_value=1200, value=600, step=20
    )
    st.info("Sending via Resend — get your API key at resend.com/api-keys")

# ─── Excel / tags ─────────────────────────────────────────────────────────────
df, TAGS = None, []
if xlsx:
    df = pd.read_excel(xlsx)
    df.columns = df.columns.str.strip()
    TAGS = df.columns.tolist()
    st.sidebar.markdown("Tags: " + " ".join(f"`{{{t}}}`" for t in TAGS))

compose_tab, status_tab = st.tabs(["✉️ Compose", "📊 Status"])

# ─── Editors ──────────────────────────────────────────────────────────────────
with compose_tab:
    st.markdown("### Header"); header_html = st_quill(html=True, key="hdr")
    st.markdown("### Body");   body_html   = st_quill(html=True, key="bdy")
    st.markdown("### Footer"); footer_html = st_quill(html=True, key="ftr")

    preview_btn, send_btn = st.columns(2)
    preview_click = preview_btn.button("Preview first email")
    send_click    = send_btn.button("Send bulk emails")

    # clean editor output
    header_html = inline_p_spacing(
        fix_inline_img_widths(clean_quill(header_html), img_width)
    )
    body_html   = inline_p_spacing(
        fix_inline_img_widths(clean_quill(body_html), img_width)
    )
    footer_html = inline_p_spacing(
        fix_inline_img_widths(clean_quill(footer_html), img_width)
    )

    hdr_tag = to_img_tag(hdr_img, img_width, br_after=True)
    ftr_tag = "<br>" + to_img_tag(ftr_img, img_width) if ftr_img else ""

    # after you compute img_width
    style_tag = (
        "{% raw %}"
        "<style>"
        f".mail-preview p {{margin:0 0 0em 0;line-height:1.4;}}"
        "</style>"
        "{% endraw %}"
    )

    wrapper_start = (
        f'<div class="mail-preview" '
        f'style="max-width:{img_width}px;margin:0 auto;">'
    )

    body_template = (
        style_tag +
        wrapper_start +
        hdr_tag +
        header_html + body_html + footer_html +
        ftr_tag +
        '</div>'
    )

    subj_template = jinja_env.from_string(subj_tpl)

    # ─── Preview ────────────────────────────────────────────────────────────
    if preview_click and df is not None:
        sample = df.iloc[0].to_dict()
        st.markdown(f"**Subject:** {subj_template.render(**sample)}")
        st.markdown(jinja_env.from_string(body_template).render(**sample),
                    unsafe_allow_html=True)

    # ─── Bulk send ──────────────────────────────────────────────────────────
    if send_click:
        if df is None or not sender or not resend_api_key:
            st.error("Please provide Excel file, sender address, **and** Resend API key.")
        else:
            logs = []
            progress = st.progress(0.0)
            status_line = st.empty()
            total = len(df)
            for i, (_, row) in enumerate(df.iterrows()):
                data = row.to_dict()
                email = data.get("Email", "")
                status_line.info(f"Sending {i + 1}/{total}: {email}")
                html = jinja_env.from_string(body_template).render(**data)
                subj = subj_template.render(**data)

                # optional one attachment
                attach = None
                pdf_raw = data.get("PDF Link", "")
                pdf = "" if pd.isna(pdf_raw) else str(pdf_raw).strip()
                if pdf:
                    try:
                        if pdf.lower().startswith("http"):
                            tmp_path = gdown.download(pdf, quiet=True)
                            fname = pathlib.Path(tmp_path).name
                            if not fname.lower().endswith(".pdf"):
                                fname += ".pdf"                       # force .pdf so Gmail knows it
                            filebytes = pathlib.Path(tmp_path).read_bytes()
                            attach = (fname, filebytes)
                        else:
                            filebytes = pathlib.Path(pdf).read_bytes()
                            attach = (pathlib.Path(pdf).name, filebytes)
                    except Exception as e:
                        logs.append({"email": email, "status": "Failed",
                                     "error": f"PDF err: {e}"})
                        progress.progress((i + 1) / total)
                        continue

                try:
                    send_via_resend(resend_api_key, sender, email,
                                     subj, html, attach)
                    logs.append({"email": email, "status": "Sent", "error": ""})
                except Exception as e:
                    logs.append({"email": email, "status": "Failed", "error": str(e)})

                progress.progress((i + 1) / total)
                time.sleep(10)  # pause before next email

            status_line.empty()
            st.session_state["logs"] = logs
            st.success("Done — see the Status tab for a full breakdown.")

# ─── Status ─────────────────────────────────────────────────────────────────
with status_tab:
    logs = st.session_state.get("logs")
    if not logs:
        st.info("No emails sent yet this session.")
    else:
        log_df = pd.DataFrame(logs)
        sent_n = (log_df["status"] == "Sent").sum()
        failed_n = (log_df["status"] == "Failed").sum()

        c1, c2, c3 = st.columns(3)
        c1.metric("Total", len(log_df))
        c2.metric("Sent", sent_n)
        c3.metric("Failed", failed_n)

        if failed_n:
            st.markdown("#### ❌ Failed")
            st.dataframe(log_df[log_df["status"] == "Failed"], use_container_width=True)

        st.markdown("#### Full log")
        st.dataframe(log_df, use_container_width=True)
