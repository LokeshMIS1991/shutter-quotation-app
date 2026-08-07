import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import io
import json
import os

st.set_page_config(page_title="Sidharth Shutters - Advanced Quotation Builder", layout="wide")

st.title("🛡️ Sidharth Shutters & Automation Pvt. Ltd.")
st.subheader("Automated Quotation Builder (With Dual Pricing & Custom Specs)")

# --- MASTER DATA FILE MANAGEMENT ---
MASTER_FILE = "master_specs.json"

DEFAULT_MASTERS = {
    "shutter_types": ["Motorized Rolling Shutter", "Manual Rolling Shutter", "Fire Rated Rolling Shutter"],
    "slat_options": [
        "90mm (H) x 1.2 mm thick Galvalume Plain slats in natural finish",
        "90mm (H) x 1.0 mm thick Galvalume Curved slat in natural finish",
        "Perforation- 7 Strip(Square)",
        "Perforation- 4 Strip(Square)",
        "78 mm (H) x 0.8 mm thick GI Slat in natural finish"
    ],
    "guide_options": [
        "TG Guide with rubber seal with grey epoxy",
        "150x150x2mmTG Guide with rubber seal with grey epoxy",
        "2mmTG Guide with rubber seal with grey epoxy"
    ],
    "bottom_hood_options": [
        "Super bottom with rubber seal with grey epoxy",
        ".80mm thick Galvalume Hood & Motor cover in natural finish",
        ".80mm thick Galvalume Hood cover in natural finish"
    ],
    "safety_lock_options": [
        "Wind Locks | Storm Anchors | External Safety Break",
        "Side Locks-2 Nos",
        "Standard Center Lock"
    ],
    "actuator_options": [
        "CE Certified indirect Drive - Brand Strong Life",
        "Direct Drive Heavy Duty Actuator",
        "Manual Gear / Chain Pulley Mechanism"
    ]
}

def load_masters():
    if os.path.exists(MASTER_FILE):
        try:
            with open(MASTER_FILE, "r") as f:
                return json.load(f)
        except:
            return DEFAULT_MASTERS
    return DEFAULT_MASTERS

def save_masters(data):
    with open(MASTER_FILE, "w") as f:
        json.dump(data, f, indent=4)

master_data = load_masters()

# --- CLIENT & HEADER DETAILS ---
st.sidebar.header("📋 Client & Proposal Details")
client_name = st.sidebar.text_input("Company Name", "M/S ABHINAV INFRA BUILD PVT. LTD.")
contact_details = st.sidebar.text_input("Contact Details", "Mob:-8889911529")
shipping_address = st.sidebar.text_area("Shipping Address", "RGLP PITHAMPUR Sector-1 Pithampur MADHYA PRADESH 454774")
site_person = st.sidebar.text_input("Site Person Mobile", "9926952398")
billing_address = st.sidebar.text_area("Billing Address", "207,208, Industry House, Indore, MP 452001")
gstin = st.sidebar.text_input("Client GSTIN", "23AAHCA9425D1ZY")

c_date, c_ref = st.sidebar.columns(2)
qtn_date = c_date.date_input("Quotation Date")
qtn_ref_no = c_ref.text_input("Qtn Ref No", "SSAPL/2026-27/319R2")

st.sidebar.markdown("---")
sales_person_1 = st.sidebar.text_input("Sales Person 1", "Mr. Nishant (90010 42914)")
sales_person_2 = st.sidebar.text_input("Sales Person 2", "Mr. Jeevan Sharma (9828771899)")

# --- SHUTTER ITEMS INGREDIENTS ---
st.header("🧱 Shutter Specifications & Pricing")

if "shutter_items" not in st.session_state:
    st.session_state["shutter_items"] = [{
        "type": master_data["shutter_types"][0],
        "slat": master_data["slat_options"][0],
        "guide": master_data["guide_options"][0],
        "bottom_hood": master_data["bottom_hood_options"][0],
        "safety_lock": master_data["safety_lock_options"][0],
        "actuator": master_data["actuator_options"][0],
        "custom_note": "",
        "hsn": "73083000",
        "width": 4650,
        "height": 7000,
        "qty": 5,
        "mat_rate": 176700,
        "inst_rate": 8000
    }]

def add_shutter():
    st.session_state["shutter_items"].append({
        "type": master_data["shutter_types"][0],
        "slat": master_data["slat_options"][0],
        "guide": master_data["guide_options"][0],
        "bottom_hood": master_data["bottom_hood_options"][0],
        "safety_lock": master_data["safety_lock_options"][0],
        "actuator": master_data["actuator_options"][0],
        "custom_note": "",
        "hsn": "73083000",
        "width": 4000,
        "height": 5000,
        "qty": 1,
        "mat_rate": 99578,
        "inst_rate": 6000
    })

for idx, item in enumerate(st.session_state["shutter_items"]):
    with st.expander(f"Shutter #{idx + 1} ({item['type']})", expanded=True):
        c1, c2, c3 = st.columns(3)
        item["type"] = c1.selectbox("Shutter Category", master_data["shutter_types"], index=0, key=f"type_{idx}")
        item["hsn"] = c2.text_input("HSN Code", value=item["hsn"], key=f"hsn_{idx}")
        item["slat"] = c3.selectbox("Slat Type", master_data["slat_options"], index=0, key=f"slat_{idx}")

        c4, c5, c6 = st.columns(3)
        item["guide"] = c4.selectbox("Guide Specification", master_data["guide_options"], index=0, key=f"guide_{idx}")
        item["bottom_hood"] = c5.selectbox("Bottom Sheet / Hood Cover", master_data["bottom_hood_options"], index=0, key=f"hood_{idx}")
        item["safety_lock"] = c6.selectbox("Locks & Safety Features", master_data["safety_lock_options"], index=0, key=f"lock_{idx}")

        c7, c8 = st.columns(2)
        item["actuator"] = c7.selectbox("Actuator / Drive", master_data["actuator_options"], index=0, key=f"act_{idx}")
        item["custom_note"] = c8.text_input("➕ Custom Note / Free-Text Spec (Optional)", value=item["custom_note"], key=f"custom_{idx}")

        st.markdown("**Sizes & Dual Rates (Material vs Installation):**")
        cw, ch, cq, cmr, cir = st.columns(5)
        item["width"] = cw.number_input("Width (mm)", value=int(item["width"]), step=50, key=f"w_{idx}")
        item["height"] = ch.number_input("Height (mm)", value=int(item["height"]), step=50, key=f"h_{idx}")
        item["qty"] = cq.number_input("Qty", value=int(item["qty"]), min_value=1, step=1, key=f"q_{idx}")
        item["mat_rate"] = cmr.number_input("Material Unit Rate", value=int(item["mat_rate"]), step=500, key=f"mr_{idx}")
        item["inst_rate"] = cir.number_input("Installation Unit Rate", value=int(item["inst_rate"]), step=100, key=f"ir_{idx}")

st.button("➕ Add Another Shutter Item", on_click=add_shutter)

# --- CHARGES & TAXES ---
st.header("🚚 Extra Charges")
col_p, col_f = st.columns(2)
packing_charges = col_p.number_input("Packing & Loading Charges (INR)", value=5000)
freight_charges = col_f.number_input("Freight Charges (INR)", value=15000)

# --- PDF GENERATOR (EXACT IMAGE FORMAT MATCHING) ---
def generate_pdf():
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=15, leftMargin=15, topMargin=15, bottomMargin=15)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('Title', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, alignment=1, textColor=colors.HexColor('#1A365D'))
    small_bold = ParagraphStyle('SmallBold', fontName='Helvetica-Bold', fontSize=7.5, leading=9)
    small_text = ParagraphStyle('SmallText', fontName='Helvetica', fontSize=7.5, leading=9)

    story.append(Paragraph("SIDHARTH SHUTTER & AUTOMATION PRIVATE LIMITED", title_style))
    story.append(Paragraph("G-1-66, Industrial Area, Prahaladpura, Sanganer, Jaipur, Rajasthan, 303903", ParagraphStyle('Sub', alignment=1, fontSize=7.5)))
    story.append(Spacer(1, 8))

    header_data = [
        [Paragraph("<b>Company Name</b>", small_bold), Paragraph(client_name, small_text), Paragraph("<b>Qtn Date:</b>", small_bold), Paragraph(str(qtn_date), small_text)],
        [Paragraph("<b>Contact Details</b>", small_bold), Paragraph(contact_details, small_text), Paragraph("<b>Sales Staff:</b>", small_bold), Paragraph(f"{sales_person_1}<br/>{sales_person_2}", small_text)],
        [Paragraph("<b>Shipping Address</b>", small_bold), Paragraph(shipping_address, small_text), Paragraph("<b>Qtn Ref No:</b>", small_bold), Paragraph(qtn_ref_no, small_bold)],
        [Paragraph("<b>Billing Address</b>", small_bold), Paragraph(billing_address, small_text), Paragraph("", small_text), Paragraph("", small_text)],
        [Paragraph("<b>GSTIN</b>", small_bold), Paragraph(gstin, small_text), Paragraph("", small_text), Paragraph("", small_text)]
    ]
    t_header = Table(header_data, colWidths=[80, 250, 65, 170])
    t_header.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(t_header)
    story.append(Spacer(1, 8))

    # Dual Rate Table Header
    items_data = [[
        Paragraph("<b>Sr. No.</b>", small_bold),
        Paragraph("<b>Description</b>", small_bold),
        Paragraph("<b>HSN Code</b>", small_bold),
        Paragraph("<b>Width (mm)</b>", small_bold),
        Paragraph("<b>Height (mm)</b>", small_bold),
        Paragraph("<b>Qty</b>", small_bold),
        Paragraph("<b>Mat. Rate (INR)</b>", small_bold),
        Paragraph("<b>Mat. Amount (INR)</b>", small_bold),
        Paragraph("<b>Inst. Rate (INR)</b>", small_bold),
        Paragraph("<b>Inst. Amount (INR)</b>", small_bold)
    ]]

    mat_grand_total = 0
    inst_grand_total = 0
    total_qty = 0

    for i, itm in enumerate(st.session_state["shutter_items"]):
        m_amt = itm["qty"] * itm["mat_rate"]
        i_amt = itm["qty"] * itm["inst_rate"]
        mat_grand_total += m_amt
        inst_grand_total += i_amt
        total_qty += itm["qty"]

        desc = f"<b>{itm['type']}</b><br/>" \
               f"- {itm['slat']}<br/>" \
               f"- {itm['guide']}<br/>" \
               f"- {itm['bottom_hood']}<br/>" \
               f"- {itm['safety_lock']}<br/>" \
               f"<b>Actuator / Drive:</b><br/>" \
               f"- {itm['actuator']}"
        
        if itm["custom_note"]:
            desc += f"<br/>- {itm['custom_note']}"

        items_data.append([
            Paragraph(str(i+1), small_text),
            Paragraph(desc, small_text),
            Paragraph(itm["hsn"], small_text),
            Paragraph(str(itm["width"]), small_text),
            Paragraph(str(itm["height"]), small_text),
            Paragraph(str(itm["qty"]), small_text),
            Paragraph(f"{itm['mat_rate']:,}", small_text),
            Paragraph(f"{m_amt:,}", small_text),
            Paragraph(f"{itm['inst_rate']:,}", small_text),
            Paragraph(f"{i_amt:,}", small_text)
        ])

    subtotal_mat = mat_grand_total + packing_charges + freight_charges
    gst_mat = round(subtotal_mat * 0.18)
    gst_inst = round(inst_grand_total * 0.18)
    
    total_mat_with_gst = subtotal_mat + gst_mat
    total_inst_with_gst = inst_grand_total + gst_inst
    final_grand_total = total_mat_with_gst + total_inst_with_gst

    items_data.append(["", Paragraph("<b>Packing & Loading charges</b>", small_bold), "", "", "", Paragraph(f"<b>{total_qty}</b>", small_bold), "<b>TOTAL</b>", Paragraph(f"<b>{mat_grand_total:,}</b>", small_bold), "", Paragraph(f"<b>{inst_grand_total:,}</b>", small_bold)])
    items_data.append(["", Paragraph("Packing Charges", small_text), "", "", "", "", "", Paragraph(f"{packing_charges:,}", small_text), "", "-"])
    items_data.append(["", Paragraph("Freight charges", small_text), "", "", "", "", "", Paragraph(f"{freight_charges:,}", small_text), "", "-"])
    items_data.append(["", Paragraph("<b>Supply & Installation Amount Excluding GST</b>", small_bold), "", "", "", "", "", Paragraph(f"<b>{subtotal_mat:,}</b>", small_bold), "", Paragraph(f"<b>{inst_grand_total:,}</b>", small_bold)])
    items_data.append(["", Paragraph("GST on supply & installation @18%", small_text), "", "", "", "", "", Paragraph(f"{gst_mat:,}", small_text), "@18%", Paragraph(f"{gst_inst:,}", small_text)])
    items_data.append(["", Paragraph("<b>Total supply & installation with GST</b>", small_bold), "", "", "", "", "", Paragraph(f"<b>{total_mat_with_gst:,}</b>", small_bold), "", Paragraph(f"<b>{total_inst_with_gst:,}</b>", small_bold)])
    items_data.append(["", Paragraph("<b>Grand total with GST</b>", small_bold), "", "", "", "", "", "", "", Paragraph(f"<b>{final_grand_total:,}</b>", small_bold)])

    t_items = Table(items_data, colWidths=[20, 185, 45, 35, 35, 20, 55, 60, 50, 60])
    t_items.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#E2E8F0')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(t_items)

    doc.build(story)
    buffer.seek(0)
    return buffer

st.markdown("---")
if st.button("📄 Generate PDF Quotation", type="primary"):
    pdf_buffer = generate_pdf()
    st.download_button(
        label="📥 Download Quotation PDF",
        data=pdf_buffer,
        file_name=f"Quotation_{qtn_ref_no.replace('/', '_')}.pdf",
        mime="application/pdf"
    )
