import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import io

st.set_page_config(page_title="Sidharth Shutters - Advanced Quotation Builder", layout="wide")

st.title("🛡️ Sidharth Shutters & Automation Pvt. Ltd.")
st.subheader("Automated Quotation Builder")

# --- MASTER DROPDOWN LISTS ---
if "slat_nat_list" not in st.session_state:
    st.session_state["slat_nat_list"] = ["90mm (H) x 1.2 mm thick Galvalume Plain slats in natural finish"]
if "slat_pow_list" not in st.session_state:
    st.session_state["slat_pow_list"] = ["90mm (H) x 1.2 mm thick Galvalume Powder Coated slats"]
if "guide_list" not in st.session_state:
    st.session_state["guide_list"] = ["TG Guide with rubber seal with grey epoxy"]
if "bottom_list" not in st.session_state:
    st.session_state["bottom_list"] = ["Super bottom with rubber seal with grey epoxy"]
if "hood_list" not in st.session_state:
    st.session_state["hood_list"] = [".80mm thick Galvalume Hood & Motor cover in natural finish"]

SAFETY_LOCK_OPTIONS = [
    "Wind Locks",
    "Storm Anchors",
    "External Safety Break",
    "Side Locks-2 Nos",
    "Standard Center Lock with 2 Keys"
]

OPERATOR_OPTIONS = [
    "CE Certified Indirect Drive Brand Strong Life Sidharth Make",
    "Three Station Push Button"
]

SHUTTER_CATEGORIES = [
    "Motorized Rolling Shutter",
    "Gear Rolling Shutter",
    "Manual Rolling Shutter"
]

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

# --- SHUTTER ITEMS MANAGEMENT ---
st.header("🧱 Shutter Specifications & Pricing")

if "shutter_items" not in st.session_state:
    st.session_state["shutter_items"] = [{
        "type": "Motorized Rolling Shutter",
        "slat_nat": [],
        "slat_pow": [],
        "guide": [],
        "bottom": [],
        "hood": [],
        "safety_locks": ["Wind Locks", "Storm Anchors", "External Safety Break"],
        "operator": ["CE Certified Indirect Drive Brand Strong Life Sidharth Make"],
        "hsn": "73083000",
        "width": 4650,
        "height": 7000,
        "qty": 5,
        "mat_rate": 176700,
        "inst_rate": 8000
    }]

def add_shutter():
    st.session_state["shutter_items"].append({
        "type": "Motorized Rolling Shutter",
        "slat_nat": [],
        "slat_pow": [],
        "guide": [],
        "bottom": [],
        "hood": [],
        "safety_locks": [],
        "operator": [],
        "hsn": "73083000",
        "width": 4000,
        "height": 5000,
        "qty": 1,
        "mat_rate": 99578,
        "inst_rate": 6000
    })

def remove_shutter(index):
    if len(st.session_state["shutter_items"]) > 1:
        st.session_state["shutter_items"].pop(index)

for idx, item in enumerate(st.session_state["shutter_items"]):
    with st.expander(f"Shutter Item #{idx + 1}: {item.get('type', 'Motorized Rolling Shutter')}", expanded=True):
        
        col_title, col_del = st.columns([5, 1])
        with col_title:
            current_type = item.get("type", "Motorized Rolling Shutter")
            item["type"] = st.selectbox(
                f"Shutter Category #{idx + 1}", 
                SHUTTER_CATEGORIES, 
                index=SHUTTER_CATEGORIES.index(current_type) if current_type in SHUTTER_CATEGORIES else 0,
                key=f"type_{idx}"
            )
        with col_del:
            st.markdown("<br>", unsafe_allow_html=True)
            if len(st.session_state["shutter_items"]) > 1:
                st.button("❌ Remove", key=f"del_{idx}", on_click=remove_shutter, args=(idx,))

        item["hsn"] = st.text_input("HSN Code", value=item.get("hsn", "73083000"), key=f"hsn_{idx}")

        # --- SLAT TYPES ---
        st.markdown("##### 📐 Slat Specifications")
        col_sn, col_sp = st.columns(2)
        
        with col_sn:
            item["slat_nat"] = st.multiselect("Slat Type - Natural Finish", st.session_state["slat_nat_list"], default=item.get("slat_nat", []), key=f"sn_{idx}")
            new_sn = st.text_input("➕ Add new 'Natural Finish Slat' option", key=f"add_sn_{idx}")
            if new_sn and new_sn not in st.session_state["slat_nat_list"]:
                st.session_state["slat_nat_list"].append(new_sn)
                st.rerun()

        with col_sp:
            item["slat_pow"] = st.multiselect("Slat Type - Powder Coating", st.session_state["slat_pow_list"], default=item.get("slat_pow", []), key=f"sp_{idx}")
            new_sp = st.text_input("➕ Add new 'Powder Coating Slat' option", key=f"add_sp_{idx}")
            if new_sp and new_sp not in st.session_state["slat_pow_list"]:
                st.session_state["slat_pow_list"].append(new_sp)
                st.rerun()

        # --- GUIDE & BOTTOM & HOOD COVER ---
        st.markdown("##### 🛠️ Guide, Bottom Sheet & Hood Cover")
        cg, cb, ch_col = st.columns(3)

        with cg:
            item["guide"] = st.multiselect("Guide Specification", st.session_state["guide_list"], default=item.get("guide", []), key=f"gd_{idx}")
            new_gd = st.text_input("➕ Add new Guide option", key=f"add_gd_{idx}")
            if new_gd and new_gd not in st.session_state["guide_list"]:
                st.session_state["guide_list"].append(new_gd)
                st.rerun()

        with cb:
            item["bottom"] = st.multiselect("Bottom Sheet Specification", st.session_state["bottom_list"], default=item.get("bottom", []), key=f"bt_{idx}")
            new_bt = st.text_input("➕ Add new Bottom Sheet option", key=f"add_bt_{idx}")
            if new_bt and new_bt not in st.session_state["bottom_list"]:
                st.session_state["bottom_list"].append(new_bt)
                st.rerun()

        with ch_col:
            item["hood"] = st.multiselect("Hood Cover Specification", st.session_state["hood_list"], default=item.get("hood", []), key=f"hd_{idx}")
            new_hd = st.text_input("➕ Add new Hood Cover option", key=f"add_hd_{idx}")
            if new_hd and new_hd not in st.session_state["hood_list"]:
                st.session_state["hood_list"].append(new_hd)
                st.rerun()

        # --- LOCKS & SAFETY ---
        st.markdown("##### 🔒 Locks & Safety Features")
        item["safety_locks"] = st.multiselect(
            "Select Locks & Safety Features (Tick / Multi-select)",
            SAFETY_LOCK_OPTIONS,
            default=item.get("safety_locks", []),
            key=f"lock_{idx}"
        )

        # --- OPERATOR FOR ROLLING SHUTTER ---
        st.markdown("##### ⚙️ Operator For Rolling Shutter")
        if item["type"] == "Motorized Rolling Shutter":
            item["operator"] = st.multiselect(
                "Operator For Rolling Shutter (Select Option)",
                OPERATOR_OPTIONS,
                default=item.get("operator", []),
                key=f"op_{idx}"
            )
        else:
            st.info("🔒 Operator selection is disabled for Gear and Manual Shutters.")
            item["operator"] = []

        st.markdown("---")
        st.markdown("**Sizes & Dual Rates (Material Supply vs Installation):**")
        cw, ch, cq, cmr, cir = st.columns(5)
        item["width"] = cw.number_input("Width (mm)", value=int(item.get("width", 4000)), step=50, key=f"w_{idx}")
        item["height"] = ch.number_input("Height (mm)", value=int(item.get("height", 5000)), step=50, key=f"h_{idx}")
        item["qty"] = cq.number_input("Qty", value=int(item.get("qty", 1)), min_value=1, step=1, key=f"q_{idx}")
        item["mat_rate"] = cmr.number_input("Material Unit Rate", value=int(item.get("mat_rate", 50000)), step=500, key=f"mr_{idx}")
        item["inst_rate"] = cir.number_input("Installation Unit Rate", value=int(item.get("inst_rate", 5000)), step=100, key=f"ir_{idx}")

st.button("➕ Add Another Shutter Item", on_click=add_shutter)

# --- CHARGES & TAXES ---
st.header("🚚 Extra Charges")
col_p, col_f = st.columns(2)
packing_charges = col_p.number_input("Packing & Loading Charges (INR)", value=5000)
freight_charges = col_f.number_input("Freight Charges (INR)", value=15000)

# --- PDF GENERATOR ---
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
        m_amt = itm.get("qty", 1) * itm.get("mat_rate", 0)
        i_amt = itm.get("qty", 1) * itm.get("inst_rate", 0)
        mat_grand_total += m_amt
        inst_grand_total += i_amt
        total_qty += itm.get("qty", 1)

        desc_lines = [f"<b>{itm.get('type', 'Motorized Rolling Shutter')}</b>"]
        
        for s in itm.get("slat_nat", []):
            desc_lines.append(f"- {s}")
        for s in itm.get("slat_pow", []):
            desc_lines.append(f"- {s}")
        for g in itm.get("guide", []):
            desc_lines.append(f"- {g}")
        for b in itm.get("bottom", []):
            desc_lines.append(f"- {b}")
        for h in itm.get("hood", []):
            desc_lines.append(f"- {h}")
        for l in itm.get("safety_locks", []):
            desc_lines.append(f"- {l}")
            
        if itm.get("type") == "Motorized Rolling Shutter" and itm.get("operator"):
            desc_lines.append("<b>Operator For Rolling Shutter:</b>")
            for op in itm.get("operator", []):
                desc_lines.append(f"- {op}")

        desc = "<br/>".join(desc_lines)

        items_data.append([
            Paragraph(str(i+1), small_text),
            Paragraph(desc, small_text),
            Paragraph(itm.get("hsn", "-"), small_text),
            Paragraph(str(itm.get("width", "-")), small_text),
            Paragraph(str(itm.get("height", "-")), small_text),
            Paragraph(str(itm.get("qty", "-")), small_text),
            Paragraph(f"{itm.get('mat_rate', 0):,}", small_text),
            Paragraph(f"{m_amt:,}", small_text),
            Paragraph(f"{itm.get('inst_rate', 0):,}", small_text),
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
