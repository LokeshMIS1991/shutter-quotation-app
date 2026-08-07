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
st.subheader("Automated Quotation Builder")

# --- PERMANENT SPECIFICATIONS FILE MANAGEMENT (Change 1) ---
SPEC_FILE = "specifications_master.json"

DEFAULT_SPECS = {
    "slat_nat_list": ["90mm (H) x 1.2 mm thick Galvalume Plain slats in natural finish"],
    "slat_pow_list": ["90mm (H) x 1.2 mm thick Galvalume Powder Coated slats"],
    "guide_list": ["TG Guide with rubber seal with grey epoxy"],
    "bottom_list": ["Super bottom with rubber seal with grey epoxy"],
    "hood_list": [".80mm thick Galvalume Hood & Motor cover in natural finish"]
}

def load_specifications():
    if os.path.exists(SPEC_FILE):
        try:
            with open(SPEC_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                # Ensure all required keys exist
                for k, v in DEFAULT_SPECS.items():
                    if k not in data:
                        data[k] = v
                return data
        except Exception:
            return DEFAULT_SPECS.copy()
    return DEFAULT_SPECS.copy()

def save_specifications():
    data = {
        "slat_nat_list": st.session_state["slat_nat_list"],
        "slat_pow_list": st.session_state["slat_pow_list"],
        "guide_list": st.session_state["guide_list"],
        "bottom_list": st.session_state["bottom_list"],
        "hood_list": st.session_state["hood_list"]
    }
    with open(SPEC_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

# Initialize Session State Specs from file
specs = load_specifications()
for key, value in specs.items():
    if key not in st.session_state:
        st.session_state[key] = value

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
# Change 2: Sales Reference
sales_reference_1 = st.sidebar.text_input("Sales Reference 1", "Mr. Nishant (90010 42914)")
sales_reference_2 = st.sidebar.text_input("Sales Reference 2", "Mr. Jeevan Sharma (9828771899)")

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
            col_in_sn, col_btn_sn = st.columns([3, 1])
            new_sn = col_in_sn.text_input("➕ New Natural Slat Option", key=f"add_sn_{idx}", label_visibility="collapsed", placeholder="Type new natural slat specification...")
            if col_btn_sn.button("Add Option", key=f"btn_add_sn_{idx}"):
                if new_sn and new_sn not in st.session_state["slat_nat_list"]:
                    st.session_state["slat_nat_list"].append(new_sn)
                    save_specifications()
                    st.rerun()

        with col_sp:
            item["slat_pow"] = st.multiselect("Slat Type - Powder Coating", st.session_state["slat_pow_list"], default=item.get("slat_pow", []), key=f"sp_{idx}")
            col_in_sp, col_btn_sp = st.columns([3, 1])
            new_sp = col_in_sp.text_input("➕ New Powder Slat Option", key=f"add_sp_{idx}", label_visibility="collapsed", placeholder="Type new powder slat specification...")
            if col_btn_sp.button("Add Option", key=f"btn_add_sp_{idx}"):
                if new_sp and new_sp not in st.session_state["slat_pow_list"]:
                    st.session_state["slat_pow_list"].append(new_sp)
                    save_specifications()
                    st.rerun()

        # --- GUIDE & BOTTOM & HOOD COVER ---
        st.markdown("##### 🛠️ Guide, Bottom Sheet & Hood Cover")
        cg, cb, ch_col = st.columns(3)

        with cg:
            item["guide"] = st.multiselect("Guide Specification", st.session_state["guide_list"], default=item.get("guide", []), key=f"gd_{idx}")
            col_in_gd, col_btn_gd = st.columns([2, 1])
            new_gd = col_in_gd.text_input("➕ New Guide", key=f"add_gd_{idx}", label_visibility="collapsed", placeholder="Type new guide...")
            if col_btn_gd.button("Add Option", key=f"btn_add_gd_{idx}"):
                if new_gd and new_gd not in st.session_state["guide_list"]:
                    st.session_state["guide_list"].append(new_gd)
                    save_specifications()
                    st.rerun()

        with cb:
            item["bottom"] = st.multiselect("Bottom Sheet Specification", st.session_state["bottom_list"], default=item.get("bottom", []), key=f"bt_{idx}")
            col_in_bt, col_btn_bt = st.columns([2, 1])
            new_bt = col_in_bt.text_input("➕ New Bottom", key=f"add_bt_{idx}", label_visibility="collapsed", placeholder="Type new bottom...")
            if col_btn_bt.button("Add Option", key=f"btn_add_bt_{idx}"):
                if new_bt and new_bt not in st.session_state["bottom_list"]:
                    st.session_state["bottom_list"].append(new_bt)
                    save_specifications()
                    st.rerun()

        with ch_col:
            item["hood"] = st.multiselect("Hood Cover Specification", st.session_state["hood_list"], default=item.get("hood", []), key=f"hd_{idx}")
            col_in_hd, col_btn_hd = st.columns([2, 1])
            new_hd = col_in_hd.text_input("➕ New Hood", key=f"add_hd_{idx}", label_visibility="collapsed", placeholder="Type new hood...")
            if col_btn_hd.button("Add Option", key=f"btn_add_hd_{idx}"):
                if new_hd and new_hd not in st.session_state["hood_list"]:
                    st.session_state["hood_list"].append(new_hd)
                    save_specifications()
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

# --- CHARGES & TAXES (Change 3: Added New Charges) ---
st.header("🚚 Extra Charges")
c1, c2, c3, c4, c5 = st.columns(5)
packing_charges = c1.number_input("Packing & Loading (INR)", value=5000)
freight_charges = c2.number_input("Freight Charges (INR)", value=15000)
unloading_charges = c3.number_input("Unloading Charges (INR)", value=0)
crane_charges = c4.number_input("Crane Charges (INR)", value=0)
scaffolding_charges = c5.number_input("Scaffolding Charges (INR)", value=0)

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
        [Paragraph("<b>Contact Details</b>", small_bold), Paragraph(contact_details, small_text), Paragraph("<b>Sales Reference:</b>", small_bold), Paragraph(f"{sales_reference_1}<br/>{sales_reference_2}", small_text)],
        [Paragraph("<b>Shipping Address</b>", small_bold), Paragraph(shipping_address, small_text), Paragraph("<b>Qtn Ref No:</b>", small_bold), Paragraph(qtn_ref_no, small_bold)],
        [Paragraph("<b>Billing Address</b>", small_bold), Paragraph(billing_address, small_text), Paragraph("", small_text), Paragraph("", small_text)],
        [Paragraph("<b>GSTIN</b>", small_bold), Paragraph(gstin, small_text), Paragraph("", small_text), Paragraph("", small_text)]
    ]
    t_header = Table(header_data, colWidths=[80, 250, 80, 155])
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

    subtotal_extras = packing_charges + freight_charges + unloading_charges + crane_charges + scaffolding_charges
    subtotal_mat = mat_grand_total + subtotal_extras
    gst_mat = round(subtotal_mat * 0.18)
    gst_inst = round(inst_grand_total * 0.18)
    
    total_mat_with_gst = subtotal_mat + gst_mat
    total_inst_with_gst = inst_grand_total + gst_inst
    final_grand_total = total_mat_with_gst + total_inst_with_gst

    items_data.append(["", Paragraph("<b>Item Total</b>", small_bold), "", "", "", Paragraph(f"<b>{total_qty}</b>", small_bold), "<b>TOTAL</b>", Paragraph(f"<b>{mat_grand_total:,}</b>", small_bold), "", Paragraph(f"<b>{inst_grand_total:,}</b>", small_bold)])
    items_data.append(["", Paragraph("Packing Charges", small_text), "", "", "", "", "", Paragraph(f"{packing_charges:,}", small_text), "", "-"])
    items_data.append(["", Paragraph("Freight Charges", small_text), "", "", "", "", "", Paragraph(f"{freight_charges:,}", small_text), "", "-"])
    if unloading_charges > 0:
        items_data.append(["", Paragraph("Unloading Charges", small_text), "", "", "", "", "", Paragraph(f"{unloading_charges:,}", small_text), "", "-"])
    if crane_charges > 0:
        items_data.append(["", Paragraph("Crane Charges", small_text), "", "", "", "", "", Paragraph(f"{crane_charges:,}", small_text), "", "-"])
    if scaffolding_charges > 0:
        items_data.append(["", Paragraph("Scaffolding Charges", small_text), "", "", "", "", "", Paragraph(f"{scaffolding_charges:,}", small_text), "", "-"])

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

# --- EXCEL GENERATOR (Change 4: Excel Download) ---
def generate_excel():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Header Info Sheet
        header_df = pd.DataFrame([
            ["Company Name", client_name, "Qtn Date", str(qtn_date)],
            ["Contact Details", contact_details, "Sales Reference 1", sales_reference_1],
            ["Shipping Address", shipping_address, "Sales Reference 2", sales_reference_2],
            ["Billing Address", billing_address, "Qtn Ref No", qtn_ref_no],
            ["GSTIN", gstin, "", ""]
        ], columns=["Field", "Value", "Field", "Value"])
        header_df.to_excel(writer, sheet_name="Client Info", index=False)

        # Items Sheet
        items_rows = []
        mat_grand_total = 0
        inst_grand_total = 0

        for i, itm in enumerate(st.session_state["shutter_items"]):
            m_amt = itm.get("qty", 1) * itm.get("mat_rate", 0)
            i_amt = itm.get("qty", 1) * itm.get("inst_rate", 0)
            mat_grand_total += m_amt
            inst_grand_total += i_amt

            specs_combined = ", ".join(
                itm.get("slat_nat", []) + 
                itm.get("slat_pow", []) + 
                itm.get("guide", []) + 
                itm.get("bottom", []) + 
                itm.get("hood", []) + 
                itm.get("safety_locks", []) + 
                itm.get("operator", [])
            )

            items_rows.append({
                "Sr. No.": i + 1,
                "Category": itm.get("type", ""),
                "Specifications": specs_combined,
                "HSN Code": itm.get("hsn", ""),
                "Width (mm)": itm.get("width", 0),
                "Height (mm)": itm.get("height", 0),
                "Qty": itm.get("qty", 1),
                "Material Rate (INR)": itm.get("mat_rate", 0),
                "Material Amount (INR)": m_amt,
                "Installation Rate (INR)": itm.get("inst_rate", 0),
                "Installation Amount (INR)": i_amt
            })

        items_df = pd.DataFrame(items_rows)
        items_df.to_excel(writer, sheet_name="Quotation Items", index=False)

        # Summary Sheet
        subtotal_extras = packing_charges + freight_charges + unloading_charges + crane_charges + scaffolding_charges
        subtotal_mat = mat_grand_total + subtotal_extras
        gst_mat = round(subtotal_mat * 0.18)
        gst_inst = round(inst_grand_total * 0.18)
        
        summary_df = pd.DataFrame([
            ["Material Total", mat_grand_total],
            ["Packing & Loading Charges", packing_charges],
            ["Freight Charges", freight_charges],
            ["Unloading Charges", unloading_charges],
            ["Crane Charges", crane_charges],
            ["Scaffolding Charges", scaffolding_charges],
            ["Subtotal Material (Excl. GST)", subtotal_mat],
            ["Installation Total (Excl. GST)", inst_grand_total],
            ["18% GST on Material", gst_mat],
            ["18% GST on Installation", gst_inst],
            ["Total Material with GST", subtotal_mat + gst_mat],
            ["Total Installation with GST", inst_grand_total + gst_inst],
            ["GRAND TOTAL WITH GST", subtotal_mat + gst_mat + inst_grand_total + gst_inst]
        ], columns=["Description", "Amount (INR)"])
        summary_df.to_excel(writer, sheet_name="Cost Summary", index=False)

    output.seek(0)
    return output

# --- DOWNLOAD BUTTONS ---
st.markdown("---")
col_pdf, col_excel = st.columns(2)

with col_pdf:
    if st.button("📄 Generate PDF Quotation", type="primary", use_container_width=True):
        pdf_buffer = generate_pdf()
        st.download_button(
            label="📥 Download Quotation PDF",
            data=pdf_buffer,
            file_name=f"Quotation_{qtn_ref_no.replace('/', '_')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )

with col_excel:
    if st.button("📊 Generate Excel Quotation", use_container_width=True):
        excel_buffer = generate_excel()
        st.download_button(
            label="📥 Download Quotation Excel (.xlsx)",
            data=excel_buffer,
            file_name=f"Quotation_{qtn_ref_no.replace('/', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
