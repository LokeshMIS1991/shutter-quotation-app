import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import io
import json
import os
import pypdfium2 as pdfium  # PDF to Image Rendering for Chrome Compatibility

# OpenPyXL styles for formatted Excel output
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

st.set_page_config(page_title="Sidharth Shutter & Automation Pvt. Ltd.", layout="wide", page_icon="🛡️")

# --- CUSTOM CSS FOR ENHANCED ENTERPRISE LOOK ---
st.markdown("""
<style>
    .stApp {
        background-color: #f8fafc;
    }
    
    /* Welcome Header Banner */
    .welcome-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%);
        padding: 35px 25px;
        border-radius: 16px;
        color: white;
        text-align: center;
        margin-bottom: 30px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    }
    .welcome-header h1 {
        color: #ffffff !important;
        font-weight: 800;
        margin: 0;
        font-size: 32px;
        letter-spacing: 0.5px;
    }
    .welcome-header p {
        color: #94a3b8 !important;
        margin-top: 10px;
        margin-bottom: 0;
        font-size: 16px;
    }

    /* Product Card Styling */
    .product-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        transition: transform 0.2s, box-shadow 0.2s;
        margin-bottom: 15px;
    }
    .product-card:hover {
        border-color: #2563eb;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.15);
    }
    
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# --- SESSION STATE FOR NAVIGATION ---
if "selected_product" not in st.session_state:
    st.session_state["selected_product"] = "Home"

def navigate_to(page_name):
    st.session_state["selected_product"] = page_name

# --- PERMANENT SPECIFICATIONS FILE MANAGEMENT ---
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

# ==============================================================================
# PAGE 1: HOME / LANDING PAGE
# ==============================================================================
if st.session_state["selected_product"] == "Home":

    st.markdown("""
    <div class="welcome-header">
        <h1>Welcome To Sidharth Shutter & Automation Pvt. Ltd.</h1>
        <p>Select a product line below to create professional quotations & proposals</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### 📦 Select Product For Quotation")

    products = [
        {"name": "Rolling Shutters", "desc": "Motorized, Gear & Manual Rolling Shutters", "icon": "🌀"},
        {"name": "Dock Leveler", "desc": "Hydraulic & Mechanical Dock Levelers", "icon": "🏗️"},
        {"name": "Gates", "desc": "Sliding, Telescopic, L-Folding, Swing & Retractable Gates", "icon": "🚪"},
        {"name": "Doors", "desc": "High Speed, Fire, HMPS, GPD & Overhead Sectional Doors", "icon": "🚪"},
        {"name": "Boom Barrier", "desc": "Automatic & Heavy-Duty Traffic Barriers", "icon": "🚧"},
        {"name": "Dock Shelter", "desc": "Retractable & Inflatable Dock Shelters", "icon": "🏬"},
        {"name": "Dock Bumper", "desc": "Heavy Rubber & Moulded Bumpers", "icon": "🛡️"}
    ]

    col1, col2, col3 = st.columns(3)
    cols = [col1, col2, col3]

    for idx, prod in enumerate(products):
        with cols[idx % 3]:
            st.markdown(f"""
            <div class="product-card">
                <h3 style="margin-bottom: 5px; color: #1e293b;">{prod['icon']} {prod['name']}</h3>
                <p style="color: #64748b; font-size: 13px; min-height: 38px;">{prod['desc']}</p>
            </div>
            """, unsafe_allow_html=True)
            
            if st.button(f"Generate Quotation ->", key=f"btn_{idx}", use_container_width=True, type="primary" if idx==0 else "secondary"):
                st.session_state["selected_product"] = prod["name"]
                st.rerun()

# ==============================================================================
# PAGE 2: ROLLING SHUTTERS QUOTATION PAGE
# ==============================================================================
elif st.session_state["selected_product"] == "Rolling Shutters":

    col_nav1, col_nav2 = st.columns([1, 6])
    with col_nav1:
        if st.button("🏠 Home", use_container_width=True):
            st.session_state["selected_product"] = "Home"
            st.rerun()
    with col_nav2:
        st.title("🌀 Rolling Shutters Quotation Builder")

    st.sidebar.header("📋 Client & Proposal Details")

    client_id = st.sidebar.text_input("Client ID", "SSAPL-CL-8842")
    client_name = st.sidebar.text_input("Company Name", "M/S ABHINAV INFRA BUILD PVT. LTD.")
    contact_details = st.sidebar.text_input("Contact Details", "Mob:-8889911529")
    shipping_address = st.sidebar.text_area("Shipping Address", "RGLP PITHAMPUR Sector-1 Pithampur MADHYA PRADESH 454774", height=80)
    site_person = st.sidebar.text_input("Site Person Mobile", "9926952398")
    billing_address = st.sidebar.text_area("Billing Address", "207,208, Industry House, Indore, MP 452001", height=80)
    gstin = st.sidebar.text_input("Client GSTIN", "23AAHCA9425D1ZY")

    st.sidebar.markdown("---")
    c_date, c_ref = st.sidebar.columns(2)
    qtn_date = c_date.date_input("Quotation Date")
    qtn_ref_no = c_ref.text_input("Qtn Ref No", "SSAPL/2026-27/319R2")

    sales_reference_1 = st.sidebar.text_input("Sales Reference 1", "Mr. Nishant (90010 42914)")
    sales_reference_2 = st.sidebar.text_input("Sales Reference 2", "Mr. Jeevan Sharma (9828771899)")

    with st.expander("📥 Bulk Import Specifications", expanded=False):
        st.markdown("Paste your complete list below. Add each option on a **new line**.")
        col_cat, col_txt = st.columns([1, 2])
        category_mapping = {
            "Natural Finish Slats": "slat_nat_list",
            "Powder Coated Slats": "slat_pow_list",
            "Guide Specifications": "guide_list",
            "Bottom Specifications": "bottom_list",
            "Hood Cover Specifications": "hood_list"
        }
        target_category = col_cat.selectbox("Select Target List:", list(category_mapping.keys()))
        bulk_text = col_txt.text_area("Paste Options (Line by Line):", height=120, placeholder="Option 1\nOption 2\nOption 3...")
        
        if st.button("🚀 Import Options in Bulk", type="primary"):
            if bulk_text.strip():
                lines = [line.strip() for line in bulk_text.split("\n") if line.strip()]
                target_key = category_mapping[target_category]
                added_count = 0
                for item in lines:
                    if item not in st.session_state[target_key]:
                        st.session_state[target_key].append(item)
                        added_count += 1
                if added_count > 0:
                    save_specifications()
                    st.success(f"✅ Successfully added {added_count} new options to {target_category}!")
                    st.rerun()
                else:
                    st.warning("⚠️ All pasted options already exist in the list.")
            else:
                st.error("Please enter some text to import.")

    st.markdown("### 🧱 Shutter Specifications & Pricing")

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
        with st.expander(f"📌 Item #{idx + 1}: {item.get('type', 'Motorized Rolling Shutter')}", expanded=True):
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
                    st.button("❌ Remove Item", key=f"del_{idx}", on_click=remove_shutter, args=(idx,))

            item["hsn"] = st.text_input("HSN Code", value=item.get("hsn", "73083000"), key=f"hsn_{idx}")

            st.markdown("##### 📐 Slat Specifications")
            col_sn, col_sp = st.columns(2)
            with col_sn:
                item["slat_nat"] = st.multiselect("Natural Finish Slats", st.session_state["slat_nat_list"], default=item.get("slat_nat", []), key=f"sn_{idx}")
                col_in_sn, col_btn_sn = st.columns([3, 1])
                new_sn = col_in_sn.text_input("➕ New Natural Option", key=f"add_sn_{idx}", label_visibility="collapsed", placeholder="Add new natural slat...")
                if col_btn_sn.button("Add Option", key=f"btn_add_sn_{idx}"):
                    if new_sn and new_sn not in st.session_state["slat_nat_list"]:
                        st.session_state["slat_nat_list"].append(new_sn)
                        save_specifications()
                        st.rerun()

            with col_sp:
                item["slat_pow"] = st.multiselect("Powder Coated Slats", st.session_state["slat_pow_list"], default=item.get("slat_pow", []), key=f"sp_{idx}")
                col_in_sp, col_btn_sp = st.columns([3, 1])
                new_sp = col_in_sp.text_input("➕ New Powder Option", key=f"add_sp_{idx}", label_visibility="collapsed", placeholder="Add new powder slat...")
                if col_btn_sp.button("Add Option", key=f"btn_add_sp_{idx}"):
                    if new_sp and new_sp not in st.session_state["slat_pow_list"]:
                        st.session_state["slat_pow_list"].append(new_sp)
                        save_specifications()
                        st.rerun()

            st.markdown("##### 🛠️ Guide, Bottom Sheet & Hood Cover")
            cg, cb, ch_col = st.columns(3)
            with cg:
                item["guide"] = st.multiselect("Guide Specification", st.session_state["guide_list"], default=item.get("guide", []), key=f"gd_{idx}")
                col_in_gd, col_btn_gd = st.columns([2, 1])
                new_gd = col_in_gd.text_input("➕ New Guide", key=f"add_gd_{idx}", label_visibility="collapsed", placeholder="Add guide...")
                if col_btn_gd.button("Add", key=f"btn_add_gd_{idx}"):
                    if new_gd and new_gd not in st.session_state["guide_list"]:
                        st.session_state["guide_list"].append(new_gd)
                        save_specifications()
                        st.rerun()

            with cb:
                item["bottom"] = st.multiselect("Bottom Specification", st.session_state["bottom_list"], default=item.get("bottom", []), key=f"bt_{idx}")
                col_in_bt, col_btn_bt = st.columns([2, 1])
                new_bt = col_in_bt.text_input("➕ New Bottom", key=f"add_bt_{idx}", label_visibility="collapsed", placeholder="Add bottom...")
                if col_btn_bt.button("Add", key=f"btn_add_bt_{idx}"):
                    if new_bt and new_bt not in st.session_state["bottom_list"]:
                        st.session_state["bottom_list"].append(new_bt)
                        save_specifications()
                        st.rerun()

            with ch_col:
                item["hood"] = st.multiselect("Hood Cover Specification", st.session_state["hood_list"], default=item.get("hood", []), key=f"hd_{idx}")
                col_in_hd, col_btn_hd = st.columns([2, 1])
                new_hd = col_in_hd.text_input("➕ New Hood", key=f"add_hd_{idx}", label_visibility="collapsed", placeholder="Add hood...")
                if col_btn_hd.button("Add", key=f"btn_add_hd_{idx}"):
                    if new_hd and new_hd not in st.session_state["hood_list"]:
                        st.session_state["hood_list"].append(new_hd)
                        save_specifications()
                        st.rerun()

            st.markdown("##### 🔒 Locks & Safety Features")
            item["safety_locks"] = st.multiselect(
                "Select Locks & Safety Features",
                SAFETY_LOCK_OPTIONS,
                default=item.get("safety_locks", []),
                key=f"lock_{idx}"
            )

            st.markdown("##### ⚙️ Operator For Rolling Shutter")
            if item["type"] == "Motorized Rolling Shutter":
                item["operator"] = st.multiselect(
                    "Operator Option",
                    OPERATOR_OPTIONS,
                    default=item.get("operator", []),
                    key=f"op_{idx}"
                )
            else:
                st.info("🔒 Operator selection disabled for Gear and Manual Shutters.")
                item["operator"] = []

            st.markdown("---")
            st.markdown("**Dimensions & Rates:**")
            cw, ch, cq, cmr, cir = st.columns(5)
            item["width"] = cw.number_input("Width (mm)", value=int(item.get("width", 4000)), step=50, key=f"w_{idx}")
            item["height"] = ch.number_input("Height (mm)", value=int(item.get("height", 5000)), step=50, key=f"h_{idx}")
            item["qty"] = cq.number_input("Qty", value=int(item.get("qty", 1)), min_value=1, step=1, key=f"q_{idx}")
            item["mat_rate"] = cmr.number_input("Material Rate (INR)", value=int(item.get("mat_rate", 50000)), step=500, key=f"mr_{idx}")
            item["inst_rate"] = cir.number_input("Installation Rate (INR)", value=int(item.get("inst_rate", 5000)), step=100, key=f"ir_{idx}")

    st.button("➕ Add Another Shutter Item", on_click=add_shutter)

    st.markdown("### 🚚 Extra Charges & Expenses")
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
        title_style = ParagraphStyle('Title', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, alignment=1, textColor=colors.HexColor('#1A365D'))
        small_bold = ParagraphStyle('SmallBold', fontName='Helvetica-Bold', fontSize=7.5, leading=9)
        small_bold_right = ParagraphStyle('SmallBoldRight', fontName='Helvetica-Bold', fontSize=7.5, leading=9, alignment=2)
        small_bold_center = ParagraphStyle('SmallBoldCenter', fontName='Helvetica-Bold', fontSize=7.5, leading=9, alignment=1)
        small_text = ParagraphStyle('SmallText', fontName='Helvetica', fontSize=7.5, leading=9)
        small_text_right = ParagraphStyle('SmallTextRight', fontName='Helvetica', fontSize=7.5, leading=9, alignment=2)

        story.append(Paragraph("SIDHARTH SHUTTER & AUTOMATION PRIVATE LIMITED", title_style))
        story.append(Paragraph("G-1-66, Industrial Area, Prahaladpura, Sanganer, Jaipur, Rajasthan, 303903", ParagraphStyle('Sub', alignment=1, fontSize=7.5)))
        story.append(Spacer(1, 8))

        header_data = [
            [Paragraph("<b>Company Name</b>", small_bold), Paragraph(client_name, small_text), Paragraph("<b>Client ID:</b>", small_bold), Paragraph(client_id, small_bold)],
            [Paragraph("<b>Contact Details</b>", small_bold), Paragraph(contact_details, small_text), Paragraph("<b>Qtn Date:</b>", small_bold), Paragraph(str(qtn_date), small_text)],
            [Paragraph("<b>Shipping Address</b>", small_bold), Paragraph(shipping_address, small_text), Paragraph("<b>Qtn Ref No:</b>", small_bold), Paragraph(qtn_ref_no, small_bold)],
            [Paragraph("<b>Billing Address</b>", small_bold), Paragraph(billing_address, small_text), Paragraph("<b>Sales Reference:</b>", small_bold), Paragraph(f"{sales_reference_1}<br/>{sales_reference_2}", small_text)],
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
            Paragraph("<b>Sr. No.</b>", small_bold_center),
            Paragraph("<b>Description</b>", small_bold),
            Paragraph("<b>HSN Code</b>", small_bold_center),
            Paragraph("<b>Width (mm)</b>", small_bold_center),
            Paragraph("<b>Height (mm)</b>", small_bold_center),
            Paragraph("<b>Qty</b>", small_bold_center),
            Paragraph("<b>Mat. Rate (INR)</b>", small_bold_right),
            Paragraph("<b>Mat. Amount (INR)</b>", small_bold_right),
            Paragraph("<b>Inst. Rate (INR)</b>", small_bold_right),
            Paragraph("<b>Inst. Amount (INR)</b>", small_bold_right)
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
            for s in itm.get("slat_nat", []): desc_lines.append(f"- {s}")
            for s in itm.get("slat_pow", []): desc_lines.append(f"- {s}")
            for g in itm.get("guide", []): desc_lines.append(f"- {g}")
            for b in itm.get("bottom", []): desc_lines.append(f"- {b}")
            for h in itm.get("hood", []): desc_lines.append(f"- {h}")
            for l in itm.get("safety_locks", []): desc_lines.append(f"- {l}")
                
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
                Paragraph(f"{itm.get('mat_rate', 0):,}", small_text_right),
                Paragraph(f"{m_amt:,}", small_text_right),
                Paragraph(f"{itm.get('inst_rate', 0):,}", small_text_right),
                Paragraph(f"{i_amt:,}", small_text_right)
            ])

        subtotal_extras = packing_charges + freight_charges + unloading_charges + crane_charges + scaffolding_charges
        subtotal_mat = mat_grand_total + subtotal_extras
        gst_mat = round(subtotal_mat * 0.18)
        gst_inst = round(inst_grand_total * 0.18)
        
        total_mat_with_gst = subtotal_mat + gst_mat
        total_inst_with_gst = inst_grand_total + gst_inst
        final_grand_total = total_mat_with_gst + total_inst_with_gst

        items_data.append(["", Paragraph("<b>Item Total</b>", small_bold), "", "", "", Paragraph(f"<b>{total_qty}</b>", small_bold_center), "", Paragraph(f"<b>{mat_grand_total:,}</b>", small_bold_right), "", Paragraph(f"<b>{inst_grand_total:,}</b>", small_bold_right)])
        items_data.append(["", Paragraph("Packing Charges", small_text), "", "", "", "", "", Paragraph(f"{packing_charges:,}", small_text_right), "", Paragraph("-", small_text_right)])
        items_data.append(["", Paragraph("Freight Charges", small_text), "", "", "", "", "", Paragraph(f"{freight_charges:,}", small_text_right), "", Paragraph("-", small_text_right)])
        if unloading_charges > 0:
            items_data.append(["", Paragraph("Unloading Charges", small_text), "", "", "", "", "", Paragraph(f"{unloading_charges:,}", small_text_right), "", Paragraph("-", small_text_right)])
        if crane_charges > 0:
            items_data.append(["", Paragraph("Crane Charges", small_text), "", "", "", "", "", Paragraph(f"{crane_charges:,}", small_text_right), "", Paragraph("-", small_text_right)])
        if scaffolding_charges > 0:
            items_data.append(["", Paragraph("Scaffolding Charges", small_text), "", "", "", "", "", Paragraph(f"{scaffolding_charges:,}", small_text_right), "", Paragraph("-", small_text_right)])

        items_data.append(["", Paragraph("<b>Supply & Installation Amount Excluding GST</b>", small_bold), "", "", "", "", "", Paragraph(f"<b>{subtotal_mat:,}</b>", small_bold_right), "", Paragraph(f"<b>{inst_grand_total:,}</b>", small_bold_right)])
        items_data.append(["", Paragraph("GST on supply & installation @18%", small_text), "", "", "", "", "", Paragraph(f"{gst_mat:,}", small_text_right), Paragraph("@18%", small_text_right), Paragraph(f"{gst_inst:,}", small_text_right)])
        items_data.append(["", Paragraph("<b>Total supply & installation with GST</b>", small_bold), "", "", "", "", "", Paragraph(f"<b>{total_mat_with_gst:,}</b>", small_bold_right), "", Paragraph(f"<b>{total_inst_with_gst:,}</b>", small_bold_right)])
        items_data.append(["", Paragraph("<b>Grand total with GST</b>", small_bold), "", "", "", "", "", "", "", Paragraph(f"<b>{final_grand_total:,}</b>", small_bold_right)])

        t_items = Table(items_data, colWidths=[22, 170, 50, 38, 38, 22, 55, 60, 50, 60])
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

    # --- FORMATTED EXCEL GENERATOR ---
    def generate_excel():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Quotation"

        font_company = Font(name="Calibri", size=14, bold=True, color="1A365D")
        font_address = Font(name="Calibri", size=9, italic=True)
        font_header_bold = Font(name="Calibri", size=10, bold=True)
        font_regular = Font(name="Calibri", size=9)
        font_bold = Font(name="Calibri", size=9, bold=True)
        
        fill_table_header = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
        fill_total_row = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

        thin_border = Border(
            left=Side(style='thin', color='D3D3D3'),
            right=Side(style='thin', color='D3D3D3'),
            top=Side(style='thin', color='D3D3D3'),
            bottom=Side(style='thin', color='D3D3D3')
        )
        
        align_center = Alignment(horizontal="center", vertical="top", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="top", wrap_text=True)
        align_right = Alignment(horizontal="right", vertical="top", wrap_text=True)

        ws.merge_cells("A1:J1")
        ws["A1"] = "SIDHARTH SHUTTER & AUTOMATION PRIVATE LIMITED"
        ws["A1"].font = font_company
        ws["A1"].alignment = Alignment(horizontal="center")

        ws.merge_cells("A2:J2")
        ws["A2"] = "G-1-66, Industrial Area, Prahaladpura, Sanganer, Jaipur, Rajasthan, 303903"
        ws["A2"].font = font_address
        ws["A2"].alignment = Alignment(horizontal="center")

        ws.row_dimensions[1].height = 25
        ws.row_dimensions[2].height = 18

        client_info = [
            [("Company Name:", font_header_bold), (client_name, font_regular), ("Client ID:", font_header_bold), (client_id, font_header_bold)],
            [("Contact Details:", font_header_bold), (contact_details, font_regular), ("Qtn Date:", font_header_bold), (str(qtn_date), font_regular)],
            [("Shipping Address:", font_header_bold), (shipping_address, font_regular), ("Qtn Ref No:", font_header_bold), (qtn_ref_no, font_header_bold)],
            [("Billing Address:", font_header_bold), (billing_address, font_regular), ("Sales Reference:", font_header_bold), (f"{sales_reference_1} / {sales_reference_2}", font_regular)],
            [("Client GSTIN:", font_header_bold), (gstin, font_regular), ("", font_regular), ("", font_regular)]
        ]

        curr_row = 4
        for row in client_info:
            ws.cell(row=curr_row, column=1, value=row[0][0]).font = row[0][1]
            
            ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=5)
            ws.cell(row=curr_row, column=2, value=row[1][0]).font = row[1][1]
            ws.cell(row=curr_row, column=2).alignment = align_left

            ws.cell(row=curr_row, column=6, value=row[2][0]).font = row[2][1]
            
            ws.merge_cells(start_row=curr_row, start_column=7, end_row=curr_row, end_column=10)
            ws.cell(row=curr_row, column=7, value=row[3][0]).font = row[3][1]
            ws.cell(row=curr_row, column=7).alignment = align_left
            
            for c in range(1, 11):
                ws.cell(row=curr_row, column=c).border = thin_border
            
            curr_row += 1

        curr_row += 1

        headers = ["Sr. No.", "Description", "HSN Code", "Width (mm)", "Height (mm)", "Qty", "Mat. Rate (INR)", "Mat. Amount (INR)", "Inst. Rate (INR)", "Inst. Amount (INR)"]
        for col_num, h_text in enumerate(headers, 1):
            cell = ws.cell(row=curr_row, column=col_num, value=h_text)
            cell.font = font_header_bold
            cell.fill = fill_table_header
            cell.alignment = align_center
            cell.border = thin_border
        
        ws.row_dimensions[curr_row].height = 25
        curr_row += 1

        mat_grand_total = 0
        inst_grand_total = 0
        total_qty = 0

        for i, itm in enumerate(st.session_state["shutter_items"]):
            m_amt = itm.get("qty", 1) * itm.get("mat_rate", 0)
            i_amt = itm.get("qty", 1) * itm.get("inst_rate", 0)
            mat_grand_total += m_amt
            inst_grand_total += i_amt
            total_qty += itm.get("qty", 1)

            desc_lines = [itm.get('type', 'Motorized Rolling Shutter')]
            for s in itm.get("slat_nat", []): desc_lines.append(f"• {s}")
            for s in itm.get("slat_pow", []): desc_lines.append(f"• {s}")
            for g in itm.get("guide", []): desc_lines.append(f"• {g}")
            for b in itm.get("bottom", []): desc_lines.append(f"• {b}")
            for h in itm.get("hood", []): desc_lines.append(f"• {h}")
            for l in itm.get("safety_locks", []): desc_lines.append(f"• {l}")
            
            if itm.get("type") == "Motorized Rolling Shutter" and itm.get("operator"):
                desc_lines.append("Operator For Rolling Shutter:")
                for op in itm.get("operator", []):
                    desc_lines.append(f"• {op}")

            desc_str = "\n".join(desc_lines)

            row_data = [
                i + 1,
                desc_str,
                itm.get("hsn", "-"),
                itm.get("width", "-"),
                itm.get("height", "-"),
                itm.get("qty", "-"),
                itm.get("mat_rate", 0),
                m_amt,
                itm.get("inst_rate", 0),
                i_amt
            ]

            for col_num, val in enumerate(row_data, 1):
                cell = ws.cell(row=curr_row, column=col_num, value=val)
                cell.font = font_regular
                cell.border = thin_border
                
                if col_num in [1, 3, 4, 5, 6]:
                    cell.alignment = align_center
                elif col_num in [7, 8, 9, 10]:
                    cell.alignment = align_right
                    cell.number_format = '#,##0'
                else:
                    cell.alignment = align_left

            curr_row += 1

        subtotal_extras = packing_charges + freight_charges + unloading_charges + crane_charges + scaffolding_charges
        subtotal_mat = mat_grand_total + subtotal_extras
        gst_mat = round(subtotal_mat * 0.18)
        gst_inst = round(inst_grand_total * 0.18)
        
        total_mat_with_gst = subtotal_mat + gst_mat
        total_inst_with_gst = inst_grand_total + gst_inst
        final_grand_total = total_mat_with_gst + total_inst_with_gst

        def add_summary_row(label, mat_val, inst_val, is_bold=False, is_fill=False):
            nonlocal curr_row
            ws.cell(row=curr_row, column=2, value=label).font = font_bold if is_bold else font_regular
            ws.cell(row=curr_row, column=2).alignment = align_left
            
            if label == "Item Total":
                ws.cell(row=curr_row, column=6, value=total_qty).font = font_bold
                ws.cell(row=curr_row, column=6).alignment = align_center

            cell_mat = ws.cell(row=curr_row, column=8, value=mat_val if isinstance(mat_val, (int, float)) else mat_val)
            cell_mat.font = font_bold if is_bold else font_regular
            cell_mat.alignment = align_right
            if isinstance(mat_val, (int, float)): cell_mat.number_format = '#,##0'

            cell_inst = ws.cell(row=curr_row, column=10, value=inst_val if isinstance(inst_val, (int, float)) else inst_val)
            cell_inst.font = font_bold if is_bold else font_regular
            cell_inst.alignment = align_right
            if isinstance(inst_val, (int, float)): cell_inst.number_format = '#,##0'

            for c in range(1, 11):
                cell = ws.cell(row=curr_row, column=c)
                cell.border = thin_border
                if is_fill:
                    cell.fill = fill_total_row
            curr_row += 1

        add_summary_row("Item Total", mat_grand_total, inst_grand_total, is_bold=True)
        add_summary_row("Packing Charges", packing_charges, "-")
        add_summary_row("Freight Charges", freight_charges, "-")
        if unloading_charges > 0: add_summary_row("Unloading Charges", unloading_charges, "-")
        if crane_charges > 0: add_summary_row("Crane Charges", crane_charges, "-")
        if scaffolding_charges > 0: add_summary_row("Scaffolding Charges", scaffolding_charges, "-")

        add_summary_row("Supply & Installation Amount Excluding GST", subtotal_mat, inst_grand_total, is_bold=True, is_fill=True)
        add_summary_row("GST on supply & installation @18%", gst_mat, gst_inst)
        add_summary_row("Total supply & installation with GST", total_mat_with_gst, total_inst_with_gst, is_bold=True, is_fill=True)
        add_summary_row("Grand total with GST", "", final_grand_total, is_bold=True, is_fill=True)

        col_widths = {1: 8, 2: 45, 3: 12, 4: 12, 5: 12, 6: 8, 7: 15, 8: 18, 9: 15, 10: 18}
        for col_idx, width in col_widths.items():
            ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    # --- DOWNLOAD & PREVIEW BUTTONS SECTION ---
    st.markdown("---")
    st.markdown("### 📥 Generate & Export Proposals")

    col_pdf, col_excel = st.columns(2)

    with col_pdf:
        if st.button("👁️ Preview PDF Quotation", type="primary", use_container_width=True):
            pdf_bytes = generate_pdf().getvalue()
            st.session_state["pdf_preview_bytes"] = pdf_bytes

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

    # --- CHROME-SAFE PDF PREVIEW AND DOWNLOAD SECTION ---
    if "pdf_preview_bytes" in st.session_state:
        st.markdown("---")
        st.markdown("### 🔍 PDF Quotation Preview & Download")
        
        pdf_data = st.session_state["pdf_preview_bytes"]
        
        c_dl, c_cls = st.columns([2, 1])
        with c_dl:
            st.download_button(
                label="📥 Download Verified PDF Quotation",
                data=pdf_data,
                file_name=f"Quotation_{qtn_ref_no.replace('/', '_')}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True
            )
        with c_cls:
            if st.button("❌ Close Preview", use_container_width=True):
                del st.session_state["pdf_preview_bytes"]
                st.rerun()

        # Render PDF pages as High-Resolution Images (Fixed for latest Streamlit version)
        pdf = pdfium.PdfDocument(pdf_data)
        for i, page in enumerate(pdf):
            image = page.render(scale=2).to_pil()
            st.image(image, caption=f"Page {i+1}", use_container_width=True)

# ==============================================================================
# PAGE 3: OTHER PRODUCTS (PLACEHOLDER)
# ==============================================================================
else:
    col_nav1, col_nav2 = st.columns([1, 6])
    with col_nav1:
        if st.button("🏠 Home", use_container_width=True):
            st.session_state["selected_product"] = "Home"
            st.rerun()
    with col_nav2:
        st.title(f"🛠️ {st.session_state['selected_product']} Quotation Module")

    st.info(f"⏳ **{st.session_state['selected_product']} Quotation Builder** is currently under development. Please check back soon!")
