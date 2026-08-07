import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import io

st.set_page_config(page_title="Sidharth Shutters - Quotation Builder", layout="wide")

st.title("🛡️ Sidharth Shutters & Automation Pvt. Ltd.")
st.subheader("Automated Quotation Builder")

# Master Dropdowns
SLAT_OPTIONS = [
    "90 mm (H) x 1 mm thick Galvalume Curved slat in natural finish",
    "78 mm (H) x 0.8 mm thick GI Slat in natural finish",
    "Aluminium Extruded Slat 75 mm"
]

PERFORATION_OPTIONS = [
    "Perforation- 7 Strip(Square)",
    "Perforation- 4 Strip(Square)",
    "None"
]

GUIDE_OPTIONS = [
    "150x150x2mmTG Guide with rubber seal with grey epoxy",
    "2mmTG Guide with rubber seal with grey epoxy",
    "100x100x2mm Guide"
]

ACTUATOR_OPTIONS = [
    "CE Certified indirect Drive - Brand Strong Life",
    "Direct Drive Heavy Duty Actuator"
]

# Client & Header Inputs
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

# Shutter Items Selection
st.header("🧱 Shutter Specifications")

if "shutter_items" not in st.session_state:
    st.session_state["shutter_items"] = [{
        "slat": SLAT_OPTIONS[0],
        "perforation": PERFORATION_OPTIONS[0],
        "guide": GUIDE_OPTIONS[0],
        "actuator": ACTUATOR_OPTIONS[0],
        "width": 5000,
        "height": 6000,
        "qty": 1,
        "unit_rate": 150000
    }]

def add_shutter():
    st.session_state["shutter_items"].append({
        "slat": SLAT_OPTIONS[0],
        "perforation": PERFORATION_OPTIONS[0],
        "guide": GUIDE_OPTIONS[0],
        "actuator": ACTUATOR_OPTIONS[0],
        "width": 5000,
        "height": 6000,
        "qty": 1,
        "unit_rate": 150000
    })

for idx, item in enumerate(st.session_state["shutter_items"]):
    with st.expander(f"Shutter #{idx + 1}", expanded=True):
        col1, col2, col3 = st.columns(3)
        item["slat"] = col1.selectbox("Slat Type", SLAT_OPTIONS, index=0, key=f"slat_{idx}")
        item["perforation"] = col2.selectbox("Perforation", PERFORATION_OPTIONS, index=0, key=f"perf_{idx}")
        item["guide"] = col3.selectbox("Guide Specification", GUIDE_OPTIONS, index=0, key=f"guide_{idx}")

        col4, col5 = st.columns(2)
        item["actuator"] = col4.selectbox("Actuator", ACTUATOR_OPTIONS, index=0, key=f"act_{idx}")
        
        c_w, c_h, c_q, c_r = st.columns(4)
        item["width"] = c_w.number_input("Width (mm)", value=int(item["width"]), step=100, key=f"w_{idx}")
        item["height"] = c_h.number_input("Height (mm)", value=int(item["height"]), step=100, key=f"h_{idx}")
        item["qty"] = c_q.number_input("Qty", value=int(item["qty"]), min_value=1, step=1, key=f"q_{idx}")
        item["unit_rate"] = c_r.number_input("Unit Rate (INR)", value=int(item["unit_rate"]), step=1000, key=f"r_{idx}")

st.button("➕ Add Another Shutter", on_click=add_shutter)

st.header("🚚 Packing & Freight")
col_p, col_f = st.columns(2)
packing_charges = col_p.number_input("Packing & Loading Charges (INR)", value=6093)
freight_charges = col_f.number_input("Freight Charges (INR)", value=50000)

# PDF Engine
def generate_pdf():
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('Title', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, alignment=1, textColor=colors.HexColor('#1A365D'))
    small_bold = ParagraphStyle('SmallBold', fontName='Helvetica-Bold', fontSize=8, leading=10)
    small_text = ParagraphStyle('SmallText', fontName='Helvetica', fontSize=8, leading=10)

    story.append(Paragraph("SIDHARTH SHUTTER & AUTOMATION PRIVATE LIMITED", title_style))
    story.append(Paragraph("G-1-66, Industrial Area, Prahaladpura, Sanganer, Jaipur, Rajasthan, 303903", ParagraphStyle('Sub', alignment=1, fontSize=8)))
    story.append(Spacer(1, 10))

    header_data = [
        [Paragraph("<b>Company Name</b>", small_bold), Paragraph(client_name, small_text), Paragraph("<b>Qtn Date:</b>", small_bold), Paragraph(str(qtn_date), small_text)],
        [Paragraph("<b>Contact Details</b>", small_bold), Paragraph(contact_details, small_text), Paragraph("<b>Sales Staff:</b>", small_bold), Paragraph(f"{sales_person_1}<br/>{sales_person_2}", small_text)],
        [Paragraph("<b>Shipping Address</b>", small_bold), Paragraph(shipping_address, small_text), Paragraph("<b>Qtn Ref No:</b>", small_bold), Paragraph(qtn_ref_no, small_bold)],
        [Paragraph("<b>Billing Address</b>", small_bold), Paragraph(billing_address, small_text), Paragraph("", small_text), Paragraph("", small_text)],
        [Paragraph("<b>GSTIN</b>", small_bold), Paragraph(gstin, small_text), Paragraph("", small_text), Paragraph("", small_text)]
    ]
    t_header = Table(header_data, colWidths=[90, 240, 70, 150])
    t_header.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(t_header)
    story.append(Spacer(1, 10))

    items_data = [[
        Paragraph("<b>Sr. No.</b>", small_bold),
        Paragraph("<b>Description</b>", small_bold),
        Paragraph("<b>HSN Code</b>", small_bold),
        Paragraph("<b>Width (mm)</b>", small_bold),
        Paragraph("<b>Height (mm)</b>", small_bold),
        Paragraph("<b>Qty</b>", small_bold),
        Paragraph("<b>Unit Rate (INR)</b>", small_bold),
        Paragraph("<b>Amount (INR)</b>", small_bold)
    ]]

    material_total = 0
    total_qty = 0

    for i, itm in enumerate(st.session_state["shutter_items"]):
        amt = itm["qty"] * itm["unit_rate"]
        material_total += amt
        total_qty += itm["qty"]

        desc = f"<b>Motorized Rolling Shutter</b><br/>" \
               f"- {itm['slat']}<br/>" \
               f"<b>{itm['perforation']}</b><br/>" \
               f"- {itm['guide']}<br/>" \
               f"- Super bottom in HR Sheet with rubber seal with grey epoxy<br/>" \
               f"- Strom Anchor | Safety Break | Wind Lock<br/>" \
               f"<b>Actuator for rolling shutter</b><br/>" \
               f"- {itm['actuator']}"

        items_data.append([
            Paragraph(str(i+1), small_text),
            Paragraph(desc, small_text),
            Paragraph("-", small_text),
            Paragraph(str(itm["width"]), small_text),
            Paragraph(str(itm["height"]), small_text),
            Paragraph(str(itm["qty"]), small_text),
            Paragraph(f"{itm['unit_rate']:,}", small_text),
            Paragraph(f"{amt:,}", small_text)
        ])

    subtotal_excl_gst = material_total + packing_charges + freight_charges
    gst_amount = round(subtotal_excl_gst * 0.18)
    grand_total = subtotal_excl_gst + gst_amount

    items_data.append(["", Paragraph("<b>Total Material Supply</b>", small_bold), "", "", "", Paragraph(f"<b>{total_qty}</b>", small_bold), "", Paragraph(f"<b>{material_total:,}</b>", small_bold)])
    items_data.append(["", Paragraph("Packing & Loading charges", small_text), "", "", "", "", "", Paragraph(f"{packing_charges:,}", small_text)])
    items_data.append(["", Paragraph("Freight charges (As per Actual)", small_text), "", "", "", "", "", Paragraph(f"{freight_charges:,}", small_text)])
    items_data.append(["", Paragraph("<b>Supply & Installation Amount Excluding GST</b>", small_bold), "", "", "", "", "", Paragraph(f"<b>{subtotal_excl_gst:,}</b>", small_bold)])
    items_data.append(["", Paragraph("GST on supply & installation @18%", small_text), "", "", "", "", "", Paragraph(f"{gst_amount:,}", small_text)])
    items_data.append(["", Paragraph("<b>Grand Total with GST</b>", small_bold), "", "", "", "", "", Paragraph(f"<b>{grand_total:,}</b>", small_bold)])

    t_items = Table(items_data, colWidths=[25, 230, 40, 45, 45, 25, 65, 75])
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
