import io
import json
import os
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pypdfium2 as pdfium  # PDF to Image Rendering
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import HRFlowable, Image as RLImage, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import streamlit as st

st.set_page_config(
    page_title="Sidharth Shutter & Automation Pvt. Ltd.",
    layout="wide",
    page_icon="🛡️",
)

# --- PRODUCTS LIST STRUCTURE ---
products = [
    {
        "productItem": "Rolling Shutters",
        "subCategory": "Motorized, Gear & Manual Rolling Shutters",
        "icon": "🌀",
    },
    {
        "productItem": "Dock Leveler",
        "subCategory": "Hydraulic & Mechanical Dock Levelers",
        "icon": "🏗️",
    },
    {
        "productItem": "Gates",
        "subCategory": "Sliding, Telescopic, L-Folding, Swing & Retractable Gates",
        "icon": "🚪",
    },
    {
        "productItem": "Doors",
        "subCategory": "High Speed, Fire, HMPS, GPD & Overhead Sectional Doors",
        "icon": "🚪",
    },
    {
        "productItem": "Boom Barrier",
        "subCategory": "Automatic & Heavy-Duty Traffic Barriers",
        "icon": "🚧",
    },
    {
        "productItem": "Dock Shelter",
        "subCategory": "Retractable & Inflatable Dock Shelters",
        "icon": "🏬",
    },
    {
        "productItem": "Dock Bumper",
        "subCategory": "Heavy Rubber & Moulded Bumpers",
        "icon": "🛡️",
    },
]

# --- CUSTOM CSS FOR ENHANCED ENTERPRISE LOOK ---
st.markdown(
    """
<style>
    .stApp {
        background-color: #f8fafc;
    }
    
    /* Welcome Header Banner */
    .welcome-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        padding: 45px 30px;
        border-radius: 18px;
        color: white;
        text-align: center;
        margin-bottom: 30px;
        box-shadow: 0 10px 20px -3px rgba(0, 0, 0, 0.15);
    }
    .welcome-header h1 {
        color: #ffffff !important;
        font-weight: 800;
        margin: 0;
        font-size: 36px;
        letter-spacing: 0.5px;
    }
    .welcome-header p {
        color: #94a3b8 !important;
        margin-top: 12px;
        margin-bottom: 0;
        font-size: 18px;
    }
    
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
    }
</style>
""",
    unsafe_allow_html=True,
)

# --- SESSION STATE FOR NAVIGATION ---
if "selected_product" not in st.session_state:
    st.session_state["selected_product"] = "Home"


def navigate_to(page_name):
    st.session_state["selected_product"] = page_name


# --- PERMANENT SPECIFICATIONS FILE MANAGEMENT ---
SPEC_FILE = "specifications_master.json"

DEFAULT_SPECS = {
    "slat_nat_list": [
        "78mm (H) x 1mm Thick Galvalume Curved Slat in Natural Finish",
        "78mm (H) x .80mm Thick Galvalume Curved Slat in Natural Finish",
        "78mm (H) x .90mm Thick Galvalume Curved Slat in Natural Finish",
        "78mm (H) x 1mm Thick Galvanized in Curved Slat",
        "78mm (H) x .80mm Thick Galvanized in Curved Slat",
        "78mm (H) x .90mm Thick Galvanized in Curved Slat",
        "90mm (H) x .80mm Thick Galvalume plain slats in Curved Slat",
        "90mm (H) x .90mm Thick Galvalume plain slats in Natural Finish",
        "90mm (H) x .80mm Thick Galvanized in plain slats",
        "90mm (H) x .90mm Thick Galvanized in plain slats"
    ],
    "slat_pow_list": [
        "Finish - Powder Coating As Per RAL",
        "Finish - Red Oxide",
        "Finish - PU Pain As Per RAL",
        "Finish - Enamel Paint As Per RAL"
    ],
    "guide_list": [
        "TG Guide with Rubber Seal",
        "TG Guide with Rubber Seal with Grey Epoxy Finish",
        "TG Guide With Grey Epoxy Finish",
        "U Guide with Grey Epoxy Finish",
        "U Guide"
    ],
    "bottom_list": [
        "Super Bottom with Rubber Seal with Grey Epoxy Finish",
        "Super Bottom with Rubber Seal",
        "Aluminium Bottom with Rubber Seal with Grey Epoxy Finish",
        "Aluminium Bottom with Rubber Seal"
    ],
    "hood_list": [
        ".80mm Thick Galvalume Hood & Motor Cover in Natural Finish",
        ".90mm Thick Galvalume Hood & Motor Cover in Natural Finish",
        "1.2mm Thick Galvalume Hood & Motor Cover in Natural Finish",
        "1mm Thick Galvalume Hood & Motor Cover in Natural Finish",
        ".80mm Thick Galvanized Hood & Motor Cover",
        ".90mm Thick Galvanized Hood & Motor Cover",
        "1mm Thick Galvanized Hood & Motor Cover",
        "1.2mm Thick Galvanized Hood & Motor Cover",
        ".80mm Thick Galvalume Hood Cover in Natural Finish",
        ".80mm Thick Galvanized Hood Cover"
    ],
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


specs = load_specifications()
for key, value in specs.items():
    if key not in st.session_state:
        st.session_state[key] = value

SAFETY_LOCK_OPTIONS = [
    "Wind Locks",
    "Storm Anchors",
    "External Safety Break",
    "Side Locks-2 Nos",
    "Standard Center Lock with 2 Keys",
    "Gear Set - Both Side",
    "Gear Set - Single Side (Outside)",
]

OPERATOR_OPTIONS = [
    "CE Certified Indirect Drive Brand Strong Life Sidharth Make",
    "Three Station Push Button",
]

# --- DYNAMIC MASTER PRODUCTS & SUB-CATEGORIES MAPPING ---
PRODUCT_HIERARCHY = {
    "Rolling Shutters": [
        "Motorized Rolling Shutter",
        "Gear Rolling Shutter",
        "Manual Rolling Shutter",
    ],
    "Dock Leveler": [
        "Hydraulic Dock Leveler",
        "Mechanical Dock Leveler",
    ],
    "Gates": [
        "Sliding Gate",
        "Telescopic Gate",
        "L-Folding Gate",
        "Swing Gate",
        "Retractable Gate",
    ],
    "Doors": [
        "High Speed Door",
        "Fire Door",
        "HMPS Door",
        "GPD Door",
        "Overhead Sectional Door",
    ],
    "Boom Barrier": [
        "Automatic Traffic Barrier",
        "Heavy-Duty Traffic Barrier",
    ],
    "Dock Shelter": [
        "Retractable Dock Shelter",
        "Inflatable Dock Shelter",
    ],
    "Dock Bumper": [
        "Heavy Rubber Bumper",
        "Moulded Bumper",
    ]
}

# DEFAULT DATA FOR SECTION A (PAGE 2)
DEFAULT_ABOUT_US_LIST = [
    {
        "title": "1. Core Company Highlights & Industrial Legacy",
        "text": "• Over Two Decades of Engineering Excellence: Established in 1998, Sidharth Shutter & Automation Pvt. Ltd. (SSAPL) has evolved into a premier national leader in high-performance industrial entrance and access automation systems across India.\n• Extensive Nationwide Deployment: Successfully engineered, custom-manufactured, and commissioned over 10,000+ heavy-duty industrial shutters, dock systems, and automated security barriers for leading public and private enterprises.\n• Pan-India Operational Footprint: Serving top-tier logistics parks, manufacturing plants, warehouse hubs, defense facilities, commercial complexes, and infrastructure projects with robust regional project execution capability.\n• Advanced Manufacturing Facility: Powered by state-of-the-art infrastructure featuring CNC roll-forming machinery, high-precision laser cutters, robotic welding units, and fully automated powder-coating lines."
    },
    {
        "title": "2. Engineering, Quality Standards & Certification",
        "text": "• ISO & CE Certified Quality Benchmarks: Strict compliance with ISO 9001:2015 Quality Management Systems and European CE safety certification standards across all product development cycles.\n• High Wind-Load & Extreme Weather Resilience: Structural curtains and frames are specifically calculated and stress-tested to withstand high wind velocity, pressure differentials, severe rain, and harsh industrial environments.\n• Premium Heavy-Gauge Raw Materials: Exclusively using certified high-tensile Galvalume, Galvanized Steel (GI), and structural aluminum alloys for exceptional structural stability, long operational lifespan, and maximum corrosion protection."
    },
    {
        "title": "3. Comprehensive Product & Solutions Portfolio",
        "text": "• Industrial & Commercial Rolling Shutters: Heavy-duty motorized, gear-operated, and manual roll-up systems designed for maximum perimeter security, thermal insulation, and continuous high-frequency operations.\n• High-Speed & Specialized Doors: Rapid roll-up PVC doors, insulated sectional overhead doors, fire-rated HMPS doors, and acoustic/cleanroom barriers engineered for dynamic airflow control.\n• Advanced Entrance Automation: Automatic boom barriers, cantilever sliding gates, telescopic gates, heavy-duty swing gates, and motor-driven retractable security gates.\n• Integrated Loading Bay Equipment: Hydraulic and mechanical dock levelers, inflatable/retractable dock shelters, heavy-duty rubber bumpers, and integrated loading dock safety accessories."
    },
    {
        "title": "4. The Sidharth Advantage: Turnkey Execution & Support",
        "text": "• End-to-End Turnkey Project Management: Complete project lifecycle execution—from initial site survey, structural shop drawings, and custom fabrication to installation, electrical integration, and final handover.\n• 24x7 Rapid Service & Maintenance Network: Highly specialized field service engineering team delivering preventative maintenance, fast emergency troubleshooting, AMC services, and guaranteed original spare parts supply.\n• Architectural & Operational Customization: Tailor-made designs engineered to integrate smoothly into non-standard structural openings, unique headroom spaces, and specialized warehouse management workflows."
    },
    {
        "title": "5. Technical R&D, Testing & Safety Protocols",
        "text": "• In-House Research & Innovation: Continuous R&D focused on ultra-quiet drive mechanisms, high-wind slat locking mechanisms, smart sensor technology, and long-life motor controls.\n• Rigorous Multi-Cycle Factory Testing: Every motor drive, control unit, and shutter curtain undergoes strict pre-dispatch factory testing for smooth mechanical movement, limit accuracy, and electrical load capacity.\n• Uncompromising Operational Safety: Integrated emergency manual hand-chain system, optical safety sensors, bottom-edge safety buffers, and external dynamic drop-brake protection standard across heavy-duty configurations."
    }
]

if "about_us_data" not in st.session_state:
    st.session_state["about_us_data"] = DEFAULT_ABOUT_US_LIST.copy()

# EXTENDED TECH SPECS FOR SECTION C (PAGE 4) - UPDATED EXACTLY TO 30 DETAILS
DEFAULT_TECH_SPECS = [
    {"param": "Shutter Type", "spec": "Motorized Heavy Duty Industrial Rolling Shutter"},
    {"param": "Slat Type", "spec": "Single Skin Curved Interlocking Galvanized / Galvalume Steel Slats"},
    {"param": "Slat Material", "spec": "Galvanized Steel (GI) / Galvalume Alloy High Tensile Grade"},
    {"param": "Slat Thickness", "spec": "0.90 mm to 1.20 mm Heavy Gauge Option"},
    {"param": "Slat Height", "spec": "78 mm / 90 mm profile pitch with anti-vibration design"},
    {"param": "Surface Finish", "spec": "Pure Polyester Outdoor Powder Coating in Standard RAL Shade (60-80 microns)"},
    {"param": "Curtain Construction", "spec": "Interlocking curved profile continuous slats with heavy nylon end locks to prevent lateral movement"},
    {"param": "Support Brackets", "spec": "Laser-cut fabricated from 5 mm to 8 mm Thick HR Steel Plates with structural gussets"},
    {"param": "Shaft Assembly", "spec": "Heavy Duty Seamless MS Pipe Shaft designed for zero deflection, with self-aligning sealed bearings"},
    {"param": "Side Guides", "spec": "Heavy Duty TG Guide fabricated from 2.0 mm / 2.5 mm Thick Galvanized Steel with continuous EPDM rubber seals"},
    {"param": "Bottom Profile", "spec": "Heavy Duty 2.5 mm Thick HR Steel Bottom Angle/Rail with weather-resistant EPDM bottom seal"},
    {"param": "Drive Motor Unit", "spec": "Side Mounted Heavy Duty Motor, Brand: Strong Life / Sidharth Make, CE Certified, IP54 rated"},
    {"param": "Operation Control", "spec": "Heavy Duty 3-Station Push Button Station (Open/Close/Stop) with Key Lock Box option"},
    {"param": "Limit Switch System", "spec": "Dual Electro-Mechanical Micro Limit Switches for precise high-accuracy height limit regulation"},
    {"param": "Safety Emergency System", "spec": "Manual Hand Chain / Crank Override System for operation during power failure"},
    {"param": "Wind Resistance", "spec": "Engineered to withstand Class 3/4 wind pressure (Up to 120 km/h) with wind lock attachments"},
    {"param": "Hood Cover Unit", "spec": "0.80 mm Thick Galvanized Steel Protective Hood & Motor Enclosure with rigid support framing"},
    {"param": "Control Panel & Supply", "spec": "IP65 Rated Control Box with integrated Thermal Overload Protection & Phase Reversing Relay"},
    {"param": "Locking Arrangement", "spec": "Electrical Self-Locking Gear Mechanism (Optional Mechanical Side Locks for dual security)"},
    {"param": "Fasteners & Hardware", "spec": "High Tensile Zinc-Plated / Hot-Dip Galvanized Fasteners & Anchor Expansion Bolts"},
    {"param": "Operating Temp & Cycle", "spec": "-10°C to +55°C Ambient Operating Range; Designed for Heavy Continuous Industrial Cycles"},
    {"param": "Opening / Closing Speed", "spec": "0.15 m/sec to 0.25 m/sec (Standard Industrial Speed)"},
    {"param": "Power Supply Requirement", "spec": "415V AC, 3-Phase, 50 Hz / 230V Single Phase (As per motor model)"},
    {"param": "Duty Cycle Rating", "spec": "60% Duty Cycle (S3 Rated Heavy Duty Industrial Motor)"},
    {"param": "Noise Level", "spec": "Low noise operation (< 65 dB at 1 meter distance)"},
    {"param": "Ingress Protection (IP Rating)", "spec": "IP54 for Motor Unit & IP65 for Control Panel Box"},
    {"param": "Safety Edge Sensor (Optional)", "spec": "Infrared Safety Photocells / Bottom Rubber Wireless Safety Edge Sensor"},
    {"param": "Manual Override Force", "spec": "Ergonomic Chain Handwheel requiring < 15 kg pull force during power loss"},
    {"param": "Guide Weather Sealing", "spec": "Dual-density EPDM / Nylon Brush seals inside side guides for dust & weather insulation"},
    {"param": "Bottom Seal Type", "spec": "Heavy-duty EPDM tubular rubber bottom profile for floor gap sealing"}
]

if "tech_specs_data" not in st.session_state:
    st.session_state["tech_specs_data"] = DEFAULT_TECH_SPECS.copy()

# EXTENDED TERMS & CONDITIONS FOR SECTION D (PAGE 5)
DEFAULT_TERMS = [
    {
        "category": "Unloading & Handling",
        "details": "• Unloading of material at site is strictly in client scope.\n• All required civil work, pocket cutting, and masonry work is in client scope.\n• Scaffolding, staging, and height access arrangements are in client scope.\n• Hydra crane, forklift, and heavy material handling equipment to be arranged by client."
    },
    {
        "category": "Storage & Material Handling",
        "details": "• Safe, dry, and locked storage space for materials at site until installation completion.\n• Material verification against packing list and shifting to exact location is client's responsibility."
    },
    {
        "category": "Site Security & Deterioration",
        "details": "• Any shortage or damage due to theft, pilferage, or misplacement at site is borne by client.\n• Refurbishment/painting necessitated by prolonged site storage in open environment is in client scope."
    },
    {
        "category": "Electrical Scope",
        "details": "• Single/Three-phase AC power supply with dedicated MCB up to installation point by client.\n• Standard 16A socket and free electric power point within 10 meters of installation spot.\n• Supply and laying of main power cables, conduits, and cable trays in client scope.\n• If permanent power is unavailable, testing will be done via client's temporary supply for handover."
    },
    {
        "category": "Power Quality & Stability",
        "details": "• Client shall ensure stabilized, uninterrupted power supply with standard voltage tolerances."
    },
    {
        "category": "Site Visits & Travel",
        "details": "• Installation scope covers maximum 2 site visits by our erection technical team.\n• Additional visits required due to site unreadiness incur ₹2,000/- per person per visit plus travel expenses."
    },
    {
        "category": "Warranty Terms",
        "details": "• Automation system is warranted against manufacturing defects for 12 months from installation or 13 months from invoice date, whichever is earlier."
    },
    {
        "category": "Payment & Price Validity",
        "details": "• 75% Advance along with formal Purchase Order.\n• 25% balance payment prior to material dispatch against proforma invoice.\n• This commercial offer remains valid for 20 days from the date of issuance."
    },
    {
        "category": "Delivery & Freight Scope",
        "details": "• Quoted freight charges apply for single full shipment. Partial shipments incur extra charges.\n• Delivery timeline: 2 weeks from drawing approval or advance receipt, whichever is later."
    },
    {
        "category": "Force Majeure & Legal",
        "details": "• Delays due to natural disasters, strikes, wars, or government restrictions are beyond our liability.\n• Any legal disputes arising out of this contract shall be subject to Jaipur Jurisdiction only."
    }
]

if "terms_data" not in st.session_state:
    st.session_state["terms_data"] = DEFAULT_TERMS.copy()

# ==============================================================================
# PAGE 1: HOME / LANDING PAGE (WELCOME NOTE ONLY)
# ==============================================================================
if st.session_state["selected_product"] == "Home":

    st.markdown(
        """
    <div class="welcome-header">
        <h1>Welcome To Sidharth Shutter & Automation Pvt. Ltd.</h1>
        <p>Industrial Access & Entrance Automation Solutions</p>
    </div>
    """,
        unsafe_allow_html=True,
    )

    st.write("<br>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown(
            """
            <div style="background-color: white; padding: 30px; border-radius: 12px; border: 1px solid #e2e8f0; text-align: center; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);">
                <h3 style="color: #0f172a; margin-bottom: 10px;">Create Professional Quotation</h3>
                <p style="color: #64748b; font-size: 14px; margin-bottom: 25px;">
                    Click below to start generating custom multi-page commercial offers and technical proposals.
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )
        st.write("")
        if st.button("🚀 Create Quotation", use_container_width=True, type="primary"):
            st.session_state["selected_product"] = "Rolling Shutters"
            st.rerun()

# ==============================================================================
# PAGE 2: ROLLING SHUTTERS / MULTI-PRODUCT QUOTATION PAGE
# ==============================================================================
elif st.session_state["selected_product"] == "Rolling Shutters":

    col_nav1, col_nav2 = st.columns([1, 6])
    with col_nav1:
        if st.button("🏠 Home", use_container_width=True):
            st.session_state["selected_product"] = "Home"
            st.rerun()
    with col_nav2:
        st.title("📋 Quotation Builder")

    st.sidebar.header("📋 Client & Proposal Details")

    client_id = st.sidebar.text_input("Client ID", "SSAPL-CL-8842")
    client_name = st.sidebar.text_input(
        "Company Name", "M/S ABHINAV INFRA BUILD PVT. LTD."
    )
    attn_person = st.sidebar.text_input("Attention To / Kind Attn", "Mr. Sharma")
    contact_details = st.sidebar.text_input("Contact Details", "Mob:-8889911529")
    shipping_address = st.sidebar.text_area(
        "Shipping Address",
        "RGLP PITHAMPUR Sector-1 Pithampur MADHYA PRADESH 454774",
        height=70,
    )
    site_person = st.sidebar.text_input("Site Person Mobile", "9926952398")
    billing_address = st.sidebar.text_area(
        "Billing Address",
        "207,208, Industry House, Indore, MP 452001",
        height=70,
    )
    gstin = st.sidebar.text_input("Client GSTIN", "23AAHCA9425D1ZY")

    st.sidebar.markdown("---")
    c_date, c_ref = st.sidebar.columns(2)
    qtn_date = c_date.date_input("Quotation Date")
    qtn_ref_no = c_ref.text_input("Qtn Ref No", "SSAPL/2026-27/319R2")

    sales_reference_1 = st.sidebar.text_input(
        "Sales Reference 1", "Mr. Nishant (90010 42914)"
    )
    sales_reference_2 = st.sidebar.text_input(
        "Sales Reference 2", "Mr. Jeevan Sharma (9828771899)"
    )

    quotation_made_by = st.sidebar.text_input(
        "Quotation Made By", "Lokesh MIS"
    )

    # --- DYNAMIC PAGE SELECTION TOGGLES IN SIDEBAR ---
    st.sidebar.markdown("---")
    st.sidebar.header("📄 Page Selection Settings")
    st.sidebar.markdown("Choose pages to include in the quotation PDF:")
    include_page_2 = st.sidebar.checkbox("Include Page 2 (About Us)", value=True)
    include_page_4 = st.sidebar.checkbox("Include Page 4 (Technical Specs)", value=True)
    include_page_5 = st.sidebar.checkbox("Include Page 5 (Terms & Conditions)", value=True)

    # --- COVER LETTER INTRO BODY IN EXPANDER ---
    with st.expander("📄 Page 1: Cover Letter Text Customization", expanded=False):
        cover_body_text = st.text_area(
            "Cover Letter Intro Body",
            "Dear Sir / Ma'am,\n\n"
            "We extend our sincere gratitude for the interest you have shown in our products and services. "
            "It is our privilege to present this detailed commercial offer, crafted to precisely address "
            "the access control and security requirements of your facility.\n\n"
            "At Sidharth Shutter & Automation Pvt. Ltd., we understand that a well-secured premise is the "
            "foundation of efficient operations. Our solutions are engineered not only to keep your unit secure "
            "but also to ensure smooth, reliable, and operationally efficient entry and exit for your facility, day in and day out.\n\n"
            "This proposal package consists of the detailed specifications and commercial offer enclosed herewith.\n\n"
            "We are confident that our proposed solution will perfectly align with your requirements. "
            "For any further clarifications, please feel free to reach out to us.\n\n"
            "Assuring you of our best services at all times.\n\n"
            "Warm Regards,",
            height=260
        )

    # --- SECTION A: ABOUT US TEXT CUSTOMIZATION IN EXPANDER (WITH ADD/REMOVE) ---
    with st.expander("📄 Page 2: About Us (Company Profile) Text Customization", expanded=False):
        st.markdown("Edit, Add or Remove items for Section A (About Us):")
        
        updated_about = []
        for idx, item in enumerate(st.session_state["about_us_data"]):
            col_t, col_d, col_rm = st.columns([2, 3, 1])
            with col_t:
                t_val = st.text_input(f"Section #{idx+1} Title", value=item["title"], key=f"ab_title_{idx}")
            with col_d:
                x_val = st.text_area(f"Section #{idx+1} Details", value=item["text"], height=100, key=f"ab_text_{idx}")
            with col_rm:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("❌ Remove", key=f"rm_ab_{idx}"):
                    st.session_state["about_us_data"].pop(idx)
                    st.rerun()
            updated_about.append({"title": t_val, "text": x_val})
        
        st.session_state["about_us_data"] = updated_about
        
        if st.button("➕ Add Section A Block"):
            st.session_state["about_us_data"].append({
                "title": f"{len(st.session_state['about_us_data'])+1}. New Section Title",
                "text": "• Add your detail point here..."
            })
            st.rerun()

    # --- SECTION C: TECHNICAL SPECIFICATIONS CUSTOMIZATION IN EXPANDER (WITH ADD/REMOVE) ---
    with st.expander("📄 Page 4: Section C - Technical Specifications Customization", expanded=False):
        st.markdown("Edit, Add or Remove parameters and specifications for Page 4:")
        tech_intro_text = st.text_area(
            "Header Introductory Text",
            "The following specifications define the precise engineering parameters for the High-Speed Door / Rolling Shutter system offered for your facility. Every component is designed for operational efficiency, durability, and safety compliance.",
            height=70
        )
        
        updated_specs = []
        for idx, item in enumerate(st.session_state["tech_specs_data"]):
            col_p, col_s, col_rm = st.columns([2, 3, 1])
            with col_p:
                p_val = st.text_input(f"Parameter #{idx+1}", value=item["param"], key=f"tp_{idx}")
            with col_s:
                s_val = st.text_input(f"Specification #{idx+1}", value=item["spec"], key=f"ts_{idx}")
            with col_rm:
                if st.button("❌ Remove", key=f"rm_ts_{idx}"):
                    st.session_state["tech_specs_data"].pop(idx)
                    st.rerun()
            updated_specs.append({"param": p_val, "spec": s_val})
        
        st.session_state["tech_specs_data"] = updated_specs

        if st.button("➕ Add Technical Specification"):
            st.session_state["tech_specs_data"].append({
                "param": "New Parameter",
                "spec": "Specification Description Details"
            })
            st.rerun()
        
        st.markdown("---")
        tech_note_text = st.text_input(
            "Footer Note Text",
            "Note: Technical specifications are subject to final site measurements and drawing approval. Custom configurations can be engineered to suit the precise dimensions of your facility opening."
        )

    # --- SECTION D: TERMS & CONDITIONS CUSTOMIZATION IN EXPANDER (WITH ADD/REMOVE) ---
    with st.expander("📄 Page 5: Section D - Terms & Conditions Customization", expanded=False):
        st.markdown("Edit, Add or Remove entries for Section D (Terms & Conditions):")
        exclusions_subhead = st.text_input("Sub-Header Title", "Exclusions - Client Scope & Operational Terms")
        
        updated_terms = []
        for idx, term in enumerate(st.session_state["terms_data"]):
            col_tc, col_td, col_rm = st.columns([2, 3, 1])
            with col_tc:
                c_val = st.text_input(f"Category #{idx+1}", value=term["category"], key=f"tc_cat_{idx}")
            with col_td:
                d_val = st.text_area(f"Details #{idx+1}", value=term["details"], height=70, key=f"tc_det_{idx}")
            with col_rm:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("❌ Remove", key=f"rm_tc_{idx}"):
                    st.session_state["terms_data"].pop(idx)
                    st.rerun()
            updated_terms.append({"category": c_val, "details": d_val})
        
        st.session_state["terms_data"] = updated_terms

        if st.button("➕ Add Term / Exclusion"):
            st.session_state["terms_data"].append({
                "category": "New Scope / Category",
                "details": "• Enter scope details here..."
            })
            st.rerun()

    # --- REVISED HEADINGS & DYNAMIC PRODUCT MAPPING ---
    st.markdown("### 📦 Product Specifications & Pricing")

    if "shutter_items" not in st.session_state:
        st.session_state["shutter_items"] = [{
            "main_product": "Rolling Shutters",
            "type": "Motorized Rolling Shutter",
            "slat_nat": [],
            "slat_pow": [],
            "guide": [],
            "bottom": [],
            "hood": [],
            "safety_locks": [
                "Wind Locks",
                "Storm Anchors",
                "External Safety Break",
            ],
            "operator": [
                "CE Certified Indirect Drive Brand Strong Life Sidharth Make"
            ],
            "hsn": "73083000",
            "width": 4650,
            "height": 7000,
            "qty": 5,
            "mat_rate": 176700,
            "inst_rate": 8000,
        }]

    def add_shutter():
        st.session_state["shutter_items"].append({
            "main_product": "Rolling Shutters",
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
            "inst_rate": 6000,
        })

    def remove_shutter(index):
        if len(st.session_state["shutter_items"]) > 1:
            st.session_state["shutter_items"].pop(index)

    for idx, item in enumerate(st.session_state["shutter_items"]):
        selected_sub_cat = item.get("type", "Motorized Rolling Shutter")
        with st.expander(
            f"📌 Item #{idx + 1}: {selected_sub_cat}",
            expanded=True,
        ):
            col_p_main, col_p_cat, col_del = st.columns([2, 2, 1])
            
            # 1. Primary Product Type Selection
            with col_p_main:
                curr_main_prod = item.get("main_product", "Rolling Shutters")
                main_prod_keys = list(PRODUCT_HIERARCHY.keys())
                main_idx = main_prod_keys.index(curr_main_prod) if curr_main_prod in main_prod_keys else 0
                
                selected_main = st.selectbox(
                    f"Product Item #{idx + 1}",
                    main_prod_keys,
                    index=main_idx,
                    key=f"main_prod_{idx}"
                )
                item["main_product"] = selected_main

            # 2. Dynamic Sub-Category Selection based on Product
            with col_p_cat:
                available_sub_cats = PRODUCT_HIERARCHY.get(selected_main, ["Standard Option"])
                curr_sub = item.get("type", available_sub_cats[0])
                sub_idx = available_sub_cats.index(curr_sub) if curr_sub in available_sub_cats else 0
                
                item["type"] = st.selectbox(
                    f"Category / Sub-Category #{idx + 1}",
                    available_sub_cats,
                    index=sub_idx,
                    key=f"sub_cat_{idx}"
                )

            # 3. Item Delete Button
            with col_del:
                st.markdown("<br>", unsafe_allow_html=True)
                if len(st.session_state["shutter_items"]) > 1:
                    st.button(
                        "❌ Remove",
                        key=f"del_{idx}",
                        on_click=remove_shutter,
                        args=(idx,),
                    )

            item["hsn"] = st.text_input(
                "HSN Code", value=item.get("hsn", "73083000"), key=f"hsn_{idx}"
            )

            # Dynamic Form Fields based on Main Product Type
            if selected_main == "Rolling Shutters":
                st.markdown("##### 📐 Technical Details")
                col_sp, col_sn = st.columns(2)
                # Reordered Fields: 1st Paint Finish, 2nd Slats
                with col_sp:
                    item["slat_pow"] = st.multiselect(
                        "Paint Finish",
                        st.session_state["slat_pow_list"],
                        default=item.get("slat_pow", []),
                        key=f"sp_{idx}",
                    )

                with col_sn:
                    item["slat_nat"] = st.multiselect(
                        "Slats",
                        st.session_state["slat_nat_list"],
                        default=item.get("slat_nat", []),
                        key=f"sn_{idx}",
                    )

                st.markdown("##### 🛠️ Guide, Bottom Sheet & Hood Cover")
                cg, cb, ch_col = st.columns(3)
                with cg:
                    item["guide"] = st.multiselect(
                        "Guide Specification",
                        st.session_state["guide_list"],
                        default=item.get("guide", []),
                        key=f"gd_{idx}",
                    )

                with cb:
                    item["bottom"] = st.multiselect(
                        "Bottom Specification",
                        st.session_state["bottom_list"],
                        default=item.get("bottom", []),
                        key=f"bt_{idx}",
                    )

                with ch_col:
                    item["hood"] = st.multiselect(
                        "Hood Cover Specification",
                        st.session_state["hood_list"],
                        default=item.get("hood", []),
                        key=f"hd_{idx}",
                    )

                st.markdown("##### 🔒 Locks & Safety Features")
                item["safety_locks"] = st.multiselect(
                    "Select Locks & Safety Features",
                    SAFETY_LOCK_OPTIONS,
                    default=item.get("safety_locks", []),
                    key=f"lock_{idx}",
                )

                st.markdown("##### ⚙️ Operator Details")
                if item["type"] == "Motorized Rolling Shutter":
                    item["operator"] = st.multiselect(
                        "Operator Option",
                        OPERATOR_OPTIONS,
                        default=item.get("operator", []),
                        key=f"op_{idx}",
                    )
                else:
                    st.info(
                        "🔒 Operator selection disabled for Gear and Manual Shutters."
                    )
                    item["operator"] = []
            else:
                st.info(f"ℹ️ Custom specification fields for **{selected_main}** will load here.")

            st.markdown("---")
            st.markdown("**Dimensions & Rates:**")
            cw, ch, cq, cmr, cir = st.columns(5)
            item["width"] = cw.number_input(
                "Width (mm)",
                value=int(item.get("width", 4000)),
                step=50,
                key=f"w_{idx}",
            )
            item["height"] = ch.number_input(
                "Height (mm)",
                value=int(item.get("height", 5000)),
                step=50,
                key=f"h_{idx}",
            )
            item["qty"] = cq.number_input(
                "Qty",
                value=int(item.get("qty", 1)),
                min_value=1,
                step=1,
                key=f"q_{idx}",
            )
            item["mat_rate"] = cmr.number_input(
                "Material Rate (INR)",
                value=int(item.get("mat_rate", 50000)),
                step=500,
                key=f"mr_{idx}",
            )
            item["inst_rate"] = cir.number_input(
                "Installation Rate (INR)",
                value=int(item.get("inst_rate", 5000)),
                step=100,
                key=f"ir_{idx}",
            )

    st.button("➕ Add Another Product Item", on_click=add_shutter)

    # --- FREIGHT CHARGES & EXTRA CHARGES DISPLAY ---
    st.markdown("### 🚚 Extra Charges & Expenses")
    c1, c2, c3, c4, c5 = st.columns(5)
    packing_charges = c1.number_input("Packing & Loading (INR)", value=5000)
    
    # Toggle to show numeric value or custom text for Freight
    show_custom_freight = c2.checkbox("Hide Amount & Show Text", key="custom_freight_toggle")
    if show_custom_freight:
        freight_text_display = c2.text_input("Freight Display Text", value="Extra Charges", key="freight_text_input")
        freight_charges = 0
    else:
        freight_charges = c2.number_input("Freight Charges (INR)", value=15000)
        freight_text_display = f"{freight_charges:,}"

    unloading_charges = c3.number_input("Unloading Charges (INR)", value=0)
    crane_charges = c4.number_input("Crane Charges (INR)", value=0)
    scaffolding_charges = c5.number_input("Scaffolding Charges (INR)", value=0)

    # --- PDF GENERATOR ---
    def generate_pdf():
        buffer = io.BytesIO()

        doc = SimpleDocTemplate(
            buffer,
            pagesize=A4,
            leftMargin=25,
            rightMargin=25,
            topMargin=20,
            bottomMargin=20,
        )
        story = []
        styles = getSampleStyleSheet()

        # --- UPDATED HEADER HELPER FUNCTION ---
        def get_header_element():
            logo_path = None
            for possible_path in ["Logo.jpeg", "logo.jpeg", "Logo.png", "logo.png", "Logo.jpg", "logo.jpg"]:
                if os.path.exists(possible_path):
                    logo_path = possible_path
                    break
                    
            if logo_path:
                try:
                    logo_element = RLImage(logo_path, width=180, height=60)
                except Exception:
                    logo_element = Paragraph("<b>SIDHARTH</b><br/><font size=8 color='#003366'>SHUTTER & AUTOMATION</font>", styles["Normal"])
            else:
                logo_element = Paragraph("<b>SIDHARTH</b><br/><font size=8 color='#003366'>SHUTTER & AUTOMATION</font>", styles["Normal"])

            right_bold = ParagraphStyle(
                "HeadRightBold",
                parent=styles["Normal"],
                fontName="Helvetica-Bold",
                fontSize=9.5,
                leading=12,
                alignment=0,
                textColor=colors.HexColor("#0F172A"),
            )
            
            right_text = ParagraphStyle(
                "HeadRightText",
                parent=styles["Normal"],
                fontName="Helvetica",
                fontSize=8.5,
                leading=11,
                alignment=0,
                textColor=colors.HexColor("#1D4ED8"),
            )

            right_table_data = [
                [Paragraph("<b>GSTIN/UIN: 08AEEPJ6848R1ZN</b>", right_bold)],
                [Paragraph("<b>Web:</b> <font color='#0284C7'>www.ssaapl.com</font>", right_text)],
                [Paragraph("<b>Email:</b> <font color='#0284C7'>sales@ssaapl.com</font>", right_text)],
                [Paragraph("<b>Ph:</b> <font color='#0284C7'>+91 90019 96526, +91 90010 42908</font>", right_text)],
                [Paragraph("<b>Add:</b> <font color='#0284C7'>H-1-89, RIICO Ind. Area, Mansarovar, Jaipur, Rajasthan, 302020</font>", right_text)],
            ]

            t_right_info = Table(right_table_data, colWidths=[240])
            t_right_info.setStyle(
                TableStyle([
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 1),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                ])
            )

            t_head = Table([[logo_element, t_right_info]], colWidths=[295, 250])
            t_head.setStyle(
                TableStyle([
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (0, 0), (0, 0), "LEFT"),
                    ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ])
            )
            return t_head

        # --- UNIFORM & ENHANCED STYLES FOR PDF PAGES ---
        style_cover_meta = ParagraphStyle(
            "CoverMeta",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=10,
            leading=15,
            textColor=colors.HexColor("#1E293B"),
        )
        style_cover_body = ParagraphStyle(
            "CoverBody",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=10,
            leading=15.5,
            textColor=colors.HexColor("#334155"),
        )
        style_sec_title = ParagraphStyle(
            "SecTitle",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=12,
            leading=16,
            textColor=colors.HexColor("#000000"),
        )

        # ==========================================
        # 📄 PAGE 1: COVER LETTER / PROPOSAL LETTER
        # ==========================================
        story.append(get_header_element())
        story.append(Spacer(1, 10))
        story.append(
            HRFlowable(
                width="100%",
                thickness=1.5,
                color=colors.HexColor("#000000"),
                spaceAfter=15,
            )
        )

        meta_text = (
            f"<b>Ref No:</b> {qtn_ref_no}"
            " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
            f" <b>Date:</b> {qtn_date}"
        )
        story.append(Paragraph(meta_text, style_cover_meta))
        story.append(Spacer(1, 15))

        ship_formatted = shipping_address.replace("\n", "<br/>")
        client_info_p1 = (
            f"<b>To,</b><br/>"
            f"<b>{client_name}</b><br/>"
            f"{ship_formatted}<br/>"
            f"<b>Kind Attn:</b> {attn_person}<br/>"
            f"<b>Contact:</b> {contact_details}"
        )
        story.append(Paragraph(client_info_p1, style_cover_meta))
        story.append(Spacer(1, 20))

        body_formatted = cover_body_text.replace("\n", "<br/>")
        story.append(Paragraph(body_formatted, style_cover_body))
        story.append(Spacer(1, 35))

        sign_off_p1 = (
            "<b>For SIDHARTH SHUTTER & AUTOMATION PVT. LTD.</b><br/><br/><br/>"
            f"<b>Authorized Signatory</b><br/>"
            f"<b>({quotation_made_by})</b>"
        )
        story.append(Paragraph(sign_off_p1, style_cover_meta))

        # ==========================================
        # 🏢 PAGE 2: SECTION A: ABOUT US (CONDITIONAL)
        # ==========================================
        if include_page_2:
            story.append(PageBreak())
            story.append(get_header_element())
            story.append(Spacer(1, 4))
            story.append(
                HRFlowable(
                    width="100%",
                    thickness=1.5,
                    color=colors.HexColor("#000000"),
                    spaceAfter=6,
                )
            )

            story.append(Paragraph("Section A: About Us & Our Expertise (Company Profile)", style_sec_title))
            story.append(Spacer(1, 6))

            style_box_head = ParagraphStyle(
                "BoxHeadFull",
                parent=styles["Normal"],
                fontName="Helvetica-Bold",
                fontSize=10,
                leading=13,
                textColor=colors.HexColor("#000000"),
            )
            style_box_text = ParagraphStyle(
                "BoxTextFull",
                parent=styles["Normal"],
                fontName="Helvetica",
                fontSize=9.5,
                leading=13,
                textColor=colors.HexColor("#1E293B"),
            )

            def build_paragraph_block(text_data):
                elements = []
                for line in text_data.split("\n"):
                    if line.strip():
                        formatted_line = line.replace("•", "&bull;")
                        elements.append(Paragraph(formatted_line, style_box_text))
                        elements.append(Spacer(1, 2))
                return elements

            about_box_content = []
            for ab_item in st.session_state["about_us_data"]:
                about_box_content.append(Paragraph(f"<b>{ab_item['title']}</b>", style_box_head))
                about_box_content.append(Spacer(1, 2))
                about_box_content.extend(build_paragraph_block(ab_item['text']))
                about_box_content.append(Spacer(1, 4))

            t_box = Table([[about_box_content]], colWidths=[545])
            t_box.setStyle(
                TableStyle([
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
                    ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#000000")),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ])
            )
            story.append(t_box)
            story.append(Spacer(1, 6))

            style_pf_header = ParagraphStyle(
                "PfHeadFull",
                parent=styles["Normal"],
                fontName="Helvetica-Bold",
                fontSize=10,
                leading=12,
                alignment=1,
                textColor=colors.HexColor("#000000"),
            )
            style_pf_col = ParagraphStyle(
                "PfColFull",
                parent=styles["Normal"],
                fontName="Helvetica-Bold",
                fontSize=9,
                leading=11.5,
                alignment=1,
                textColor=colors.HexColor("#1E293B"),
            )
            style_pf_body = ParagraphStyle(
                "PfBodyFull",
                parent=styles["Normal"],
                fontName="Helvetica",
                fontSize=8.5,
                leading=11,
                alignment=1,
                textColor=colors.HexColor("#334155"),
            )

            portfolio_data = [
                [Paragraph("<b>Our Key Enterprise Pillars & Service Footprint</b>", style_pf_header), "", "", ""],
                [
                    Paragraph("<b>Rolling Shutters</b>", style_pf_col),
                    Paragraph("<b>Loading Equipment</b>", style_pf_col),
                    Paragraph("<b>Automation Systems</b>", style_pf_col),
                    Paragraph("<b>Pan-India AMC Support</b>", style_pf_col),
                ],
                [
                    Paragraph("Motorized & gear shutters for industrial, warehouse & commercial openings.", style_pf_body),
                    Paragraph("Dock levelers, shelters & bumpers for integrated logistics bays.", style_pf_body),
                    Paragraph("Boom barriers, sliding gates & high-speed doors for controlled access.", style_pf_body),
                    Paragraph("Dedicated regional engineer teams for fast maintenance & original spare support.", style_pf_body),
                ]
            ]

            t_portfolio = Table(portfolio_data, colWidths=[136, 136, 136, 137])
            t_portfolio.setStyle(
                TableStyle([
                    ("SPAN", (0, 0), (3, 0)),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E2E8F0")),
                    ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#F1F5F9")),
                    ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#000000")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#000000")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ])
            )
            story.append(t_portfolio)

        # ==========================================
        # 🧱 PAGE 3: COMMERCIAL OFFER & PRICING TABLE (ALWAYS INCLUDED)
        # ==========================================
        story.append(PageBreak())
        story.append(get_header_element())
        story.append(Spacer(1, 8))
        story.append(
            HRFlowable(
                width="100%",
                thickness=1.5,
                color=colors.HexColor("#000000"),
                spaceAfter=8,
            )
        )

        small_bold = ParagraphStyle(
            "SmallBold",
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=10,
            alignment=0,
            textColor=colors.HexColor("#1E293B"),
        )
        small_bold_center = ParagraphStyle(
            "SmallBoldCenter",
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=10,
            alignment=1,
            textColor=colors.HexColor("#1E293B"),
        )
        small_bold_right = ParagraphStyle(
            "SmallBoldRight",
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=10,
            alignment=2,
            textColor=colors.HexColor("#1E293B"),
        )
        small_text = ParagraphStyle(
            "SmallText",
            fontName="Helvetica",
            fontSize=8,
            leading=10,
            alignment=0,
            textColor=colors.HexColor("#334155"),
        )
        small_text_center = ParagraphStyle(
            "SmallTextCenter",
            fontName="Helvetica",
            fontSize=8,
            leading=10,
            alignment=1,
            textColor=colors.HexColor("#334155"),
        )
        small_text_right = ParagraphStyle(
            "SmallTextRight",
            fontName="Helvetica",
            fontSize=8,
            leading=10,
            alignment=2,
            textColor=colors.HexColor("#334155"),
        )

        header_data = [
            [
                Paragraph("<b>Company Name</b>", small_bold),
                Paragraph(client_name, small_text),
                Paragraph("<b>Client ID:</b>", small_bold),
                Paragraph(client_id, small_bold),
            ],
            [
                Paragraph("<b>Contact Details</b>", small_bold),
                Paragraph(contact_details, small_text),
                Paragraph("<b>Qtn Date:</b>", small_bold),
                Paragraph(str(qtn_date), small_text),
            ],
            [
                Paragraph("<b>Shipping Address</b>", small_bold),
                Paragraph(shipping_address, small_text),
                Paragraph("<b>Qtn Ref No:</b>", small_bold),
                Paragraph(qtn_ref_no, small_bold),
            ],
            [
                Paragraph("<b>Billing Address</b>", small_bold),
                Paragraph(billing_address, small_text),
                Paragraph("<b>Sales Reference:</b>", small_bold),
                Paragraph(
                    f"{sales_reference_1}<br/>{sales_reference_2}", small_text
                ),
            ],
            [
                Paragraph("<b>GSTIN</b>", small_bold),
                Paragraph(gstin, small_text),
                Paragraph("<b>Site Contact:</b>", small_bold),
                Paragraph(site_person, small_text),
            ],
        ]

        t_header = Table(header_data, colWidths=[90, 230, 85, 140])
        t_header.setStyle(
            TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
                ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#000000")),
                ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#1E293B")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ])
        )
        story.append(t_header)
        story.append(Spacer(1, 10))

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
            Paragraph("<b>Inst. Amount (INR)</b>", small_bold_right),
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
            # REORDERED OUTPUT: 1st Paint Finish, 2nd Slats
            for s in itm.get("slat_pow", []):
                desc_lines.append(f"- {s}")
            for s in itm.get("slat_nat", []):
                desc_lines.append(f"- {s}")
            for g in itm.get("guide", []):
                desc_lines.append(f"- {g}")
            for b in itm.get("bottom", []):
                desc_lines.append(f"- {b}")
            for h in itm.get("hood", []):
                desc_lines.append(f"- {h}")
            for l in itm.get("safety_locks", []):
                desc_lines.append(f"- {l}")

            if (
                itm.get("type") == "Motorized Rolling Shutter"
                and itm.get("operator")
            ):
                desc_lines.append("<b>Operator Details:</b>")
                for op in itm.get("operator", []):
                    desc_lines.append(f"- {op}")

            desc = "<br/>".join(desc_lines)

            items_data.append([
                Paragraph(str(i + 1), small_text_center),
                Paragraph(desc, small_text),
                Paragraph(itm.get("hsn", "-"), small_text_center),
                Paragraph(str(itm.get("width", "-")), small_text_center),
                Paragraph(str(itm.get("height", "-")), small_text_center),
                Paragraph(str(itm.get("qty", "-")), small_text_center),
                Paragraph(f"{itm.get('mat_rate', 0):,}", small_text_right),
                Paragraph(f"{m_amt:,}", small_text_right),
                Paragraph(f"{itm.get('inst_rate', 0):,}", small_text_right),
                Paragraph(f"{i_amt:,}", small_text_right),
            ])

        subtotal_extras = (
            packing_charges
            + freight_charges
            + unloading_charges
            + crane_charges
            + scaffolding_charges
        )
        subtotal_mat = mat_grand_total + subtotal_extras
        gst_mat = round(subtotal_mat * 0.18)
        gst_inst = round(inst_grand_total * 0.18)

        total_mat_with_gst = subtotal_mat + gst_mat
        total_inst_with_gst = inst_grand_total + gst_inst
        final_grand_total = total_mat_with_gst + total_inst_with_gst

        items_data.append([
            "",
            Paragraph("<b>Item Total</b>", small_bold),
            "",
            "",
            "",
            Paragraph(f"<b>{total_qty}</b>", small_bold_center),
            "",
            Paragraph(f"<b>{mat_grand_total:,}</b>", small_bold_right),
            "",
            Paragraph(f"<b>{inst_grand_total:,}</b>", small_bold_right),
        ])
        items_data.append([
            "",
            Paragraph("Packing Charges", small_text),
            "",
            "",
            "",
            "",
            "",
            Paragraph(f"{packing_charges:,}", small_text_right),
            "",
            Paragraph("-", small_text_right),
        ])
        items_data.append([
            "",
            Paragraph("Freight Charges (As Per Actual)", small_text),
            "",
            "",
            "",
            "",
            "",
            Paragraph(freight_text_display, small_text_right),
            "",
            Paragraph("-", small_text_right),
        ])
        if unloading_charges > 0:
            items_data.append([
                "",
                Paragraph("Unloading Charges", small_text),
                "",
                "",
                "",
                "",
                "",
                Paragraph(f"{unloading_charges:,}", small_text_right),
                "",
                Paragraph("-", small_text_right),
            ])
        if crane_charges > 0:
            items_data.append([
                "",
                Paragraph("Crane Charges", small_text),
                "",
                "",
                "",
                "",
                "",
                Paragraph(f"{crane_charges:,}", small_text_right),
                "",
                Paragraph("-", small_text_right),
            ])
        if scaffolding_charges > 0:
            items_data.append([
                "",
                Paragraph("Scaffolding Charges", small_text),
                "",
                "",
                "",
                "",
                "",
                Paragraph(f"{scaffolding_charges:,}", small_text_right),
                "",
                Paragraph("-", small_text_right),
            ])

        items_data.append([
            "",
            Paragraph(
                "<b>Supply & Installation Amount Excluding GST</b>", small_bold
            ),
            "",
            "",
            "",
            "",
            "",
            Paragraph(f"<b>{subtotal_mat:,}</b>", small_bold_right),
            "",
            Paragraph(f"<b>{inst_grand_total:,}</b>", small_bold_right),
        ])
        items_data.append([
            "",
            Paragraph("GST on supply & installation @18%", small_text),
            "",
            "",
            "",
            "",
            "",
            Paragraph(f"{gst_mat:,}", small_text_right),
            Paragraph("@18%", small_text_right),
            Paragraph(f"{gst_inst:,}", small_text_right),
        ])
        items_data.append([
            "",
            Paragraph("<b>Total supply & installation with GST</b>", small_bold),
            "",
            "",
            "",
            "",
            "",
            Paragraph(f"<b>{total_mat_with_gst:,}</b>", small_bold_right),
            "",
            Paragraph(f"<b>{total_inst_with_gst:,}</b>", small_bold_right),
        ])
        items_data.append([
            "",
            Paragraph("<b>Grand total with GST</b>", small_bold),
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            Paragraph(f"<b>{final_grand_total:,}</b>", small_bold_right),
        ])

        t_items = Table(
            items_data, colWidths=[22, 160, 50, 38, 38, 22, 55, 60, 50, 60]
        )
        t_items.setStyle(
            TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E2E8F0")),
                ("BOX", (0, 0), (-1, -1), 1, colors.black),
                ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#334155")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ])
        )
        story.append(t_items)

        # ==========================================
        # ⚙️ PAGE 4: SECTION C: TECHNICAL SPECS (CONDITIONAL)
        # ==========================================
        if include_page_4:
            story.append(PageBreak())
            story.append(get_header_element())
            story.append(Spacer(1, 6))
            story.append(
                HRFlowable(
                    width="100%",
                    thickness=1.5,
                    color=colors.HexColor("#000000"),
                    spaceAfter=8,
                )
            )

            story.append(Paragraph("Section C: Technical Specifications", style_sec_title))
            story.append(Spacer(1, 4))

            style_intro_text = ParagraphStyle(
                "IntroText",
                parent=styles["Normal"],
                fontName="Helvetica",
                fontSize=9.5,
                leading=13,
                textColor=colors.HexColor("#1E293B"),
            )
            story.append(Paragraph(tech_intro_text, style_intro_text))
            story.append(Spacer(1, 6))

            style_th_param = ParagraphStyle(
                "ThParam",
                parent=styles["Normal"],
                fontName="Helvetica-Bold",
                fontSize=9.5,
                leading=12,
                alignment=0,
                textColor=colors.HexColor("#000000"),
            )
            style_th_spec = ParagraphStyle(
                "ThSpec",
                parent=styles["Normal"],
                fontName="Helvetica-Bold",
                fontSize=9.5,
                leading=12,
                alignment=1,
                textColor=colors.HexColor("#000000"),
            )
            style_td_param = ParagraphStyle(
                "TdParam",
                parent=styles["Normal"],
                fontName="Helvetica-Bold",
                fontSize=9,
                leading=11.5,
                alignment=0,
                textColor=colors.HexColor("#0F172A"),
            )
            style_td_spec = ParagraphStyle(
                "TdSpec",
                parent=styles["Normal"],
                fontName="Helvetica",
                fontSize=9,
                leading=11.5,
                alignment=0,
                textColor=colors.HexColor("#334155"),
            )

            tech_table_data = [[
                Paragraph("Parameter", style_th_param),
                Paragraph("Specification Details", style_th_spec)
            ]]

            for spec_item in st.session_state["tech_specs_data"]:
                tech_table_data.append([
                    Paragraph(spec_item["param"], style_td_param),
                    Paragraph(spec_item["spec"], style_td_spec)
                ])

            t_tech = Table(tech_table_data, colWidths=[140, 405])
            t_tech.setStyle(
                TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E2E8F0")),
                    ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#000000")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#000000")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ])
            )
            story.append(t_tech)
            story.append(Spacer(1, 6))

            style_note_text = ParagraphStyle(
                "NoteText",
                parent=styles["Normal"],
                fontName="Helvetica-Oblique",
                fontSize=8.5,
                leading=11,
                textColor=colors.HexColor("#475569"),
            )
            story.append(Paragraph(tech_note_text, style_note_text))

        # ==========================================
        # 📋 PAGE 5: SECTION D: TERMS & CONDITIONS (CONDITIONAL)
        # ==========================================
        if include_page_5:
            story.append(PageBreak())
            story.append(get_header_element())
            story.append(Spacer(1, 6))
            story.append(
                HRFlowable(
                    width="100%",
                    thickness=1.5,
                    color=colors.HexColor("#000000"),
                    spaceAfter=8,
                )
            )

            story.append(Paragraph("Section D: Annexure – Terms & Condition", style_sec_title))
            story.append(Spacer(1, 4))

            style_subhead = ParagraphStyle(
                "SubHeadEx",
                parent=styles["Normal"],
                fontName="Helvetica-Bold",
                fontSize=10,
                leading=13,
                textColor=colors.HexColor("#000000"),
            )
            story.append(Paragraph(f"<b>{exclusions_subhead}</b>", style_subhead))
            story.append(Spacer(1, 6))

            style_tc_cat = ParagraphStyle(
                "TcCat",
                parent=styles["Normal"],
                fontName="Helvetica-Bold",
                fontSize=9,
                leading=11.5,
                alignment=0,
                textColor=colors.HexColor("#000000"),
            )
            style_tc_det = ParagraphStyle(
                "TcDet",
                parent=styles["Normal"],
                fontName="Helvetica",
                fontSize=8.5,
                leading=11,
                alignment=0,
                textColor=colors.HexColor("#1E293B"),
            )

            terms_table_data = []

            for term in st.session_state["terms_data"]:
                details_formatted = term["details"].replace("\n", "<br/>").replace("•", "&bull;")
                terms_table_data.append([
                    Paragraph(f"<b>{term['category']}</b>", style_tc_cat),
                    Paragraph(details_formatted, style_tc_det)
                ])

            t_terms = Table(terms_table_data, colWidths=[135, 410])
            t_terms.setStyle(
                TableStyle([
                    ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#000000")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#000000")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ])
            )
            story.append(t_terms)

            story.append(Spacer(1, 10))
            style_sign_box = ParagraphStyle(
                "SignBox",
                parent=styles["Normal"],
                fontName="Helvetica-Bold",
                fontSize=9,
                leading=12,
                textColor=colors.HexColor("#000000"),
            )
            sign_box_content = [
                Paragraph("<b>ACCEPTANCE OF OFFER & ORDER CONFIRMATION</b>", style_sign_box),
                Paragraph("<font size=8>We hereby accept the commercial offer, technical specifications, and terms & conditions outlined above.</font>", style_tc_det),
                Spacer(1, 15),
                Paragraph("<b>Client Seal & Signature: _______________________ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Date: _______________</b>", style_sign_box)
            ]
            t_sign = Table([[sign_box_content]], colWidths=[545])
            t_sign.setStyle(
                TableStyle([
                    ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#000000")),
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ])
            )
            story.append(t_sign)

        doc.build(story)
        buffer.seek(0)
        return buffer

    # --- FORMATTED EXCEL GENERATOR ---
    def generate_excel():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Quotation"

        font_company = Font(name="Calibri", size=14, bold=True, color="000000")
        font_address = Font(name="Calibri", size=9, italic=True)
        font_header_bold = Font(name="Calibri", size=10, bold=True)
        font_regular = Font(name="Calibri", size=9)
        font_bold = Font(name="Calibri", size=9, bold=True)

        fill_table_header = PatternFill(
            start_color="E2E8F0", end_color="E2E8F0", fill_type="solid"
        )
        fill_total_row = PatternFill(
            start_color="F1F5F9", end_color="F1F5F9", fill_type="solid"
        )

        thin_border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000"),
        )

        align_center = Alignment(horizontal="center", vertical="top", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="top", wrap_text=True)
        align_right = Alignment(horizontal="right", vertical="top", wrap_text=True)

        ws.merge_cells("A1:J1")
        ws["A1"] = "SIDHARTH SHUTTER & AUTOMATION PRIVATE LIMITED"
        ws["A1"].font = font_company
        ws["A1"].alignment = Alignment(horizontal="center")

        ws.merge_cells("A2:J2")
        ws["A2"] = (
            "GSTIN: 08AEEPJ6848R1ZN | www.ssaapl.com | sales@ssaapl.com | H-1-89,"
            " RIICO Ind. Area, Mansarovar, Jaipur"
        )
        ws["A2"].font = font_address
        ws["A2"].alignment = Alignment(horizontal="center")

        ws.row_dimensions[1].height = 25
        ws.row_dimensions[2].height = 18

        client_info = [
            [
                ("Company Name:", font_header_bold),
                (client_name, font_regular),
                ("Client ID:", font_header_bold),
                (client_id, font_header_bold),
            ],
            [
                ("Contact Details:", font_header_bold),
                (contact_details, font_regular),
                ("Qtn Date:", font_header_bold),
                (str(qtn_date), font_regular),
            ],
            [
                ("Shipping Address:", font_header_bold),
                (shipping_address, font_regular),
                ("Qtn Ref No:", font_header_bold),
                (qtn_ref_no, font_header_bold),
            ],
            [
                ("Billing Address:", font_header_bold),
                (billing_address, font_regular),
                ("Sales Reference:", font_header_bold),
                (f"{sales_reference_1} / {sales_reference_2}", font_regular),
            ],
            [
                ("Client GSTIN:", font_header_bold),
                (gstin, font_regular),
                ("Quotation Made By:", font_header_bold),
                (quotation_made_by, font_header_bold),
            ],
        ]

        curr_row = 4
        for row in client_info:
            ws.cell(row=curr_row, column=1, value=row[0][0]).font = row[0][1]

            ws.merge_cells(
                start_row=curr_row, start_column=2, end_row=curr_row, end_column=5
            )
            ws.cell(row=curr_row, column=2, value=row[1][0]).font = row[1][1]
            ws.cell(row=curr_row, column=2).alignment = align_left

            ws.cell(row=curr_row, column=6, value=row[2][0]).font = row[2][1]

            ws.merge_cells(
                start_row=curr_row, start_column=7, end_row=curr_row, end_column=10
            )
            ws.cell(row=curr_row, column=7, value=row[3][0]).font = row[3][1]
            ws.cell(row=curr_row, column=7).alignment = align_left

            for c in range(1, 11):
                ws.cell(row=curr_row, column=c).border = thin_border

            curr_row += 1

        curr_row += 1

        headers = [
            "Sr. No.",
            "Description",
            "HSN Code",
            "Width (mm)",
            "Height (mm)",
            "Qty",
            "Mat. Rate (INR)",
            "Mat. Amount (INR)",
            "Inst. Rate (INR)",
            "Inst. Amount (INR)",
        ]
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

            desc_lines = [itm.get("type", "Motorized Rolling Shutter")]
            # REORDERED EXCEL OUTPUT: 1st Paint Finish, 2nd Slats
            for s in itm.get("slat_pow", []):
                desc_lines.append(f"• {s}")
            for s in itm.get("slat_nat", []):
                desc_lines.append(f"• {s}")
            for g in itm.get("guide", []):
                desc_lines.append(f"• {g}")
            for b in itm.get("bottom", []):
                desc_lines.append(f"• {b}")
            for h in itm.get("hood", []):
                desc_lines.append(f"• {h}")
            for l in itm.get("safety_locks", []):
                desc_lines.append(f"• {l}")

            if (
                itm.get("type") == "Motorized Rolling Shutter"
                and itm.get("operator")
            ):
                desc_lines.append("Operator Details:")
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
                i_amt,
            ]

            for col_num, val in enumerate(row_data, 1):
                cell = ws.cell(row=curr_row, column=col_num, value=val)
                cell.font = font_regular
                cell.border = thin_border

                if col_num in [1, 3, 4, 5, 6]:
                    cell.alignment = align_center
                elif col_num in [7, 8, 9, 10]:
                    cell.alignment = align_right
                    cell.number_format = "#,##0"
                else:
                    cell.alignment = align_left

            curr_row += 1

        subtotal_extras = (
            packing_charges
            + freight_charges
            + unloading_charges
            + crane_charges
            + scaffolding_charges
        )
        subtotal_mat = mat_grand_total + subtotal_extras
        gst_mat = round(subtotal_mat * 0.18)
        gst_inst = round(inst_grand_total * 0.18)

        total_mat_with_gst = subtotal_mat + gst_mat
        total_inst_with_gst = inst_grand_total + gst_inst
        final_grand_total = total_mat_with_gst + total_inst_with_gst

        def add_summary_row(
            label, mat_val, inst_val, is_bold=False, is_fill=False
        ):
            nonlocal curr_row
            ws.cell(row=curr_row, column=2, value=label).font = (
                font_bold if is_bold else font_regular
            )
            ws.cell(row=curr_row, column=2).alignment = align_left

            if label == "Item Total":
                ws.cell(row=curr_row, column=6, value=total_qty).font = font_bold
                ws.cell(row=curr_row, column=6).alignment = align_center

            cell_mat = ws.cell(
                row=curr_row,
                column=8,
                value=mat_val if isinstance(mat_val, (int, float)) else mat_val,
            )
            cell_mat.font = font_bold if is_bold else font_regular
            cell_mat.alignment = align_right
            if isinstance(mat_val, (int, float)):
                cell_mat.number_format = "#,##0"

            cell_inst = ws.cell(
                row=curr_row,
                column=10,
                value=inst_val if isinstance(inst_val, (int, float)) else inst_val,
            )
            cell_inst.font = font_bold if is_bold else font_regular
            cell_inst.alignment = align_right
            if isinstance(inst_val, (int, float)):
                cell_inst.number_format = "#,##0"

            for c in range(1, 11):
                cell = ws.cell(row=curr_row, column=c)
                cell.border = thin_border
                if is_fill:
                    cell.fill = fill_total_row
            curr_row += 1

        add_summary_row(
            "Item Total", mat_grand_total, inst_grand_total, is_bold=True
        )
        add_summary_row("Packing Charges", packing_charges, "-")
        add_summary_row("Freight Charges (As Per Actual)", freight_text_display, "-")
        if unloading_charges > 0:
            add_summary_row("Unloading Charges", unloading_charges, "-")
        if crane_charges > 0:
            add_summary_row("Crane Charges", crane_charges, "-")
        if scaffolding_charges > 0:
            add_summary_row("Scaffolding Charges", scaffolding_charges, "-")

        add_summary_row(
            "Supply & Installation Amount Excluding GST",
            subtotal_mat,
            inst_grand_total,
            is_bold=True,
            is_fill=True,
        )
        add_summary_row("GST on supply & installation @18%", gst_mat, gst_inst)
        add_summary_row(
            "Total supply & installation with GST",
            total_mat_with_gst,
            total_inst_with_gst,
            is_bold=True,
            is_fill=True,
        )
        add_summary_row(
            "Grand total with GST", "", final_grand_total, is_bold=True, is_fill=True
        )

        col_widths = {
            1: 8,
            2: 45,
            3: 12,
            4: 12,
            5: 12,
            6: 8,
            7: 15,
            8: 18,
            9: 15,
            10: 18,
        }
        for col_idx, width in col_widths.items():
            ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = (
                width
            )

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    # --- DOWNLOAD & PREVIEW BUTTONS SECTION ---
    st.markdown("---")
    st.markdown("### 📥 Generate & Export Proposals")

    col_pdf, col_excel = st.columns(2)

    with col_pdf:
        if st.button(
            "👁️ Preview PDF Quotation", type="primary", use_container_width=True
        ):
            pdf_bytes = generate_pdf().getvalue()
            st.session_state["pdf_preview_bytes"] = pdf_bytes

    with col_excel:
        if st.button("📊 Generate Excel Quotation", use_container_width=True):
            excel_buffer = generate_excel()
            st.download_button(
                label="📥 Download Quotation Excel (.xlsx)",
                data=excel_buffer,
                file_name=f"Quotation_{qtn_ref_no.replace('/', '_')}.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

    # --- PDF PREVIEW AND DOWNLOAD SECTION ---
    if "pdf_preview_bytes" in st.session_state:
        st.markdown("---")
        st.markdown("### 🔍 PDF Quotation Preview & Download")

        pdf_data = st.session_state["pdf_preview_bytes"]

        c_dl, c_cls = st.columns([2, 1])
        with c_dl:
            st.download_button(
                label="📥 Download Verified Multi-Page PDF Quotation",
                data=pdf_data,
                file_name=f"Quotation_{qtn_ref_no.replace('/', '_')}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True,
            )
        with c_cls:
            if st.button("❌ Close Preview", use_container_width=True):
                del st.session_state["pdf_preview_bytes"]
                st.rerun()

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

    st.info(
        f"⏳ **{st.session_state['selected_product']} Quotation Builder** is"
        " currently under development. Please check back soon!"
    )
