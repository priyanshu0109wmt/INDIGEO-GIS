import os
import shutil
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Paths
TEMPLATE_PATH = r"D:\SEM 7\Review 3- 28-10- 2026.pptx"
OUTPUT_PATH = r"D:\SEM 7\indigeo-gis\priyanshu ppt.pptx"

# Color Palette (Crisp, Academic, Neat)
NAVY_PRIMARY = RGBColor(15, 34, 64)       # Deep Navy #0F2240
BLUE_ACCENT  = RGBColor(30, 80, 160)      # Academic Blue #1E50A0
BLUE_LIGHT   = RGBColor(242, 246, 252)    # Soft Card Fill #F2F6FC
TEXT_DARK    = RGBColor(25, 30, 36)       # Near Black
TEXT_MUTED   = RGBColor(80, 90, 105)      # Slate Gray
BORDER_LIGHT = RGBColor(210, 220, 232)    # Subtle Card Border
ROW_ALT_FILL = RGBColor(247, 250, 254)    # Subtle Table Alt Row
WHITE        = RGBColor(255, 255, 255)
GREEN_DARK   = RGBColor(16, 128, 64)      # Success Green #108040
RED_ACCENT   = RGBColor(185, 35, 35)      # Gap / Warning Red

FONT_FAMILY = "Times New Roman"

def set_run_font(run, size_pt, bold=False, italic=False, color=TEXT_DARK, font_name=FONT_FAMILY):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color

def create_card(slide, left, top, width, height, bg_color=BLUE_LIGHT, border_color=BORDER_LIGHT):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    return shape

def format_cell(cell, text, font_size=10, bold=False, italic=False, text_color=TEXT_DARK, bg_color=None, align=PP_ALIGN.LEFT):
    if bg_color:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg_color
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = cell.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)
    tf.text = ""
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run_font(run, font_size, bold=bold, italic=italic, color=text_color)

def remove_shape(slide, shape_name):
    for s in list(slide.shapes):
        if s.name == shape_name:
            slide.shapes._spTree.remove(s._element)
            return True
    return False

def main():
    print(f"Loading template from: {TEMPLATE_PATH}")
    prs = pptx.Presentation(TEMPLATE_PATH)
    
    # =============================================================
    # SLIDE 1: TITLE SLIDE
    # =============================================================
    print("Formatting Slide 1 (Title)...")
    s1 = prs.slides[0]
    
    # Title 6 Placeholder: Project Header and Title
    for s in s1.shapes:
        if s.name == "Title 6":
            tf = s.text_frame
            tf.word_wrap = True
            tf.text = ""
            
            p1 = tf.paragraphs[0]
            p1.alignment = PP_ALIGN.CENTER
            r1 = p1.add_run()
            r1.text = "B.Tech. Computer Science and Engineering  |  BCSE497J - Project-I (Review 3)"
            set_run_font(r1, 15, bold=True, color=BLUE_ACCENT)
            
            p2 = tf.add_paragraph()
            p2.alignment = PP_ALIGN.CENTER
            p2.space_before = Pt(6)
            r2 = p2.add_run()
            r2.text = "GEOGRAPHICAL INFORMATION SYSTEM BASED DATA MANAGEMENT SYSTEM\n(INDIGEO-GIS)"
            set_run_font(r2, 21, bold=True, color=NAVY_PRIMARY)

    # Remove template object 3 and create 2 structured, neat cards
    remove_shape(s1, "object 3")
    
    # Left Card: Team Members
    card_tm = create_card(s1, Inches(1.6), Inches(3.55), Inches(4.8), Inches(2.65))
    tf_tm = card_tm.text_frame
    tf_tm.word_wrap = True
    tf_tm.margin_left = Inches(0.2)
    tf_tm.margin_right = Inches(0.2)
    tf_tm.margin_top = Inches(0.15)
    
    p = tf_tm.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "PROJECT TEAM MEMBERS"
    set_run_font(r, 13, bold=True, color=NAVY_PRIMARY)
    
    tm_list = [
        ("PRIYANSHU", "23BCE0303"),
        ("UJJAWAL KUMAR", "23BCE0354"),
        ("PIYUSH RAJ", "23BCE0459")
    ]
    for name, reg in tm_list:
        pt = tf_tm.add_paragraph()
        pt.alignment = PP_ALIGN.CENTER
        pt.space_before = Pt(8)
        rn = pt.add_run()
        rn.text = f"{name}  "
        set_run_font(rn, 12, bold=True, color=TEXT_DARK)
        rr = pt.add_run()
        rr.text = f"({reg})"
        set_run_font(rr, 11, bold=False, color=BLUE_ACCENT)
        
    p_inst = tf_tm.add_paragraph()
    p_inst.alignment = PP_ALIGN.CENTER
    p_inst.space_before = Pt(10)
    ri = p_inst.add_run()
    ri.text = "School of Computer Science and Engineering (SCOPE)"
    set_run_font(ri, 10, italic=True, color=TEXT_MUTED)
    
    # Right Card: Faculty Guide
    card_fg = create_card(s1, Inches(6.9), Inches(3.55), Inches(4.8), Inches(2.65))
    tf_fg = card_fg.text_frame
    tf_fg.word_wrap = True
    tf_fg.margin_left = Inches(0.2)
    tf_fg.margin_right = Inches(0.2)
    tf_fg.margin_top = Inches(0.15)
    
    p = tf_fg.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "UNDER THE SUPERVISION OF"
    set_run_font(r, 13, bold=True, color=NAVY_PRIMARY)
    
    pg1 = tf_fg.add_paragraph()
    pg1.alignment = PP_ALIGN.CENTER
    pg1.space_before = Pt(8)
    rg1 = pg1.add_run()
    rg1.text = "Dr. NARESH K"
    set_run_font(rg1, 14, bold=True, color=TEXT_DARK)
    
    pg2 = tf_fg.add_paragraph()
    pg2.alignment = PP_ALIGN.CENTER
    pg2.space_before = Pt(4)
    rg2 = pg2.add_run()
    rg2.text = "Associate Professor, Department of Software Systems\nSchool of Computer Science and Engineering (SCOPE)\nVellore Institute of Technology (VIT), Vellore – 632014"
    set_run_font(rg2, 10.5, italic=True, color=TEXT_MUTED)

    # =============================================================
    # SLIDE 2: APPROVAL MAIL FROM GUIDE (LEAVE EXACTLY AS IS!)
    # =============================================================
    print("Slide 2 left intact as in template (Approval Mail From Guide).")

    # =============================================================
    # SLIDE 3: AIM
    # =============================================================
    print("Formatting Slide 3 (Aim)...")
    s3 = prs.slides[2]
    remove_shape(s3, "Content Placeholder 2")
    
    aims = [
        ("Unified Web-GIS Data Management Architecture",
         "Design and deploy an end-to-end, high-performance, open-source Web-GIS platform engineered specifically for nationwide multi-hazard risk assessment and spatial data management across India."),
        ("Sub-Second Synchronous Spatial Query Engine",
         "Deliver coordinate-level point hazard profiling within 0.27s – 0.80s across 760 Indian districts, integrating in-memory Ray-Casting with live cloud telemetry (AWS Terrarium DEM, USGS ComCat, Open-Meteo)."),
        ("Concurrent Multi-Hazard Risk Evaluation",
         "Transition from isolated single-hazard research models to a concurrent 8-vector hazard evaluation engine (Flood, Coastal Surge, Earthquake, Cyclone, Drought, Heatwave, Landslide, Wildfire) normalized to [0.0, 1.0]."),
        ("Actionable Emergency Logistics & Resilient Failover",
         "Bridge passive hazard map visualization with active emergency operations by mapping risk scorecards to 13 evidence-based preparedness advisories, backed by an async multi-mirror OpenStreetMap Overpass failover cluster (99.5% uptime)."),
        ("Zero-Synthetic Data Integrity & Provenance Tracking",
         "Enforce strict end-to-end telemetry provenance with explicit 'unverified' flags for unconfigured sensor feeds (e.g. NASA FIRMS), eliminating fabricated fallback values in life-critical decision making.")
    ]
    
    top_pos = Inches(1.55)
    card_height = Inches(0.85)
    spacing = Inches(0.12)
    card_width = Inches(11.8)
    left_pos = Inches(0.76)
    
    for idx, (title, desc) in enumerate(aims):
        curr_top = top_pos + idx * (card_height + spacing)
        card = create_card(s3, left_pos, curr_top, card_width, card_height)
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.2)
        tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.08)
        tf.margin_bottom = Inches(0.08)
        
        p = tf.paragraphs[0]
        r_num = p.add_run()
        r_num.text = f"Goal {idx+1}: "
        set_run_font(r_num, 12.5, bold=True, color=BLUE_ACCENT)
        
        r_title = p.add_run()
        r_title.text = f"{title}\n"
        set_run_font(r_title, 12.5, bold=True, color=NAVY_PRIMARY)
        
        p_desc = tf.add_paragraph()
        p_desc.space_before = Pt(2)
        r_desc = p_desc.add_run()
        r_desc.text = desc
        set_run_font(r_desc, 10.5, bold=False, color=TEXT_DARK)

    # =============================================================
    # SLIDE 4: ABSTRACT
    # =============================================================
    print("Formatting Slide 4 (Abstract)...")
    s4 = prs.slides[3]
    remove_shape(s4, "Content Placeholder 2")
    
    # Left Card: Problem & Context
    card_left = create_card(s4, Inches(0.76), Inches(1.55), Inches(5.75), Inches(4.65))
    tf_left = card_left.text_frame
    tf_left.word_wrap = True
    tf_left.margin_left = Inches(0.2)
    tf_left.margin_right = Inches(0.2)
    tf_left.margin_top = Inches(0.15)
    
    p = tf_left.paragraphs[0]
    r = p.add_run()
    r.text = "RESEARCH PROBLEM & CONTEXT"
    set_run_font(r, 13, bold=True, color=NAVY_PRIMARY)
    
    bullets_left = [
        ("Fragmentation in Existing Portals", "Conventional Web-GIS disaster systems only provide passive map visualizations, lacking spatial data versioning, dynamic schema flexibility, and dataset audit tracking."),
        ("Rigid Relational Schemas", "Traditional PostGIS/GeoServer setups rely on hardcoded table columns that crash during field uploads of heterogeneous spatial files with varying coordinate systems."),
        ("Offline Machine Learning Trap", "Cutting-edge GeoAI and Explainable AI (XAI) models remain confined to offline Python notebooks (>80s execution), isolated from live user-facing Web-GIS pipelines."),
        ("Single-Hazard Isolation", "Existing models evaluate hazards in isolation, neglecting cascading compound disasters (e.g. cyclonic rainfall causing coastal surge and landslide triggers).")
    ]
    for title, body in bullets_left:
        pb = tf_left.add_paragraph()
        pb.space_before = Pt(8)
        rt = pb.add_run()
        rt.text = f"• {title}: "
        set_run_font(rt, 11, bold=True, color=BLUE_ACCENT)
        rb = pb.add_run()
        rb.text = body
        set_run_font(rb, 10, color=TEXT_DARK)
        
    # Right Card: Proposed Solution & Key Results
    card_right = create_card(s4, Inches(6.8), Inches(1.55), Inches(5.75), Inches(4.65))
    tf_right = card_right.text_frame
    tf_right.word_wrap = True
    tf_right.margin_left = Inches(0.2)
    tf_right.margin_right = Inches(0.2)
    tf_right.margin_top = Inches(0.15)
    
    p = tf_right.paragraphs[0]
    r = p.add_run()
    r.text = "PROPOSED SOLUTION & KEY OUTCOMES"
    set_run_font(r, 13, bold=True, color=NAVY_PRIMARY)
    
    bullets_right = [
        ("Decoupled Client-Server Platform", "Engineered INDIGEO-GIS with a responsive Next.js 16 / React 19 / Leaflet thin-client frontend and an asynchronous Python FastAPI backend orchestrating live cloud APIs."),
        ("Nationwide Spatial Ingestion", "Automated GeoPandas & Shapely pipeline standardizing 760 Indian administrative boundaries in WGS 84 (EPSG:4326) with sub-40ms Ray-Casting point queries."),
        ("Sub-Second Synchronous Profiling", "Achieved 0.54s average query latency (0.27s–0.80s), accelerating spatial hazard profiling by 21x to 150x compared to legacy GIS and offline notebooks."),
        ("Actionable Logistics & Resilience", "Connected multi-hazard risk scores to 13 evidence-based preparedness advisories, backed by an async Overpass failover cluster (99.5% uptime)."),
        ("Zero-Synthetic Integrity & TCO Savings", "Strict data provenance with zero synthetic substitutions, saving $285,000 USD over 5-year TCO via a 100% open-source stack.")
    ]
    for title, body in bullets_right:
        pb = tf_right.add_paragraph()
        pb.space_before = Pt(6)
        rt = pb.add_run()
        rt.text = f"• {title}: "
        set_run_font(rt, 11, bold=True, color=GREEN_DARK)
        rb = pb.add_run()
        rb.text = body
        set_run_font(rb, 10, color=TEXT_DARK)

    # =============================================================
    # SLIDE 5: LITERATURE REVIEW
    # =============================================================
    print("Formatting Slide 5 (Literature Review)...")
    s5 = prs.slides[4]
    remove_shape(s5, "Content Placeholder 4")
    remove_shape(s5, "Picture 6")
    
    # Top Card 1: Base Paper 1
    bp1_card = create_card(s5, Inches(0.76), Inches(1.5), Inches(5.75), Inches(1.6))
    tf1 = bp1_card.text_frame
    tf1.word_wrap = True
    tf1.margin_left = Inches(0.15)
    tf1.margin_right = Inches(0.15)
    tf1.margin_top = Inches(0.1)
    p = tf1.paragraphs[0]
    r = p.add_run()
    r.text = "BASE PAPER 1: Web-GIS Systematic Review (2024)"
    set_run_font(r, 11.5, bold=True, color=NAVY_PRIMARY)
    p_ref = tf1.add_paragraph()
    r_ref = p_ref.add_run()
    r_ref.text = "A. F. Al-Hussein et al., MDPI Sustainability, vol. 16, no. 10, p. 4238, May 2024."
    set_run_font(r_ref, 9.5, italic=True, color=TEXT_MUTED)
    p_body = tf1.add_paragraph()
    p_body.space_before = Pt(3)
    r_b = p_body.add_run()
    r_b.text = "• Findings: Evaluated global Web-GIS hazard platforms.\n• Gaps Exposed: Rigid relational schemas crash during field uploads; high latency (>10s); passive visualization without emergency logistics or change tracking."
    set_run_font(r_b, 9.5, color=TEXT_DARK)
    
    # Top Card 2: Base Paper 2
    bp2_card = create_card(s5, Inches(6.8), Inches(1.5), Inches(5.75), Inches(1.6))
    tf2 = bp2_card.text_frame
    tf2.word_wrap = True
    tf2.margin_left = Inches(0.15)
    tf2.margin_right = Inches(0.15)
    tf2.margin_top = Inches(0.1)
    p = tf2.paragraphs[0]
    r = p.add_run()
    r.text = "BASE PAPER 2: Interpretable ML for Floods (2024)"
    set_run_font(r, 11.5, bold=True, color=NAVY_PRIMARY)
    p_ref = tf2.add_paragraph()
    r_ref = p_ref.add_run()
    r_ref.text = "Z. Chen et al., MDPI Remote Sensing, vol. 16, no. 20, p. 3902, Oct. 2024."
    set_run_font(r_ref, 9.5, italic=True, color=TEXT_MUTED)
    p_body = tf2.add_paragraph()
    p_body.space_before = Pt(3)
    r_b = p_body.add_run()
    r_b.text = "• Findings: Tree boosting (RF, XGBoost, LightGBM) + SHAP explainability for flood susceptibility.\n• Gaps Exposed: Restricted to offline Python scripts (>80s); single-hazard scope only; zero integration with live web-GIS, spatial query systems, or evacuation routing."
    set_run_font(r_b, 9.5, color=TEXT_DARK)

    # Bottom Table: Comparative Benchmark Matrix (Table 2.1 in report)
    t_top = Inches(3.25)
    t_width = Inches(11.8)
    t_height = Inches(2.9)
    tbl_shape = s5.shapes.add_table(5, 6, Inches(0.76), t_top, t_width, t_height)
    tbl = tbl_shape.table
    
    col_widths = [Inches(2.5), Inches(1.9), Inches(1.7), Inches(1.6), Inches(1.9), Inches(2.2)]
    for idx, w in enumerate(col_widths):
        tbl.columns[idx].width = w
        
    headers = ["Platform / Benchmark", "Spatial Data Model", "Hazard Scope", "Query Latency", "Explainability (XAI)", "Actionable Logistics"]
    for col_idx, h in enumerate(headers):
        format_cell(tbl.cell(0, col_idx), h, font_size=10, bold=True, text_color=WHITE, bg_color=NAVY_PRIMARY, align=PP_ALIGN.CENTER)
        
    matrix_data = [
        ("MDPI Sustainability 2024\n[Base Paper 1]", "Relational tables (WMS/WFS)", "Single/Dual (Flood/Fire)", "High (>10 s)", "None (Static Maps)", "None (Passive View)"),
        ("MDPI Remote Sensing 2024\n[Base Paper 2]", "Static Raster GeoTIFF Grids", "Single Hazard (Flood Only)", "Offline (>82.5 s)", "SHAP (Static Notebook)", "None (No Evacuation Logic)"),
        ("State Disaster Portals (India)\n[Bandrova et al. 2018]", "Legacy Relational / Fixed Schema", "Flood / Cyclone Disconnected", "High (8 - 25 s)", "None", "Static PDF Bulletins / Basic Overlay"),
        ("INDIGEO-GIS\n[Proposed System]", "Dynamic GeoJSON + Fast In-Memory", "8 Concurrent Hazard Vectors", "Sub-Second (0.27 - 0.80 s)", "Roadmap SHAP + Zero Fallback", "13 Action Rules + OR-Tools / NetworkX")
    ]
    
    for row_idx, r_data in enumerate(matrix_data):
        is_indigeo = (row_idx == 3)
        bg = RGBColor(230, 245, 235) if is_indigeo else (ROW_ALT_FILL if row_idx % 2 == 1 else WHITE)
        t_col = GREEN_DARK if is_indigeo else TEXT_DARK
        bld = is_indigeo
        for col_idx, val in enumerate(r_data):
            format_cell(tbl.cell(row_idx+1, col_idx), val, font_size=9.5, bold=bld, text_color=t_col, bg_color=bg, align=PP_ALIGN.CENTER if col_idx > 0 else PP_ALIGN.LEFT)

    # =============================================================
    # SLIDE 6: RESEARCH GAP
    # =============================================================
    print("Formatting Slide 6 (Research Gap)...")
    s6 = prs.slides[5]
    remove_shape(s6, "Content Placeholder 2")
    
    gaps = [
        ("Gap 1: Rigid Relational Schemas & Ingestion Failures",
         "Traditional Web-GIS portals crash on heterogeneous spatial files with varying attributes and CRS.",
         "Automated GeoPandas/Shapely ingestion with instant reprojection to EPSG:4326 and polygon geometry repair."),
        ("Gap 2: Zero Spatial History & Boundary Overwrite",
         "Updating hazard boundaries permanently destroys past geometries, removing any temporal audit trail.",
         "Spatial dataset versioning, metadata provenance tracking, and geometry rollback engine."),
        ("Gap 3: Passive Visualization Without Decision Support",
         "Platforms render thematic risk polygons but offer zero actionable guidance or emergency logistics.",
         "Unified decision support mapping risk vectors directly to 13 evidence-based preparedness advisories."),
        ("Gap 4: Single-Hazard Isolation & Missing Compound Triggers",
         "Existing ML models evaluate hazards in isolation, missing cascading compound multi-disasters.",
         "Unified Multi-Hazard Engine evaluating 8 vectors concurrently (Flood, Coastal, Quake, Cyclone, etc.)."),
        ("Gap 5: Offline Notebook Execution vs. Live Spatial Queries",
         "Advanced ML models take >80s in static scripts, completely unusable for live user-facing Web-GIS.",
         "Asynchronous FastAPI backend coupled with Next.js/Leaflet delivering live point queries in 0.27s–0.80s."),
        ("Gap 6: Blind Assumption of Complete Remote Sensing Data",
         "Pipelines assume complete data streams and fail silently or hallucinate synthetic risk numbers.",
         "Transparent coverage monitoring and explicit 'unverified' flags for missing keys (e.g. NASA FIRMS).")
    ]
    
    card_w = Inches(5.75)
    card_h = Inches(1.42)
    for idx, (title, gap_desc, sol_desc) in enumerate(gaps):
        col = idx % 2
        row = idx // 2
        l = Inches(0.76) if col == 0 else Inches(6.8)
        t = Inches(1.55) + row * (card_h + Inches(0.15))
        
        card = create_card(s6, l, t, card_w, card_h)
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.15)
        tf.margin_right = Inches(0.15)
        tf.margin_top = Inches(0.08)
        tf.margin_bottom = Inches(0.08)
        
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = title
        set_run_font(r, 11, bold=True, color=NAVY_PRIMARY)
        
        pg = tf.add_paragraph()
        pg.space_before = Pt(2)
        rg_lbl = pg.add_run()
        rg_lbl.text = "• Gap: "
        set_run_font(rg_lbl, 10, bold=True, color=RED_ACCENT)
        rg_txt = pg.add_run()
        rg_txt.text = gap_desc
        set_run_font(rg_txt, 9.5, color=TEXT_DARK)
        
        ps = tf.add_paragraph()
        ps.space_before = Pt(2)
        rs_lbl = ps.add_run()
        rs_lbl.text = "• Solution: "
        set_run_font(rs_lbl, 10, bold=True, color=GREEN_DARK)
        rs_txt = ps.add_run()
        rs_txt.text = sol_desc
        set_run_font(rs_txt, 9.5, color=TEXT_DARK)

    # =============================================================
    # SLIDE 7: OBJECTIVES
    # =============================================================
    print("Formatting Slide 7 (Objectives)...")
    s7 = prs.slides[6]
    remove_shape(s7, "Content Placeholder 2")
    remove_shape(s7, "TextBox 6")
    
    # Top Card: 5 Core Technical Objectives
    top_box = create_card(s7, Inches(0.76), Inches(1.45), Inches(11.8), Inches(2.75))
    tf_top = top_box.text_frame
    tf_top.word_wrap = True
    tf_top.margin_left = Inches(0.2)
    tf_top.margin_right = Inches(0.2)
    tf_top.margin_top = Inches(0.1)
    
    p = tf_top.paragraphs[0]
    r = p.add_run()
    r.text = "CORE TECHNICAL & ENGINEERING OBJECTIVES"
    set_run_font(r, 12.5, bold=True, color=NAVY_PRIMARY)
    
    objs = [
        ("Automated Spatial Ingestion & Standardized Indexing",
         "Ingest, sanitize, and index 760 Indian district and state boundaries in standard WGS 84 (EPSG:4326) with automated geometry repair."),
        ("Sub-Second Synchronous Multi-Hazard Engine",
         "Implement coordinate-level point analysis within 0.27s–0.80s, decoding AWS Terrarium DEM tiles and streaming USGS & Open-Meteo feeds."),
        ("Concurrent 8-Vector Hazard Computation",
         "Simultaneously evaluate 8 hazard vectors (Flood, Inundation, Quake, Cyclone, Drought, Heat, Landslide, Fire) normalized to [0.0, 1.0]."),
        ("Resilient Lifeline Discovery with Automatic Failover",
         "Deploy an asynchronous circuit-breaker query across OpenStreetMap Overpass API mirror pools, achieving 99.5% operational uptime."),
        ("Evidence-Based Decision Support & Zero-Synthetic Integrity",
         "Translate multi-hazard scorecards into 13 prioritized mitigation advisories with strict data provenance and zero fake substitutions.")
    ]
    for idx, (title, desc) in enumerate(objs):
        po = tf_top.add_paragraph()
        po.space_before = Pt(3)
        ro_num = po.add_run()
        ro_num.text = f"{idx+1}. {title}: "
        set_run_font(ro_num, 10, bold=True, color=BLUE_ACCENT)
        ro_desc = po.add_run()
        ro_desc.text = desc
        set_run_font(ro_desc, 9.5, color=TEXT_DARK)
        
    # Reposition Picture 7 & 8 cleanly
    for s in s7.shapes:
        if s.name == "Picture 7":
            s.left = Inches(0.9)
            s.top = Inches(4.7)
            s.width = Inches(2.2)
            s.height = Inches(1.7)
        elif s.name == "Picture 8":
            s.left = Inches(3.3)
            s.top = Inches(4.7)
            s.width = Inches(2.1)
            s.height = Inches(1.9)

    # Left SDG Header text box
    sdg_head = s7.shapes.add_textbox(Inches(0.76), Inches(4.3), Inches(4.8), Inches(0.35))
    tf_sh = sdg_head.text_frame
    tf_sh.margin_left = Inches(0.05)
    p_sh = tf_sh.paragraphs[0]
    r_sh = p_sh.add_run()
    r_sh.text = "Identified Sustainable Development Goals (SDGs):"
    set_run_font(r_sh, 11, bold=True, color=NAVY_PRIMARY)

    # Right Card: Target Project Outcomes
    card_out = create_card(s7, Inches(5.6), Inches(4.35), Inches(6.96), Inches(2.35))
    tf_out = card_out.text_frame
    tf_out.word_wrap = True
    tf_out.margin_left = Inches(0.2)
    tf_out.margin_right = Inches(0.2)
    tf_out.margin_top = Inches(0.1)
    
    p = tf_out.paragraphs[0]
    r = p.add_run()
    r.text = "TARGET PROJECT OUTCOMES & SOCIETAL IMPACT"
    set_run_font(r, 11.5, bold=True, color=NAVY_PRIMARY)
    
    outcomes = [
        ("Target Scopus Publication", "Manuscript targeted for Scopus-Indexed International Conference / Journal (Disaster Geoinformatics / Remote Sensing)."),
        ("Software Product Deliverable", "Open-source, enterprise-grade Web-GIS multi-hazard platform (INDIGEO-GIS v1.0.0, MIT License)."),
        ("Societal & Humanitarian Impact", "Zero-install, high-resiliency decision-support tool serving district authorities, panchayats, and field responders."),
        ("Academic & Engineering Rigor", "Complete test suite (34 test modules), 100% data provenance, and zero synthetic data fabrication.")
    ]
    for o_title, o_desc in outcomes:
        p_o = tf_out.add_paragraph()
        p_o.space_before = Pt(3)
        ro_t = p_o.add_run()
        ro_t.text = f"• {o_title}: "
        set_run_font(ro_t, 9.5, bold=True, color=BLUE_ACCENT)
        ro_d = p_o.add_run()
        ro_d.text = o_desc
        set_run_font(ro_d, 9, color=TEXT_DARK)

    # =============================================================
    # SLIDE 8: FRAMEWORK / ARCHITECTURE / BLOCK DIAGRAM
    # =============================================================
    print("Formatting Slide 8 (Architecture)...")
    s8 = prs.slides[7]
    remove_shape(s8, "Content Placeholder 2")
    
    # Left Side: 7-Layer Architecture Description
    card_arch = create_card(s8, Inches(0.76), Inches(1.5), Inches(5.4), Inches(4.7))
    tf_arch = card_arch.text_frame
    tf_arch.word_wrap = True
    tf_arch.margin_left = Inches(0.15)
    tf_arch.margin_right = Inches(0.15)
    tf_arch.margin_top = Inches(0.1)
    
    p = tf_arch.paragraphs[0]
    r = p.add_run()
    r.text = "7-LAYER SYSTEM ARCHITECTURE"
    set_run_font(r, 12.5, bold=True, color=NAVY_PRIMARY)
    
    layers = [
        ("Layer 1: Presentation Layer", "Next.js 16 + React 19 + Leaflet thin-client web dashboard; zero client fallback values."),
        ("Layer 2: API Gateway Layer", "Python FastAPI asynchronous router, CORS, Pydantic schema validation, and health probes."),
        ("Layer 3: Spatial Ingestion Layer", "GeoPandas, GDAL, Shapely; standardizes 760 districts in WGS 84 (EPSG:4326) with sub-40ms PIP."),
        ("Layer 4: Telemetry Streams", "Real-time AWS Terrarium DEM (RGB decoding), USGS ComCat (Seismic), Open-Meteo ERA5 (Atmospheric)."),
        ("Layer 5: Multi-Hazard Core", "Concurrent 8-vector deterministic normalization engine [0.0, 1.0] across environmental thresholds."),
        ("Layer 6: Lifelines & Failover", "Async Overpass API cluster querying nearest hospitals, shelters, fire stations with 99.5% uptime."),
        ("Layer 7: Decision Support Layer", "13 evidence-based preparedness advisories + roadmap for Google OR-Tools CVRP & NetworkX routing.")
    ]
    for l_title, l_desc in layers:
        pl = tf_arch.add_paragraph()
        pl.space_before = Pt(4)
        rl_t = pl.add_run()
        rl_t.text = f"• {l_title}: "
        set_run_font(rl_t, 10, bold=True, color=BLUE_ACCENT)
        rl_d = pl.add_run()
        rl_d.text = l_desc
        set_run_font(rl_d, 9.5, color=TEXT_DARK)
        
    # Right Side: Embed Architecture Diagram (img_2.png)
    img_arch_path = os.path.join("temp_docx_pics", "img_2.png")
    if os.path.exists(img_arch_path):
        img_left = Inches(6.35)
        img_top = Inches(1.5)
        img_width = Inches(6.2)
        img_height = Inches(4.5)
        pic = s8.shapes.add_picture(img_arch_path, img_left, img_top, img_width, img_height)
        
        tb_cap = s8.shapes.add_textbox(img_left, Inches(6.05), img_width, Inches(0.3))
        tf_cap = tb_cap.text_frame
        p_c = tf_cap.paragraphs[0]
        p_c.alignment = PP_ALIGN.CENTER
        r_c = p_c.add_run()
        r_c.text = "Figure 4.1: INDIGEO-GIS Multi-Tier Client-Server Architecture (Report Page 30)"
        set_run_font(r_c, 9.5, bold=True, italic=True, color=TEXT_MUTED)

    # =============================================================
    # SLIDE 9: FUNCTIONAL REQUIREMENTS
    # =============================================================
    print("Formatting Slide 9 (Functional Requirements)...")
    s9 = prs.slides[8]
    remove_shape(s9, "Content Placeholder 2")
    
    t_top = Inches(1.5)
    t_width = Inches(11.8)
    t_height = Inches(4.65)
    tbl_fr_shape = s9.shapes.add_table(11, 4, Inches(0.76), t_top, t_width, t_height)
    tbl_fr = tbl_fr_shape.table
    
    col_w = [Inches(1.2), Inches(2.6), Inches(2.8), Inches(5.2)]
    for idx, w in enumerate(col_w):
        tbl_fr.columns[idx].width = w
        
    fr_headers = ["Req ID", "Functional Module", "Input Parameters", "System Processing & Expected Output"]
    for col_idx, h in enumerate(fr_headers):
        format_cell(tbl_fr.cell(0, col_idx), h, font_size=10, bold=True, text_color=WHITE, bg_color=NAVY_PRIMARY, align=PP_ALIGN.CENTER)
        
    fr_data = [
        ("FR-01", "Boundary Search", "District / State Name", "Fast case-insensitive spatial index lookup across 760 boundaries; highlights polygon (<40ms)."),
        ("FR-02", "Point Query", "Latitude, Longitude (EPSG:4326)", "Ray-Casting point-in-polygon verification; returns parent district identity and Census code."),
        ("FR-03", "DEM Elevation & Slope", "Coordinate Bounding Box", "Decodes AWS Terrarium RGB tiles: Elev = (R*256 + G + B/256) - 32768; calculates slope & depression."),
        ("FR-04", "Seismic Assessment", "Coordinates & Search Radius", "Queries live USGS ComCat API for M >= 4.0 events (30 days); computes distance and focal depth."),
        ("FR-05", "Climate Analytics", "Point Coordinates", "Fetches Open-Meteo ERA5 reanalysis & 14-day rainfall forecast; computes soil moisture & deficiency."),
        ("FR-06", "Multi-Hazard Profiling", "Synthesized Telemetry Streams", "Simultaneously evaluates 8 hazard vectors using deterministic formulas; outputs scores [0.0 - 1.0]."),
        ("FR-07", "Lifeline Discovery", "Coordinates & Radius (km)", "Async query to OSM Overpass pool for nearest hospitals, fire stations, shelters, and roads."),
        ("FR-08", "Automatic Failover", "Overpass Timeout / HTTP 5xx", "Dynamic round-robin failover to healthy secondary mirror (e.g. maps.mail.ru) without user delay."),
        ("FR-09", "Preparedness Advisory", "Multi-Hazard Scorecard", "Rule engine mapping risk triggers to 13 specific mitigation actions with agency and urgency flags."),
        ("FR-10", "Provenance Verification", "External Integration Feeds", "Validates credential states; explicitly flags unconfigured APIs as 'unverified' with zero fake data.")
    ]
    
    for row_idx, r_data in enumerate(fr_data):
        bg = ROW_ALT_FILL if row_idx % 2 == 1 else WHITE
        for col_idx, val in enumerate(r_data):
            align = PP_ALIGN.CENTER if col_idx in [0, 1] else PP_ALIGN.LEFT
            bold = (col_idx == 0)
            format_cell(tbl_fr.cell(row_idx+1, col_idx), val, font_size=9, bold=bold, text_color=TEXT_DARK, bg_color=bg, align=align)

    # =============================================================
    # SLIDE 10: MODULES
    # =============================================================
    print("Formatting Slide 10 (Modules)...")
    s10 = prs.slides[9]
    remove_shape(s10, "Content Placeholder 2")
    
    modules = [
        ("Module 1: Administrative GeoJSON Ingestion & Indexing",
         "backend/app/gis/coords.py, spatial.py",
         "• Standardizes 760 Indian district and state boundaries in WGS 84 (EPSG:4326).\n• Performs in-memory Shapely Ray-Casting point-in-polygon queries in <0.04s.\n• Sanitizes corrupted boundary geometries and eliminates relational schema crashes."),
        ("Module 2: Remote Sensing & Terrain Analysis Engine",
         "backend/app/gis/dem.py, local_dem.py",
         "• Real-time decoding of AWS Terrarium RGB PNG elevation tiles in 0.15s.\n• Mathematical formula: Elevation = (R × 256 + G + B / 256) - 32768 meters.\n• Calculates terrain slope gradient (°) and relative topographic depression."),
        ("Module 3: Concurrent Multi-Hazard Computation Core",
         "backend/app/services/hazard.py, risk.py",
         "• Computes 8 hazard vectors simultaneously: Flood, Coastal, Quake, Cyclone, Drought, Heat, Landslide, Fire.\n• Normalizes all outputs to continuous [0.0, 1.0] risk scores based on scientific physical triggers.\n• Resolves Base Paper 2's single-hazard limitation."),
        ("Module 4: Resilient Lifeline Infrastructure Discovery",
         "backend/app/integrations/osm_infra.py, services/lifelines.py",
         "• Asynchronous circuit-breaker querying OpenStreetMap Overpass mirror cluster (DE, Mail.ru, Kumi).\n• Automatically switches mirrors upon timeout, achieving 99.5% operational uptime.\n• Identifies nearest hospitals, shelters, fire stations, and critical road segments in 1.42s."),
        ("Module 5: Actionable Preparedness Advisory Engine",
         "backend/app/services/recommendations.py, evidence.py",
         "• Translates multi-hazard risk scorecards into 13 evidence-based mitigation actions.\n• Assigns operational urgency levels (Critical, High, Moderate) and target agencies (NDRF, SDMA, PWD).\n• Bridges passive hazard rendering with active field emergency operations."),
        ("Module 6: GeoAI Machine Learning & Evacuation Roadmap",
         "backend/app/ml/, Base Paper 2 Extension",
         "• Blueprint for XGBoost & LightGBM training on 20-year historical disaster records (NDMA/IMD).\n• Cloud GPU Sentinel-1 SAR U-Net deep learning for all-weather radar flood segmentation.\n• TreeSHAP local explainability popups and Google OR-Tools CVRP evacuation routing.")
    ]
    
    card_w = Inches(5.75)
    card_h = Inches(1.42)
    for idx, (title, files, desc) in enumerate(modules):
        col = idx % 2
        row = idx // 2
        l = Inches(0.76) if col == 0 else Inches(6.8)
        t = Inches(1.55) + row * (card_h + Inches(0.15))
        
        card = create_card(s10, l, t, card_w, card_h)
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.15)
        tf.margin_right = Inches(0.15)
        tf.margin_top = Inches(0.08)
        tf.margin_bottom = Inches(0.08)
        
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = title
        set_run_font(r, 11, bold=True, color=NAVY_PRIMARY)
        
        pf = tf.add_paragraph()
        rf = pf.add_run()
        rf.text = f"Files: {files}"
        set_run_font(rf, 9, italic=True, color=BLUE_ACCENT)
        
        pd = tf.add_paragraph()
        pd.space_before = Pt(2)
        rd = pd.add_run()
        rd.text = desc
        set_run_font(rd, 9, color=TEXT_DARK)

    # =============================================================
    # SLIDE 11: EXPERIMENTS AND RESULTS
    # =============================================================
    print("Formatting Slide 11 (Experiments and Results)...")
    s11 = prs.slides[10]
    remove_shape(s11, "Content Placeholder 2")
    
    # Left Side: Benchmarking Table (Table 7.1 from report)
    t_top = Inches(1.5)
    t_width = Inches(6.4)
    t_height = Inches(3.2)
    tbl_res_shape = s11.shapes.add_table(7, 4, Inches(0.76), t_top, t_width, t_height)
    tbl_res = tbl_res_shape.table
    
    col_w_res = [Inches(2.5), Inches(1.3), Inches(1.3), Inches(1.3)]
    for idx, w in enumerate(col_w_res):
        tbl_res.columns[idx].width = w
        
    res_headers = ["Geospatial Operation", "Legacy GIS", "INDIGEO", "Speedup"]
    for col_idx, h in enumerate(res_headers):
        format_cell(tbl_res.cell(0, col_idx), h, font_size=10, bold=True, text_color=WHITE, bg_color=NAVY_PRIMARY, align=PP_ALIGN.CENTER)
        
    res_data = [
        ("Point-in-Polygon Lookup", "1.85 s", "0.04 s", "46x Faster"),
        ("DEM Elevation & Slope Extraction", "4.20 s", "0.15 s", "28x Faster"),
        ("Atmospheric Telemetry Ingestion", "3.50 s", "0.22 s", "16x Faster"),
        ("Seismic Hazard Buffer Scan", "2.10 s", "0.34 s", "6x Faster"),
        ("Lifeline Mirror Failover Search", "18.50 s", "1.42 s", "13x Faster"),
        ("End-to-End Synchronous Query", "11.30 s", "0.54 s", "21x - 150x")
    ]
    for row_idx, r_data in enumerate(res_data):
        is_total = (row_idx == 5)
        bg = RGBColor(230, 245, 235) if is_total else (ROW_ALT_FILL if row_idx % 2 == 1 else WHITE)
        t_col = GREEN_DARK if is_total else TEXT_DARK
        for col_idx, val in enumerate(r_data):
            bold = is_total or (col_idx == 3)
            col_color = GREEN_DARK if col_idx == 3 else t_col
            align = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.CENTER
            format_cell(tbl_res.cell(row_idx+1, col_idx), val, font_size=9.5, bold=bold, text_color=col_color, bg_color=bg, align=align)

    # Key Empirical Insights Card below table
    card_ins = create_card(s11, Inches(0.76), Inches(4.85), Inches(6.4), Inches(1.35))
    tf_ins = card_ins.text_frame
    tf_ins.word_wrap = True
    tf_ins.margin_left = Inches(0.15)
    tf_ins.margin_right = Inches(0.15)
    tf_ins.margin_top = Inches(0.08)
    
    p = tf_ins.paragraphs[0]
    r = p.add_run()
    r.text = "EMPIRICAL VALIDATION & DATA INTEGRITY"
    set_run_font(r, 11, bold=True, color=NAVY_PRIMARY)
    
    p_b1 = tf_ins.add_paragraph()
    p_b1.space_before = Pt(2)
    rb1 = p_b1.add_run()
    rb1.text = "• Rigorous Stress Testing: Benchmarked across 1,000 randomized coordinate queries across all physiographic zones of India; observed query range: 0.27s – 0.80s.\n• Zero-Synthetic Data Integrity: Missing credentials (e.g. NASA FIRMS) strictly return 'unverified' state; 0% synthetic or fabricated risk scores generated.\n• Test Suite Reliability: 34 comprehensive unit & integration test suites passing in automated CI/CD."
    set_run_font(rb1, 9, color=TEXT_DARK)

    # Right Side: Embed Latency Bar Chart (img_11.png)
    img_lat_path = os.path.join("temp_docx_pics", "img_11.png")
    if os.path.exists(img_lat_path):
        img_left = Inches(7.35)
        img_top = Inches(1.5)
        img_width = Inches(5.2)
        img_height = Inches(4.3)
        pic = s11.shapes.add_picture(img_lat_path, img_left, img_top, img_width, img_height)
        
        tb_cap = s11.shapes.add_textbox(img_left, Inches(5.85), img_width, Inches(0.35))
        tf_cap = tb_cap.text_frame
        p_c = tf_cap.paragraphs[0]
        p_c.alignment = PP_ALIGN.CENTER
        r_c = p_c.add_run()
        r_c.text = "Figure 7.1: End-to-End Latency: Legacy GIS vs INDIGEO-GIS (Report Page 44)"
        set_run_font(r_c, 9.5, bold=True, italic=True, color=TEXT_MUTED)

    # =============================================================
    # SLIDE 12: CONCLUSION AND FUTURE WORK
    # =============================================================
    print("Formatting Slide 12 (Conclusion and Future Work)...")
    s12 = prs.slides[11]
    remove_shape(s12, "Content Placeholder 2")
    
    # Left Card: Summary of Contributions
    card_c = create_card(s12, Inches(0.76), Inches(1.55), Inches(5.75), Inches(4.65))
    tf_c = card_c.text_frame
    tf_c.word_wrap = True
    tf_c.margin_left = Inches(0.2)
    tf_c.margin_right = Inches(0.2)
    tf_c.margin_top = Inches(0.15)
    
    p = tf_c.paragraphs[0]
    r = p.add_run()
    r.text = "SUMMARY OF CONTRIBUTIONS"
    set_run_font(r, 13, bold=True, color=NAVY_PRIMARY)
    
    c_points = [
        ("Unified Spatial Ingestion", "Resolved the relational schema crash barrier by engineering an automated GeoPandas/Shapely ingestion layer indexing 760 Indian administrative boundaries in EPSG:4326."),
        ("Sub-Second Synchronous Speed", "Demonstrated 0.54s average query response (20x–150x faster than legacy Web-GIS and offline research notebooks) via asynchronous FastAPI orchestration."),
        ("Concurrent Multi-Hazard Intelligence", "Overcame single-hazard research silos by evaluating 8 hazard vectors simultaneously with normalized continuous risk scores in [0.0, 1.0]."),
        ("Resilient Lifeline Mapping", "Achieved 99.5% operational uptime for discovering emergency infrastructure using an async multi-mirror Overpass circuit-breaker network."),
        ("Actionable Decision Support & TCO Savings", "Generated 13 evidence-based preparedness advisories with strict data provenance, eliminating $285k USD in 5-year commercial licensing fees.")
    ]
    for title, desc in c_points:
        pc = tf_c.add_paragraph()
        pc.space_before = Pt(5)
        rc_t = pc.add_run()
        rc_t.text = f"• {title}: "
        set_run_font(rc_t, 10.5, bold=True, color=BLUE_ACCENT)
        rc_d = pc.add_run()
        rc_d.text = desc
        set_run_font(rc_d, 10, color=TEXT_DARK)

    # Right Card: Future Scope and GeoAI Engineering Roadmap
    card_f = create_card(s12, Inches(6.8), Inches(1.55), Inches(5.75), Inches(4.65))
    tf_f = card_f.text_frame
    tf_f.word_wrap = True
    tf_f.margin_left = Inches(0.2)
    tf_f.margin_right = Inches(0.2)
    tf_f.margin_top = Inches(0.15)
    
    p = tf_f.paragraphs[0]
    r = p.add_run()
    r.text = "FUTURE SCOPE & GeoAI ENGINEERING ROADMAP"
    set_run_font(r, 13, bold=True, color=NAVY_PRIMARY)
    
    f_points = [
        ("Machine Learning Multi-Hazard Training", "Train gradient boosted decision trees (XGBoost, LightGBM) on 20-year NDMA and IMD historical disaster datasets to learn non-linear multi-hazard susceptibility."),
        ("Satellite Radar SAR U-Net Deep Learning", "Deploy lightweight U-Net convolutional neural network on cloud GPUs to segment Sentinel-1 SAR imagery for all-weather flood boundary detection through monsoon clouds."),
        ("Live Explainable AI (SHAP) Dashboard", "Integrate TreeSHAP directly inside the FastAPI backend to generate live interactive waterfall attribution plots in the Leaflet map popup cards."),
        ("Algorithmic Evacuation Logistics", "Implement Capacitated Vehicle Routing Problem (CVRP) via Google OR-Tools for shelter allocation, coupled with NetworkX Dijkstra dynamic routing to bypass flooded road segments.")
    ]
    for title, desc in f_points:
        pf = tf_f.add_paragraph()
        pf.space_before = Pt(7)
        rf_t = pf.add_run()
        rf_t.text = f"• {title}: "
        set_run_font(rf_t, 10.5, bold=True, color=GREEN_DARK)
        rf_d = pf.add_run()
        rf_d.text = desc
        set_run_font(rf_d, 10, color=TEXT_DARK)

    # =============================================================
    # SLIDE 13: OUTCOME ACHIEVED (LEAVE EXACTLY AS IS!)
    # =============================================================
    print("Slide 13 left intact as in template (Outcome Achieved).")

    # =============================================================
    # SLIDE 14: REFERENCES
    # =============================================================
    print("Formatting Slide 14 (References)...")
    s14 = prs.slides[13]
    remove_shape(s14, "Rectangle 3")
    
    card_ref = create_card(s14, Inches(0.76), Inches(1.5), Inches(11.8), Inches(4.75))
    tf_ref = card_ref.text_frame
    tf_ref.word_wrap = True
    tf_ref.margin_left = Inches(0.2)
    tf_ref.margin_right = Inches(0.2)
    tf_ref.margin_top = Inches(0.12)
    
    refs = [
        ("[1]", "A. F. Al-Hussein, M. R. Karim, and S. M. Hassan, \"Comprehensive Analysis of the Use of Web-GIS for Natural Hazard Management: A Systematic Review,\" Sustainability, vol. 16, no. 10, p. 4238, May 2024. doi: 10.3390/su16104238. [Base Paper 1]"),
        ("[2]", "Z. Chen, Y. Liang, and X. Zhang, \"From Data to Decision: Interpretable Machine Learning for Predicting Flood Susceptibility,\" Remote Sensing, vol. 16, no. 20, p. 3902, Oct. 2024. doi: 10.3390/rs16203902. [Base Paper 2]"),
        ("[3]", "S. M. Lundberg and S.-I. Lee, \"A Unified Approach to Interpreting Model Predictions,\" in Advances in Neural Information Processing Systems (NeurIPS 2017), vol. 30, Long Beach, CA, USA, Dec. 2017, pp. 4765–4774."),
        ("[4]", "T. Chen and C. Guestrin, \"XGBoost: A Scalable Tree Boosting System,\" in Proc. 22nd ACM SIGKDD Int. Conf. on Knowledge Discovery and Data Mining (KDD '16), San Francisco, CA, USA, Aug. 2016, pp. 785–794. doi: 10.1145/2939672.2939785."),
        ("[5]", "T. Bandrova, S. Zlatanova, and M. Konecny, \"Geoinformatics for Disaster Management: Advances and Challenges,\" ISPRS International Journal of Geo-Information, vol. 7, no. 5, p. 184, May 2018. doi: 10.3390/ijgi7050184."),
        ("[6]", "H. Hersbach et al., \"The ERA5 Global Reanalysis,\" Quarterly Journal of the Royal Meteorological Society, vol. 146, no. 730, pp. 1999–2049, Jul. 2020. doi: 10.1002/qj.3803."),
        ("[7]", "S. R. Earle, B. J. Walker, and J. C. Dewey, \"USGS Comprehensive Earthquake Catalog (ComCat): Real-Time Global Earthquake Information Architecture,\" Seismological Research Letters, vol. 88, no. 5, pp. 1234–1246, Sep. 2017. doi: 10.1785/0220170045."),
        ("[8]", "M. Haklay and P. Weber, \"OpenStreetMap: User-Generated Street Maps,\" IEEE Pervasive Computing, vol. 7, no. 4, pp. 12–18, Oct.-Dec. 2008. doi: 10.1109/MPRV.2008.80.")
    ]
    
    for idx, (num, text) in enumerate(refs):
        pr = tf_ref.paragraphs[0] if idx == 0 else tf_ref.add_paragraph()
        if idx > 0:
            pr.space_before = Pt(5)
        rn = pr.add_run()
        rn.text = f"{num} "
        set_run_font(rn, 10, bold=True, color=BLUE_ACCENT)
        rt = pr.add_run()
        rt.text = text
        set_run_font(rt, 9.5, color=TEXT_DARK)

    # Save presentation
    print(f"Saving final presentation to: {OUTPUT_PATH}")
    prs.save(OUTPUT_PATH)
    print("SUCCESS: priyanshu ppt.pptx generated successfully!")

if __name__ == "__main__":
    main()
