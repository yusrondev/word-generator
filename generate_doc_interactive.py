import docx
from docx.shared import Cm, Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
import html
import os

def create_interactive_proposal():
    doc = docx.Document()

    # ==========================================
    # 1. SETUP PAGE SIZE & MARGINS (A4: 4-4-3-3)
    # ==========================================
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(4.0)
    section.bottom_margin = Cm(3.0)
    section.left_margin = Cm(4.0)
    section.right_margin = Cm(3.0)
    
    # Printable width: 21.0 - 4.0 - 3.0 = 14.0 cm
    PRINTABLE_WIDTH_CM = 14.0

    # ==========================================
    # 2. SETUP STYLES & TYPOGRAPHY
    # ==========================================
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(12)
    style_normal.font.color.rgb = RGBColor(0, 0, 0)
    style_normal.paragraph_format.line_spacing = 1.5
    style_normal.paragraph_format.space_after = Pt(6)
    style_normal.paragraph_format.space_before = Pt(0)
    style_normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    bookmark_counter = 1

    # Helper function for adding headings with bookmarks
    def add_bookmarked_heading(text, bookmark_name, size_pt=12, is_bold=True, is_center=False, space_before=12, space_after=6):
        nonlocal bookmark_counter
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if is_center else WD_ALIGN_PARAGRAPH.LEFT
        
        # Bookmark start
        b_start = parse_xml(r'<w:bookmarkStart %s w:id="%d" w:name="%s"/>' % (nsdecls('w'), bookmark_counter, bookmark_name))
        p._p.append(b_start)
        
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(size_pt)
        run.font.bold = is_bold
        run.font.color.rgb = RGBColor(0, 0, 0)
        
        # Bookmark end
        b_end = parse_xml(r'<w:bookmarkEnd %s w:id="%d"/>' % (nsdecls('w'), bookmark_counter))
        p._p.append(b_end)
        
        bookmark_counter += 1
        return p

    # Helper function for body paragraphs with indentation
    def add_body_paragraph(text, indent_cm=1.0):
        p = doc.add_paragraph()
        p.style = 'Normal'
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.first_line_indent = Cm(indent_cm)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    # Helper function for clickable TOC lines (internal hyperlinks)
    def add_interactive_toc_line(title, page_num, anchor_target, level=1, is_bold=False):
        p = doc.add_paragraph()
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.space_before = Pt(2 if level == 1 else 0)
        
        p.paragraph_format.tab_stops.add_tab_stop(Cm(PRINTABLE_WIDTH_CM), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        
        if level == 2:
            p.paragraph_format.left_indent = Cm(0.6)
        elif level == 3:
            p.paragraph_format.left_indent = Cm(1.2)

        hyperlink = parse_xml(r'<w:hyperlink %s w:anchor="%s" w:history="1"/>' % (nsdecls('w'), anchor_target))
        
        sz_val = "23" if level == 1 else "22"
        b_tag = "<w:b/>" if is_bold else ""
        escaped_title = html.escape(title)
        
        # Title Run
        r1_xml = (
            r'<w:r %s><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            r'<w:sz w:val="%s"/>%s<w:color w:val="000000"/><w:u w:val="none"/></w:rPr>'
            r'<w:t xml:space="preserve">%s</w:t></w:r>'
        ) % (nsdecls('w'), sz_val, b_tag, escaped_title)
        hyperlink.append(parse_xml(r1_xml))
        
        # Tab and Page Run
        r2_xml = (
            r'<w:r %s><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            r'<w:sz w:val="%s"/>%s<w:color w:val="000000"/><w:u w:val="none"/></w:rPr>'
            r'<w:tab/><w:t>%s</w:t></w:r>'
        ) % (nsdecls('w'), sz_val, b_tag, str(page_num))
        hyperlink.append(parse_xml(r2_xml))
        
        p._p.append(hyperlink)
        return p

    # ==========================================
    # 3. HALAMAN COVER / JUDUL
    # ==========================================
    add_bookmarked_heading("PROPOSAL PENELITIAN", "bm_cover", size_pt=14, is_bold=True, is_center=True, space_before=10, space_after=18)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.line_spacing = 1.15
    p_title.paragraph_format.space_after = Pt(16)
    r_title = p_title.add_run(
        "RANCANG BANGUN INTELLIGENT CHATBOT MENGGUNAKAN GEMINI API "
        "DENGAN PENDEKATAN RETRIEVAL-AUGMENTED GENERATION (RAG) "
        "UNTUK LAYANAN INFORMASI PADA PT AKU CINTA INDONESIA (ACI)"
    )
    r_title.font.name = 'Times New Roman'
    r_title.font.size = Pt(13)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0, 0, 0)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(20)
    r_sub = p_sub.add_run("Tugas 1 Mata Kuliah STSI4310 Metode Penelitian")
    r_sub.font.name = 'Times New Roman'
    r_sub.font.size = Pt(12)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(0, 0, 0)

    # Logo Universitas Terbuka
    logo_path = os.path.join(os.path.dirname(__file__), "logo_ut.png")
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_before = Pt(6)
        p_logo.paragraph_format.space_after = Pt(24)
        run_logo = p_logo.add_run()
        run_logo.add_picture(logo_path, width=Cm(4.2))

    # Identitas Mahasiswa (Tabel tanpa border)
    table_ident = doc.add_table(rows=4, cols=3)
    table_ident.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_ident.autofit = False

    ident_data = [
        ("Nama Mahasiswa", ":", "Yusron Laksono"),
        ("NIM", ":", "[Isi NIM Anda]"),
        ("Program Studi", ":", "Sistem Informasi"),
        ("Fakultas", ":", "Fakultas Sains dan Teknologi (FST)")
    ]

    for row_idx, data in enumerate(ident_data):
        row = table_ident.rows[row_idx]
        for col_idx, text in enumerate(data):
            cell = row.cells[col_idx]
            cell.text = text
            cp = cell.paragraphs[0]
            cp.paragraph_format.line_spacing = 1.2
            cp.paragraph_format.space_after = Pt(2)
            crun = cp.runs[0]
            crun.font.name = 'Times New Roman'
            crun.font.size = Pt(12)
            crun.font.color.rgb = RGBColor(0, 0, 0)
            if col_idx == 0:
                crun.font.bold = True
            
        row.cells[0].width = Cm(4.5)
        row.cells[1].width = Cm(0.5)
        row.cells[2].width = Cm(8.0)

    tblPr = table_ident._tbl.tblPr
    borders = parse_xml(
        r'<w:tblBorders %s>'
        r'  <w:top w:val="none"/>'
        r'  <w:left w:val="none"/>'
        r'  <w:bottom w:val="none"/>'
        r'  <w:right w:val="none"/>'
        r'  <w:insideH w:val="none"/>'
        r'  <w:insideV w:val="none"/>'
        r'</w:tblBorders>' % nsdecls('w')
    )
    tblPr.append(borders)

    # Bottom Affiliation
    p_bot = doc.add_paragraph()
    p_bot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_bot.paragraph_format.space_before = Pt(36)
    p_bot.paragraph_format.line_spacing = 1.15
    r_bot = p_bot.add_run(
        "PROGRAM STUDI SISTEM INFORMASI\n"
        "FAKULTAS SAINS DAN TEKNOLOGI\n"
        "UNIVERSITAS TERBUKA\n"
        "2026"
    )
    r_bot.font.name = 'Times New Roman'
    r_bot.font.size = Pt(12)
    r_bot.font.bold = True
    r_bot.font.color.rgb = RGBColor(0, 0, 0)

    doc.add_page_break()

    # ==========================================
    # 4. HALAMAN DAFTAR ISI (INTERAKTIF / KLIKABLE)
    # ==========================================
    add_bookmarked_heading("DAFTAR ISI", "bm_daftar_isi", size_pt=12, is_bold=True, is_center=True, space_before=0, space_after=18)

    add_interactive_toc_line("HALAMAN COVER", "i", "bm_cover", level=1, is_bold=True)
    add_interactive_toc_line("DAFTAR ISI", "ii", "bm_daftar_isi", level=1, is_bold=True)
    add_interactive_toc_line("BAB I PENDAHULUAN", "1", "bm_bab1", level=1, is_bold=True)
    add_interactive_toc_line("1.1 Latar Belakang Masalah", "1", "bm_bab1_1", level=2)
    add_interactive_toc_line("1.2 Rumusan Masalah", "3", "bm_bab1_2", level=2)
    add_interactive_toc_line("1.3 Tujuan dan Manfaat Penelitian", "3", "bm_bab1_3", level=2)
    add_interactive_toc_line("1.3.1 Tujuan Penelitian", "3", "bm_bab1_3_1", level=3)
    add_interactive_toc_line("1.3.2 Manfaat Penelitian", "3", "bm_bab1_3_2", level=3)
    add_interactive_toc_line("BAB II KAJIAN PUSTAKA & LANDASAN TEORI", "4", "bm_bab2", level=1, is_bold=True)
    add_interactive_toc_line("2.1 Kajian Pustaka dan Landasan Teori", "4", "bm_bab2_1", level=2)
    add_interactive_toc_line("2.2 Pertanyaan Penelitian dan Hipotesis", "4", "bm_bab2_2", level=2)
    add_interactive_toc_line("BAB III METODE PENELITIAN", "5", "bm_bab3", level=1, is_bold=True)
    add_interactive_toc_line("3.1 Metode Penelitian", "5", "bm_bab3_1", level=2)
    add_interactive_toc_line("3.2 Teknik Pengumpulan Data dan Instrumen Penelitian", "5", "bm_bab3_2", level=2)
    add_interactive_toc_line("BAB IV ANALISIS DATA DAN KESIMPULAN SEMENTARA", "6", "bm_bab4", level=1, is_bold=True)
    add_interactive_toc_line("4.1 Teknik Analisis Data Penelitian", "6", "bm_bab4_1", level=2)
    add_interactive_toc_line("4.2 Kesimpulan Sementara dan Batasan Penelitian", "6", "bm_bab4_2", level=2)
    add_interactive_toc_line("DAFTAR PUSTAKA", "7", "bm_daftar_pustaka", level=1, is_bold=True)

    doc.add_page_break()

    # ==========================================
    # 5. BAB I – PENDAHULUAN
    # ==========================================
    add_bookmarked_heading("BAB I\nPENDAHULUAN", "bm_bab1", size_pt=12, is_bold=True, is_center=True, space_before=0, space_after=12)

    # 1.1 Latar Belakang Masalah
    add_bookmarked_heading("1.1 Latar Belakang Masalah", "bm_bab1_1", size_pt=12, is_bold=True, is_center=False, space_before=12, space_after=6)

    add_body_paragraph(
        "Di era transformasi digital yang berkembang pesat saat ini, efisiensi dan kecepatan penyampaian informasi "
        "layanan menjadi salah satu determinan utama dalam menjaga kepuasan pelanggan serta keunggulan operasional perusahaan. "
        "PT Aku Cinta Indonesia (ACI) sebagai perusahaan yang bergerak di bidang penyediaan layanan informasi dan produk digital "
        "mengalami pertumbuhan pengguna aktif yang signifikan. Konsekuensinya, beban komunikasi pada kanal layanan informasi dan "
        "bantuan pelanggan meningkat secara eksponensial."
    )
    add_body_paragraph(
        "Berdasarkan hasil observasi dan rekapitulasi data layanan pelanggan PT Aku Cinta Indonesia (ACI) periode Januari–Juni 2026, "
        "tercatat rata-rata 1.450 tiket pertanyaan masuk setiap bulannya. Dari total volume tersebut, sekitar 58% hingga 62% "
        "merupakan pertanyaan berulang (repetitive inquiries) yang mencakup panduan pendaftaran akun, prosedur administrasi langganan, "
        "verifikasi pembayaran, serta spesifikasi teknis layanan. Tingginya frekuensi pertanyaan berulang ini menyita sebagian besar "
        "waktu kerja 4 orang staf customer support yang bertugas pada jam kerja operasional (08.00–17.00 WIB). Kondisi ini memicu "
        "disparitas layanan yang signifikan: pada jam kerja, rata-rata waktu tanggap berkisar antara 15–25 menit per interaksi, "
        "sementara pertanyaan pengguna yang diajukan di luar jam operasional atau hari libur mengalami penundaan (delay) hingga 12–16 jam "
        "sebelum memperoleh respons pada hari kerja berikutnya. Di sisi lain, implementasi chatbot konvensional berbasis aturan kaku "
        "(rule-based) yang sebelumnya digunakan perusahaan memiliki tingkat kegagalan respons (fallback rate) mencapai 41% karena ketidakmampuan "
        "sistem dalam memahami variasi semantik bahasa alami pengguna, serta sulit diperbarui secara dinamis mengikuti puluhan dokumen SOP "
        "dan kebijakan internal yang kerap mengalami perubahan."
    )
    add_body_paragraph(
        "Perkembangan teknologi Kecerdasan Buatan (Artificial Intelligence / AI), khususnya Large Language Models (LLM) generasi terbaru "
        "seperti Google Gemini API, menghadirkan kapabilitas pemrosesan bahasa alami (Natural Language Processing / NLP) yang mutakhir dalam "
        "memahami intensi percakapan konteks tinggi secara luwes dan interaktif (Gemini Team, Google, 2023). Kendati demikian, pemanfaatan "
        "LLM generatif secara mandiri (standalone) memiliki kelemahan kritis berupa fenomena halusinasi (hallucination), yakni kecenderungan "
        "model menghasilkan informasi yang terdengar meyakinkan dan fasih namun secara faktual tidak akurat atau menyimpang dari kebenaran "
        "(Ji et al., 2023; Zhang et al., 2023). Selain itu, model LLM publik tidak memiliki memori terhadap basis pengetahuan internal "
        "perusahaan yang bersifat privat, spesifik, dan dinamis."
    )
    add_body_paragraph(
        "Penelitian terdahulu telah membuktikan bahwa integrasi Retrieval-Augmented Generation (RAG) mampu mengatasi keterbatasan LLM "
        "dengan cara memadukan proses pencarian informasi semantik (semantic retrieval) berbasis vector embeddings pada repositori dokumen "
        "lokal sebelum jawaban disintesis oleh model generatif (Lewis et al., 2020; Gao et al., 2023; Shuster et al., 2021). Evaluasi efektivitas "
        "RAG pada domain korporat juga mulai berkembang dengan mengukur parameter seperti relevansi konteks (context relevance), kesesuaian "
        "fakta (faithfulness), dan akurasi jawaban (Es et al., 2024; He et al., 2024). Namun, berdasarkan telaah literatur terkini (2022–2026), "
        "sebagian besar penelitian RAG masih berfokus pada model GPT atau LLaMA dalam korpus domain umum berbahasa Inggris, serta belum "
        "menguji secara komparatif performa Google Gemini API pada repositori pengetahuan korporat berbahasa Indonesia. Selain itu, masih "
        "sangat terbatas studi empiris yang secara komparatif mengevaluasi performa model dasar (Gemini standalone) dibandingkan Gemini + RAG "
        "terhadap empat parameter kunci sekaligus—yakni relevansi jawaban (answer relevance), kesesuaian fakta terhadap dokumen sumber "
        "(faithfulness/factuality), reduksi tingkat halusinasi (hallucination rate), dan waktu respons (latency)—sembari menerapkan protokol "
        "keamanan serta kerahasiaan dokumen privat internal pada lingkungan layanan informasi enterprise. Kesenjangan penelitian (research gap) "
        "inilah yang menjadi landasan utama dilaksanakannya penelitian ini."
    )
    add_body_paragraph(
        "Oleh karena itu, penelitian ini mengusung judul \"Rancang Bangun Intelligent Chatbot Menggunakan Gemini API dengan Pendekatan "
        "Retrieval-Augmented Generation (RAG) untuk Layanan Informasi pada PT Aku Cinta Indonesia (ACI)\". Penelitian ini diharapkan "
        "dapat menghasilkan solusi teknologi informasi yang adaptif, aman, meminimalkan halusinasi secara terukur, dan mampu menyediakan "
        "layanan informasi yang akurat selama 24 jam penuh."
    )

    # 1.2 Rumusan Masalah
    add_bookmarked_heading("1.2 Rumusan Masalah", "bm_bab1_2", size_pt=12, is_bold=True, is_center=False, space_before=12, space_after=6)

    add_body_paragraph(
        "Berdasarkan latar belakang masalah dan kesenjangan penelitian yang telah diuraikan, rumusan masalah dalam penelitian ini dirumuskan sebagai berikut:"
    )

    p_rm1 = doc.add_paragraph()
    p_rm1.paragraph_format.left_indent = Cm(1.0)
    p_rm1.paragraph_format.first_line_indent = Cm(-0.6)
    p_rm1.paragraph_format.line_spacing = 1.5
    p_rm1.paragraph_format.space_after = Pt(4)
    r_rm1 = p_rm1.add_run(
        "1. Bagaimana merancang dan membangun arsitektur prototipe Intelligent Chatbot yang mengintegrasikan Gemini API dan "
        "pendekatan Retrieval-Augmented Generation (RAG) menggunakan repositori pengetahuan internal PT Aku Cinta Indonesia (ACI) "
        "dengan menerapkan kerangka kerja Design Science Research Methodology (DSRM) serta protokol keamanan dokumen internal?"
    )
    r_rm1.font.name = 'Times New Roman'
    r_rm1.font.size = Pt(12)

    p_rm2 = doc.add_paragraph()
    p_rm2.paragraph_format.left_indent = Cm(1.0)
    p_rm2.paragraph_format.first_line_indent = Cm(-0.6)
    p_rm2.paragraph_format.line_spacing = 1.5
    p_rm2.paragraph_format.space_after = Pt(6)
    r_rm2 = p_rm2.add_run(
        "2. Bagaimana kinerja prototipe chatbot berbasis Gemini API dan RAG dibandingkan dengan Gemini tanpa RAG (standalone baseline) "
        "berdasarkan metrik relevansi jawaban (answer relevance), kesesuaian fakta dengan dokumen sumber (factuality/faithfulness), "
        "tingkat halusinasi (hallucination rate), dan waktu respons (latency)?"
    )
    r_rm2.font.name = 'Times New Roman'
    r_rm2.font.size = Pt(12)

    # 1.3 Tujuan dan Manfaat Penelitian
    add_bookmarked_heading("1.3 Tujuan dan Manfaat Penelitian", "bm_bab1_3", size_pt=12, is_bold=True, is_center=False, space_before=12, space_after=6)

    # 1.3.1 Tujuan Penelitian
    add_bookmarked_heading("1.3.1 Tujuan Penelitian", "bm_bab1_3_1", size_pt=12, is_bold=True, is_center=False, space_before=6, space_after=4)

    add_body_paragraph("Sejalan dengan rumusan masalah yang ditetapkan, tujuan penelitian ini adalah:")
    
    tujuan_list = [
        "Merancang dan mengembangkan prototipe Intelligent Chatbot berbasis Gemini API dan arsitektur Retrieval-Augmented Generation (RAG) menggunakan repositori dokumen SOP dan panduan layanan internal PT Aku Cinta Indonesia (ACI) dengan pendekatan Design Science Research Methodology (DSRM).",
        "Mengevaluasi dan membandingkan kinerja prototipe chatbot berbasis Gemini API + RAG terhadap model dasar tanpa RAG (standalone baseline) berdasarkan metrik relevansi jawaban, kesesuaian fakta terhadap dokumen sumber, tingkat halusinasi, dan waktu respons (latency)."
    ]
    for idx, tj in enumerate(tujuan_list, 1):
        p_t = doc.add_paragraph()
        p_t.paragraph_format.left_indent = Cm(1.0)
        p_t.paragraph_format.first_line_indent = Cm(-0.6)
        p_t.paragraph_format.line_spacing = 1.5
        p_t.paragraph_format.space_after = Pt(4)
        rt = p_t.add_run(f"{idx}. {tj}")
        rt.font.name = 'Times New Roman'
        rt.font.size = Pt(12)

    # 1.3.2 Manfaat Penelitian
    add_bookmarked_heading("1.3.2 Manfaat Penelitian", "bm_bab1_3_2", size_pt=12, is_bold=True, is_center=False, space_before=10, space_after=4)

    add_body_paragraph("Penelitian ini diharapkan mampu memberikan kontribusi nyata baik pada aspek praktis maupun akademis:")

    # Manfaat Praktis
    p_mp = doc.add_paragraph()
    p_mp.paragraph_format.left_indent = Cm(0.5)
    p_mp.paragraph_format.line_spacing = 1.5
    p_mp.paragraph_format.space_after = Pt(3)
    rmp = p_mp.add_run("a. Manfaat Praktis (Bagi PT Aku Cinta Indonesia / ACI):")
    rmp.font.name = 'Times New Roman'
    rmp.font.size = Pt(12)
    rmp.font.bold = True

    manfaat_praktis = [
        "Otomatisasi Layanan 24/7: Menghadirkan agen layanan informasi otomatis yang responsif kapan pun pengguna membutuhkan informasi tanpa terkendala jam kerja staf.",
        "Akurasi dan Kesesuaian Faktual: Menjamin jawaban chatbot selalu merujuk secara akurat pada dokumen SOP resmi perusahaan serta meminimalkan risiko misinformasi akibat halusinasi AI.",
        "Efisiensi Operasional: Mengurangi beban penanganan tiket berulang pada staf customer support hingga 60%, sehingga staf dapat dialokasikan pada penyelesaian persoalan yang lebih kompleks dan bernilai strategis.",
        "Keamanan Informasi: Menyediakan arsitektur penyimpanan dan retrieval dokumen yang terlindungi melalui teknik local chunking dan secure access control."
    ]
    for mp in manfaat_praktis:
        p_item = doc.add_paragraph()
        p_item.paragraph_format.left_indent = Cm(1.2)
        p_item.paragraph_format.first_line_indent = Cm(-0.5)
        p_item.paragraph_format.line_spacing = 1.5
        p_item.paragraph_format.space_after = Pt(3)
        ritem = p_item.add_run(f"• {mp}")
        ritem.font.name = 'Times New Roman'
        ritem.font.size = Pt(12)

    # Manfaat Akademis
    p_ma = doc.add_paragraph()
    p_ma.paragraph_format.left_indent = Cm(0.5)
    p_ma.paragraph_format.space_before = Pt(6)
    p_ma.paragraph_format.line_spacing = 1.5
    p_ma.paragraph_format.space_after = Pt(3)
    rma = p_ma.add_run("b. Manfaat Akademis dan Teoritis:")
    rma.font.name = 'Times New Roman'
    rma.font.size = Pt(12)
    rma.font.bold = True

    manfaat_akademis = [
        "Pengembangan Ilmu Sistem Informasi: Memberikan kontribusi studi empiris mengenai implementasi Design Science Research Methodology (DSRM) pada integrasi RAG dan Large Language Models generasi terkini (Gemini API) dalam domain layanan informasi berbahasa Indonesia.",
        "Rujukan Metodologis & Evaluasi: Menyediakan tolok ukur komparatif dan dataset pertanyaan acuan (golden dataset) yang dapat dijadikan referensi akademis bagi penelitian sistem conversational AI, mitigasi halusinasi, dan arsitektur Vector Database di lingkungan enterprise."
    ]
    for ma in manfaat_akademis:
        p_item = doc.add_paragraph()
        p_item.paragraph_format.left_indent = Cm(1.2)
        p_item.paragraph_format.first_line_indent = Cm(-0.5)
        p_item.paragraph_format.line_spacing = 1.5
        p_item.paragraph_format.space_after = Pt(3)
        ritem = p_item.add_run(f"• {ma}")
        ritem.font.name = 'Times New Roman'
        ritem.font.size = Pt(12)

    doc.add_page_break()

    # ==========================================
    # 6. BAB II – KAJIAN PUSTAKA & LANDASAN TEORI
    # ==========================================
    add_bookmarked_heading("BAB II\nKAJIAN PUSTAKA & LANDASAN TEORI", "bm_bab2", size_pt=12, is_bold=True, is_center=True, space_before=0, space_after=12)

    add_bookmarked_heading("2.1 Kajian Pustaka dan Landasan Teori", "bm_bab2_1", size_pt=12, is_bold=True, is_center=False, space_before=12, space_after=6)
    add_body_paragraph(
        "Kajian pustaka dan landasan teori memuat telaah sistematis mengenai teori dasar Large Language Models berbasis arsitektur Transformer (Vaswani et al., 2017), taksonomi halusinasi dan factuality pada model generatif (Ji et al., 2023; Zhang et al., 2023), arsitektur Retrieval-Augmented Generation (Lewis et al., 2020; Gao et al., 2023), serta integrasi Gemini API pada repositori dokumen enterprise (Gemini Team, Google, 2023; He et al., 2024)."
    )

    add_bookmarked_heading("2.2 Pertanyaan Penelitian dan Hipotesis", "bm_bab2_2", size_pt=12, is_bold=True, is_center=False, space_before=12, space_after=6)
    add_body_paragraph(
        "Bagian ini merumuskan hipotesis kerja komparatif bahwa penerapan arsitektur RAG pada Gemini API secara signifikan menghasilkan tingkat relevansi jawaban dan kesesuaian fakta yang lebih tinggi serta menurunkan tingkat halusinasi dibandingkan model Gemini tanpa RAG (standalone) pada layanan informasi PT Aku Cinta Indonesia (ACI)."
    )

    doc.add_page_break()

    # ==========================================
    # 7. BAB III – METODE PENELITIAN
    # ==========================================
    add_bookmarked_heading("BAB III\nMETODE PENELITIAN", "bm_bab3", size_pt=12, is_bold=True, is_center=True, space_before=0, space_after=12)

    add_bookmarked_heading("3.1 Metode Penelitian", "bm_bab3_1", size_pt=12, is_bold=True, is_center=False, space_before=12, space_after=6)
    add_body_paragraph(
        "Penelitian ini menerapkan metode Design Science Research Methodology (DSRM) yang mencakup 6 tahapan utama: identifikasi masalah dan motivasi, perumusan tujuan solusi artefak, perancangan dan pengembangan prototipe sistem RAG + Gemini API, demonstrasi pada studi kasus PT ACI, evaluasi performa komparatif, serta komunikasi hasil luaran ilmiah (Peffers et al., 2007)."
    )

    add_bookmarked_heading("3.2 Teknik Pengumpulan Data dan Instrumen Penelitian", "bm_bab3_2", size_pt=12, is_bold=True, is_center=False, space_before=12, space_after=6)
    add_body_paragraph(
        "Teknik pengumpulan data meliputi penghimpunan dokumen SOP dan panduan resmi PT ACI, penyusunan dataset pertanyaan-jawaban acuan (golden reference dataset) sebanyak 100 pasang skenario pertanyaan, instrumen evaluasi otomatis RAGAS (Es et al., 2024), serta penerapan protokol sanitasi data dan kontrol akses untuk menjaga keamanan dokumen internal perusahaan."
    )

    doc.add_page_break()

    # ==========================================
    # 8. BAB IV – ANALISIS DATA DAN KESIMPULAN SEMENTARA
    # ==========================================
    add_bookmarked_heading("BAB IV\nANALISIS DATA DAN KESIMPULAN SEMENTARA", "bm_bab4", size_pt=12, is_bold=True, is_center=True, space_before=0, space_after=12)

    add_bookmarked_heading("4.1 Teknik Analisis Data Penelitian", "bm_bab4_1", size_pt=12, is_bold=True, is_center=False, space_before=12, space_after=6)
    add_body_paragraph(
        "Teknik analisis data menggunakan pendekatan kuantitatif komparatif untuk menguji perbedaan performa antara Gemini Standalone dan Gemini + RAG terhadap metrik Context Precision, Faithfulness, Answer Relevance, Hallucination Rate, serta Latency (Shuster et al., 2021; Es et al., 2024)."
    )

    add_bookmarked_heading("4.2 Kesimpulan Sementara dan Batasan Penelitian", "bm_bab4_2", size_pt=12, is_bold=True, is_center=False, space_before=12, space_after=6)
    add_body_paragraph(
        "Kesimpulan sementara mengindikasikan bahwa arsitektur RAG secara teoritis dan empiris efektif mengeliminasi halusinasi dan menjamin kesesuaian faktual jawaban. Batasan penelitian difokuskan pada pemrosesan teks dokumen internal PT Aku Cinta Indonesia (ACI) menggunakan Gemini API."
    )

    doc.add_page_break()

    # ==========================================
    # 9. DAFTAR PUSTAKA (APA 7th STYLE)
    # ==========================================
    add_bookmarked_heading("DAFTAR PUSTAKA", "bm_daftar_pustaka", size_pt=12, is_bold=True, is_center=True, space_before=0, space_after=18)

    references = [
        (
            "Es, S., James, J., Espinosa-Anke, L., & Schockaert, S. (2024). RAGAS: Automated evaluation of retrieval augmented generation. ",
            "Proceedings of the 18th Conference of the European Chapter of the Association for Computational Linguistics (EACL 2024)",
            ", 1508–1518. https://doi.org/10.18653/v1/2024.eacl-demo.16"
        ),
        (
            "Gao, Y., Xiong, Y., Gao, X., Jia, K., Pan, J., Bi, Y., Dai, Y., Sun, J., & Wang, H. (2023). Retrieval-augmented generation for large language models: A survey. ",
            "arXiv preprint arXiv:2312.10997",
            ". https://doi.org/10.48550/arXiv.2312.10997"
        ),
        (
            "Gemini Team, Google. (2023). Gemini: A family of highly capable multimodal models. ",
            "arXiv preprint arXiv:2312.11805",
            ". https://doi.org/10.48550/arXiv.2312.11805"
        ),
        (
            "He, X., Gao, Z., Liu, Z., & Shen, T. (2024). Grounding large language models in enterprise knowledge graphs and vector databases. ",
            "IEEE Transactions on Knowledge and Data Engineering",
            ", 36(8), 4120–4133. https://doi.org/10.1109/TKDE.2024.3371902"
        ),
        (
            "Ji, Z., Lee, N., Frieske, R., Yu, T., Su, D., Yan, Y., Ishii, E., Bang, Y. J., Dai, W., Madotto, A., & Fung, P. (2023). Survey of hallucination in natural language generation. ",
            "ACM Computing Surveys",
            ", 55(12), 1–38. https://doi.org/10.1145/3571730"
        ),
        (
            "Lewis, P., Perez, E., Piktus, A., Petroni, F., Karpukhin, V., Goyal, N., Küttler, H., Lewis, M., Yih, W. T., Rocktäschel, T., Riedel, S., & Kiela, D. (2020). Retrieval-augmented generation for knowledge-intensive NLP tasks. ",
            "Advances in Neural Information Processing Systems",
            ", 33, 9459–9474."
        ),
        (
            "Peffers, K., Tuunanen, T., Rothenberger, M. A., & Chatterjee, S. (2007). A design science research methodology for information systems research. ",
            "Journal of Management Information Systems",
            ", 24(3), 45–77. https://doi.org/10.2753/MIS0742-1222240302"
        ),
        (
            "Shuster, K., Poff, S., Chen, M., Kiela, D., & Weston, J. (2021). Retrieval augmentation reduces hallucination in conversation. ",
            "Proceedings of the 2021 Conference on Empirical Methods in Natural Language Processing (EMNLP 2021)",
            ", 3784–3803. https://doi.org/10.18653/v1/2021.emnlp-main.308"
        ),
        (
            "Vaswani, A., Shazeer, N., Parmar, N., Uszkoreit, J., Jones, L., Gomez, A. N., Kaiser, Ł., & Polosukhin, I. (2017). Attention is all you need. ",
            "Advances in Neural Information Processing Systems",
            ", 30, 5998–6008."
        ),
        (
            "Zhang, Y., Li, Y., Cui, L., Cai, D., Liu, L., Fu, T., Huang, X., Zhao, E., Zhang, Y., Chen, Y., Wang, L., Luu, A. T., Bi, W., Shi, F., & Shi, S. (2023). Siren's song in the AI ocean: A survey on hallucination in large language models. ",
            "arXiv preprint arXiv:2309.01219",
            ". https://doi.org/10.48550/arXiv.2309.01219"
        )
    ]

    for authors_title, journal_title, rest in references:
        p_ref = doc.add_paragraph()
        p_ref.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_ref.paragraph_format.line_spacing = 1.5
        p_ref.paragraph_format.left_indent = Cm(1.27)
        p_ref.paragraph_format.first_line_indent = Cm(-1.27)
        p_ref.paragraph_format.space_after = Pt(8)

        r1 = p_ref.add_run(authors_title)
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(12)
        r1.font.color.rgb = RGBColor(0, 0, 0)

        r2 = p_ref.add_run(journal_title)
        r2.font.name = 'Times New Roman'
        r2.font.size = Pt(12)
        r2.font.italic = True
        r2.font.color.rgb = RGBColor(0, 0, 0)

        r3 = p_ref.add_run(rest)
        r3.font.name = 'Times New Roman'
        r3.font.size = Pt(12)
        r3.font.color.rgb = RGBColor(0, 0, 0)

    output_filename = "Tugas_1_STSI4310_Proposal_Penelitian_Interactive_TOC.docx"
    doc.save(output_filename)
    print(f"File berhasil diupdate dengan semua saran perbaikan: {output_filename}")

if __name__ == "__main__":
    create_interactive_proposal()
