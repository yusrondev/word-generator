import docx
from docx.shared import Cm, Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def create_proposal_document():
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
    # Normal Style (Body text)
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(12)
    style_normal.font.color.rgb = RGBColor(0, 0, 0)
    style_normal.paragraph_format.line_spacing = 1.5
    style_normal.paragraph_format.space_after = Pt(6)
    style_normal.paragraph_format.space_before = Pt(0)
    style_normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Custom function to style headings
    def configure_heading_style(style_name, size_pt, is_bold=True, is_center=False):
        try:
            st = doc.styles[style_name]
        except KeyError:
            st = doc.styles.add_style(style_name, docx.enum.style.WD_STYLE_TYPE.PARAGRAPH)
        st.font.name = 'Times New Roman'
        st.font.size = Pt(size_pt)
        st.font.bold = is_bold
        st.font.color.rgb = RGBColor(0, 0, 0)
        st.paragraph_format.line_spacing = 1.5
        st.paragraph_format.space_before = Pt(12)
        st.paragraph_format.space_after = Pt(6)
        st.paragraph_format.keep_with_next = True
        if is_center:
            st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        return st

    configure_heading_style('Heading 1', 12, is_bold=True, is_center=True)
    configure_heading_style('Heading 2', 12, is_bold=True, is_center=False)
    configure_heading_style('Heading 3', 12, is_bold=True, is_center=False)

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

    # Helper function for TOC entries with dot leaders
    def add_toc_line(title, page_num, level=1, is_bold=False):
        p = doc.add_paragraph()
        p.style = 'Normal'
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.space_before = Pt(2 if level == 1 else 0)
        
        # Tab stop at the right margin (14.0 cm) with dots
        p.paragraph_format.tab_stops.add_tab_stop(Cm(PRINTABLE_WIDTH_CM), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        
        if level == 2:
            p.paragraph_format.left_indent = Cm(0.6)
        elif level == 3:
            p.paragraph_format.left_indent = Cm(1.2)
            
        r_title = p.add_run(title)
        r_title.font.name = 'Times New Roman'
        r_title.font.size = Pt(11.5 if level == 1 else 11)
        r_title.font.bold = is_bold
        r_title.font.color.rgb = RGBColor(0, 0, 0)

        # Tab character
        p.add_run("\t")

        # Page number
        r_page = p.add_run(str(page_num))
        r_page.font.name = 'Times New Roman'
        r_page.font.size = Pt(11.5 if level == 1 else 11)
        r_page.font.bold = is_bold
        r_page.font.color.rgb = RGBColor(0, 0, 0)
        return p

    # ==========================================
    # 3. HALAMAN COVER / JUDUL
    # ==========================================
    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_top.paragraph_format.space_before = Pt(10)
    p_top.paragraph_format.space_after = Pt(24)
    r = p_top.add_run("PROPOSAL PENELITIAN")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.line_spacing = 1.15
    p_title.paragraph_format.space_after = Pt(24)
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
    p_sub.paragraph_format.space_after = Pt(60)
    r_sub = p_sub.add_run("Tugas 1 Mata Kuliah STSI4310 Metode Penelitian")
    r_sub.font.name = 'Times New Roman'
    r_sub.font.size = Pt(12)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(0, 0, 0)

    # Identitas Penyusun (Tabel bersih tanpa border agar rapi)
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
            
        # Set column widths
        row.cells[0].width = Cm(4.5)
        row.cells[1].width = Cm(0.5)
        row.cells[2].width = Cm(8.0)

    # Remove borders from identity table
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
    p_bot.paragraph_format.space_before = Pt(80)
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
    # 4. HALAMAN DAFTAR ISI (TAB STOPS + DOT LEADER)
    # ==========================================
    p_toc_head = doc.add_paragraph()
    p_toc_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_toc_head.paragraph_format.space_before = Pt(0)
    p_toc_head.paragraph_format.space_after = Pt(18)
    r_toc_h = p_toc_head.add_run("DAFTAR ISI")
    r_toc_h.font.name = 'Times New Roman'
    r_toc_h.font.size = Pt(12)
    r_toc_h.font.bold = True
    r_toc_h.font.color.rgb = RGBColor(0, 0, 0)

    # Isi Daftar Isi sesuai sistematika PDF tugas STSI4310
    add_toc_line("HALAMAN COVER", "i", level=1, is_bold=True)
    add_toc_line("DAFTAR ISI", "ii", level=1, is_bold=True)
    add_toc_line("BAB I PENDAHULUAN", "1", level=1, is_bold=True)
    add_toc_line("1.1 Latar Belakang Masalah", "1", level=2)
    add_toc_line("1.2 Rumusan Masalah", "2", level=2)
    add_toc_line("1.3 Tujuan dan Manfaat Penelitian", "2", level=2)
    add_toc_line("1.3.1 Tujuan Penelitian", "2", level=3)
    add_toc_line("1.3.2 Manfaat Penelitian", "3", level=3)
    add_toc_line("BAB II KAJIAN PUSTAKA & LANDASAN TEORI", "4", level=1, is_bold=True)
    add_toc_line("2.1 Kajian Pustaka dan Landasan Teori (KOSONGKAN)", "4", level=2)
    add_toc_line("2.2 Pertanyaan Penelitian dan Hipotesis (KOSONGKAN)", "4", level=2)
    add_toc_line("BAB III METODE PENELITIAN", "5", level=1, is_bold=True)
    add_toc_line("3.1 Metode Penelitian (KOSONGKAN)", "5", level=2)
    add_toc_line("3.2 Teknik Pengumpulan Data dan Instrumen Penelitian (KOSONGKAN)", "5", level=2)
    add_toc_line("BAB IV ANALISIS DATA DAN KESIMPULAN SEMENTARA", "6", level=1, is_bold=True)
    add_toc_line("4.1 Teknik Analisis Data Penelitian (KOSONGKAN)", "6", level=2)
    add_toc_line("4.2 Kesimpulan Sementara dan Batasan Penelitian (KOSONGKAN)", "6", level=2)
    add_toc_line("DAFTAR PUSTAKA PENELITIAN", "7", level=1, is_bold=True)

    doc.add_page_break()

    # ==========================================
    # 5. BAB I – PENDAHULUAN
    # ==========================================
    p_b1 = doc.add_paragraph()
    p_b1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_b1.paragraph_format.space_before = Pt(0)
    p_b1.paragraph_format.space_after = Pt(12)
    r_b1 = p_b1.add_run("BAB I\nPENDAHULUAN")
    r_b1.font.name = 'Times New Roman'
    r_b1.font.size = Pt(12)
    r_b1.font.bold = True
    r_b1.font.color.rgb = RGBColor(0, 0, 0)

    # 1.1 Latar Belakang Masalah
    p_h11 = doc.add_paragraph()
    p_h11.paragraph_format.space_before = Pt(12)
    p_h11.paragraph_format.space_after = Pt(6)
    p_h11.paragraph_format.keep_with_next = True
    r_h11 = p_h11.add_run("1.1 Latar Belakang Masalah")
    r_h11.font.name = 'Times New Roman'
    r_h11.font.size = Pt(12)
    r_h11.font.bold = True
    r_h11.font.color.rgb = RGBColor(0, 0, 0)

    add_body_paragraph(
        "Di era transformasi digital yang berkembang pesat saat ini, efisiensi dan kecepatan dalam penyampaian "
        "informasi layanan menjadi salah satu faktor penentu kepuasan pengguna serta keberlanjutan operasional perusahaan. "
        "PT Aku Cinta Indonesia (ACI) sebagai perusahaan yang bergerak di bidang penyediaan layanan informasi dan produk "
        "digital terus mengalami pertumbuhan basis pelanggan yang signifikan. Seiring dengan peningkatan tersebut, volume "
        "pertanyaan umum (Frequently Asked Questions / FAQ) yang masuk ke kanal layanan informasi perusahaan meningkat "
        "drastis setiap bulannya."
    )
    add_body_paragraph(
        "Berdasarkan observasi pada operasional layanan di PT Aku Cinta Indonesia (ACI), terdapat beberapa kendala utama "
        "yang dihadapi dalam penanganan informasi pelanggan. Pertama, keterbatasan jam kerja staf operasional/customer support "
        "mengakibatkan pertanyaan pengguna yang diajukan di luar jam kerja (malam hari atau hari libur) mengalami penundaan "
        "respons (delay response) hingga hari kerja berikutnya. Kedua, tingginya frekuensi pertanyaan berulang (repetitive inquiries) "
        "mengenai panduan layanan, prosedur administrasi, dan spesifikasi produk menyita hingga 60% waktu produktif staf operasional, "
        "sehingga mengurangi efektivitas penanganan kasus-kasus yang membutuhkan penalaran kompleks. Ketiga, solusi chatbot "
        "konvensional berbasis aturan kaku (rule-based chatbot) yang sebelumnya digunakan terbukti tidak fleksibel, kerap gagal "
        "memahami variasi bahasa alami pengguna, serta memiliki keterbatasan dalam memperbarui basis data pengetahuan tanpa pemrograman ulang."
    )
    add_body_paragraph(
        "Perkembangan teknologi Kecerdasan Buatan (Artificial Intelligence / AI), khususnya Large Language Models (LLM) seperti "
        "Google Gemini API, menawarkan kapabilitas pemrosesan bahasa alami (Natural Language Processing / NLP) yang mutakhir dalam "
        "memahami konteks percakapan manusia secara luwes dan generatif. Kendati demikian, pemanfaatan LLM secara mandiri (standalone) "
        "memiliki risiko kelemahan kritis berupa fenomena halusinasi (hallucination), yakni kondisi di mana model menghasilkan "
        "jawaban yang terdengar meyakinkan namun secara faktual keliru atau tidak sesuai dengan data internal perusahaan (Vaswani et al., 2017; Team et al., 2023)."
    )
    add_body_paragraph(
        "Untuk mengatasi kesenjangan tersebut (research gap), pendekatan Retrieval-Augmented Generation (RAG) diintegrasikan sebagai "
        "arsitektur solusi. Metode RAG bekerja dengan cara menggabungkan teknik pencarian informasi (Information Retrieval) berbasis "
        "pencocokan semantik (vector embeddings) pada basis pengetahuan internal dokumen PT Aku Cinta Indonesia (ACI), kemudian "
        "menyertakan konteks dokumen relevan tersebut ke dalam prompt Gemini API (Lewis et al., 2020). Melalui mekanisme ini, respon "
        "yang dihasilkan chatbot tetap berakar kuat (grounded) pada data faktual perusahaan, aman dari kebocoran konteks, serta minim halusinasi."
    )
    add_body_paragraph(
        "Oleh karena itu, penelitian mengenai \"Rancang Bangun Intelligent Chatbot Menggunakan Gemini API dengan Pendekatan Retrieval-Augmented Generation (RAG) untuk Layanan Informasi pada PT Aku Cinta Indonesia (ACI)\" menjadi sangat strategis dan mendesak untuk dilaksanakan guna mewujudkan sistem layanan informasi yang cerdas, akurat, responsif, dan dapat diakses selama 24 jam penuh."
    )

    # 1.2 Rumusan Masalah
    p_h12 = doc.add_paragraph()
    p_h12.paragraph_format.space_before = Pt(12)
    p_h12.paragraph_format.space_after = Pt(6)
    p_h12.paragraph_format.keep_with_next = True
    r_h12 = p_h12.add_run("1.2 Rumusan Masalah")
    r_h12.font.name = 'Times New Roman'
    r_h12.font.size = Pt(12)
    r_h12.font.bold = True
    r_h12.font.color.rgb = RGBColor(0, 0, 0)

    add_body_paragraph(
        "Berdasarkan latar belakang masalah yang telah diuraikan, rumusan masalah dalam penelitian ini disajikan dalam bentuk pertanyaan penelitian sebagai berikut:"
    )

    p_rm1 = doc.add_paragraph()
    p_rm1.paragraph_format.left_indent = Cm(1.0)
    p_rm1.paragraph_format.first_line_indent = Cm(-0.6)
    p_rm1.paragraph_format.line_spacing = 1.5
    p_rm1.paragraph_format.space_after = Pt(4)
    r_rm1 = p_rm1.add_run("1. Bagaimana merancang dan membangun arsitektur Intelligent Chatbot yang mengintegrasikan Gemini API dan pendekatan Retrieval-Augmented Generation (RAG) untuk layanan informasi pada PT Aku Cinta Indonesia (ACI)?")
    r_rm1.font.name = 'Times New Roman'
    r_rm1.font.size = Pt(12)

    p_rm2 = doc.add_paragraph()
    p_rm2.paragraph_format.left_indent = Cm(1.0)
    p_rm2.paragraph_format.first_line_indent = Cm(-0.6)
    p_rm2.paragraph_format.line_spacing = 1.5
    p_rm2.paragraph_format.space_after = Pt(6)
    r_rm2 = p_rm2.add_run("2. Bagaimana kinerja dan akurasi Intelligent Chatbot berbasis RAG dan Gemini API dalam memberikan respons yang relevan, faktual, serta minim halusinasi terhadap pertanyaan layanan di PT Aku Cinta Indonesia (ACI)?")
    r_rm2.font.name = 'Times New Roman'
    r_rm2.font.size = Pt(12)

    # 1.3 Tujuan dan Manfaat Penelitian
    p_h13 = doc.add_paragraph()
    p_h13.paragraph_format.space_before = Pt(12)
    p_h13.paragraph_format.space_after = Pt(6)
    p_h13.paragraph_format.keep_with_next = True
    r_h13 = p_h13.add_run("1.3 Tujuan dan Manfaat Penelitian")
    r_h13.font.name = 'Times New Roman'
    r_h13.font.size = Pt(12)
    r_h13.font.bold = True
    r_h13.font.color.rgb = RGBColor(0, 0, 0)

    # 1.3.1 Tujuan Penelitian
    p_h131 = doc.add_paragraph()
    p_h131.paragraph_format.space_before = Pt(6)
    p_h131.paragraph_format.space_after = Pt(4)
    p_h131.paragraph_format.left_indent = Cm(0.0)
    p_h131.paragraph_format.keep_with_next = True
    r_h131 = p_h131.add_run("1.3.1 Tujuan Penelitian")
    r_h131.font.name = 'Times New Roman'
    r_h131.font.size = Pt(12)
    r_h131.font.bold = True

    add_body_paragraph("Sejalan dengan rumusan masalah di atas, tujuan yang hendak dicapai dalam penelitian ini adalah:")
    
    tujuan_list = [
        "Mengembangkan dan mengimplementasikan prototipe Intelligent Chatbot dengan mengintegrasikan Gemini API dan arsitektur Retrieval-Augmented Generation (RAG) pada repositori pengetahuan internal PT Aku Cinta Indonesia (ACI).",
        "Menguji dan menganalisis performa akurasi jawaban, relevansi informasi, serta waktu tanggap (responsivitas) sistem chatbot dalam melayani pertanyaan seputar layanan informasi perusahaan."
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
    p_h132 = doc.add_paragraph()
    p_h132.paragraph_format.space_before = Pt(10)
    p_h132.paragraph_format.space_after = Pt(4)
    p_h132.paragraph_format.left_indent = Cm(0.0)
    p_h132.paragraph_format.keep_with_next = True
    r_h132 = p_h132.add_run("1.3.2 Manfaat Penelitian")
    r_h132.font.name = 'Times New Roman'
    r_h132.font.size = Pt(12)
    r_h132.font.bold = True

    add_body_paragraph("Penelitian ini diharapkan dapat memberikan kontribusi nyata baik secara praktis maupun akademis:")

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
        "Meningkatkan efisiensi dan ketersediaan layanan informasi secara otomatis selama 24 jam 7 hari tanpa jeda.",
        "Meningkatkan konsistensi dan akurasi informasi yang diberikan kepada pengguna karena jawaban mengacu langsung pada dokumen resmi perusahaan.",
        "Mengurangi beban kerja repetitif dari staf admin/customer support sehingga dapat dialokasikan untuk tugas strategis lainnya."
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
        "Memberikan sumbangsih empiris dalam disiplin ilmu Sistem Informasi terkait integrasi teknologi Retrieval-Augmented Generation (RAG) dengan Large Language Models (LLM).",
        "Menjadi rujukan penelitian lanjutan dalam perancangan sistem Conversational AI dan pemanfaatan Vector Database pada domain enterprise."
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
    p_b2 = doc.add_paragraph()
    p_b2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_b2.paragraph_format.space_before = Pt(0)
    p_b2.paragraph_format.space_after = Pt(12)
    r_b2 = p_b2.add_run("BAB II\nKAJIAN PUSTAKA & LANDASAN TEORI")
    r_b2.font.name = 'Times New Roman'
    r_b2.font.size = Pt(12)
    r_b2.font.bold = True
    r_b2.font.color.rgb = RGBColor(0, 0, 0)

    # 2.1
    p_h21 = doc.add_paragraph()
    p_h21.paragraph_format.space_before = Pt(12)
    p_h21.paragraph_format.space_after = Pt(6)
    p_h21.paragraph_format.keep_with_next = True
    r_h21 = p_h21.add_run("2.1 Kajian Pustaka dan Landasan Teori")
    r_h21.font.name = 'Times New Roman'
    r_h21.font.size = Pt(12)
    r_h21.font.bold = True

    p_k21 = doc.add_paragraph()
    p_k21.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_k21.paragraph_format.space_after = Pt(12)
    r_k21 = p_k21.add_run("- KOSONGKAN -\n(Sesuai petunjuk pengerjaan Tugas 1 STSI4310)")
    r_k21.font.name = 'Times New Roman'
    r_k21.font.size = Pt(11)
    r_k21.font.italic = True

    # 2.2
    p_h22 = doc.add_paragraph()
    p_h22.paragraph_format.space_before = Pt(12)
    p_h22.paragraph_format.space_after = Pt(6)
    p_h22.paragraph_format.keep_with_next = True
    r_h22 = p_h22.add_run("2.2 Pertanyaan Penelitian dan Hipotesis")
    r_h22.font.name = 'Times New Roman'
    r_h22.font.size = Pt(12)
    r_h22.font.bold = True

    p_k22 = doc.add_paragraph()
    p_k22.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_k22.paragraph_format.space_after = Pt(12)
    r_k22 = p_k22.add_run("- KOSONGKAN -\n(Sesuai petunjuk pengerjaan Tugas 1 STSI4310)")
    r_k22.font.name = 'Times New Roman'
    r_k22.font.size = Pt(11)
    r_k22.font.italic = True

    doc.add_page_break()

    # ==========================================
    # 7. BAB III – METODE PENELITIAN
    # ==========================================
    p_b3 = doc.add_paragraph()
    p_b3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_b3.paragraph_format.space_before = Pt(0)
    p_b3.paragraph_format.space_after = Pt(12)
    r_b3 = p_b3.add_run("BAB III\nMETODE PENELITIAN")
    r_b3.font.name = 'Times New Roman'
    r_b3.font.size = Pt(12)
    r_b3.font.bold = True
    r_b3.font.color.rgb = RGBColor(0, 0, 0)

    # 3.1
    p_h31 = doc.add_paragraph()
    p_h31.paragraph_format.space_before = Pt(12)
    p_h31.paragraph_format.space_after = Pt(6)
    p_h31.paragraph_format.keep_with_next = True
    r_h31 = p_h31.add_run("3.1 Metode Penelitian")
    r_h31.font.name = 'Times New Roman'
    r_h31.font.size = Pt(12)
    r_h31.font.bold = True

    p_k31 = doc.add_paragraph()
    p_k31.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_k31.paragraph_format.space_after = Pt(12)
    r_k31 = p_k31.add_run("- KOSONGKAN -\n(Sesuai petunjuk pengerjaan Tugas 1 STSI4310)")
    r_k31.font.name = 'Times New Roman'
    r_k31.font.size = Pt(11)
    r_k31.font.italic = True

    # 3.2
    p_h32 = doc.add_paragraph()
    p_h32.paragraph_format.space_before = Pt(12)
    p_h32.paragraph_format.space_after = Pt(6)
    p_h32.paragraph_format.keep_with_next = True
    r_h32 = p_h32.add_run("3.2 Teknik Pengumpulan Data dan Instrumen Penelitian")
    r_h32.font.name = 'Times New Roman'
    r_h32.font.size = Pt(12)
    r_h32.font.bold = True

    p_k32 = doc.add_paragraph()
    p_k32.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_k32.paragraph_format.space_after = Pt(12)
    r_k32 = p_k32.add_run("- KOSONGKAN -\n(Sesuai petunjuk pengerjaan Tugas 1 STSI4310)")
    r_k32.font.name = 'Times New Roman'
    r_k32.font.size = Pt(11)
    r_k32.font.italic = True

    doc.add_page_break()

    # ==========================================
    # 8. BAB IV – ANALISIS DATA DAN KESIMPULAN SEMENTARA
    # ==========================================
    p_b4 = doc.add_paragraph()
    p_b4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_b4.paragraph_format.space_before = Pt(0)
    p_b4.paragraph_format.space_after = Pt(12)
    r_b4 = p_b4.add_run("BAB IV\nANALISIS DATA DAN KESIMPULAN SEMENTARA")
    r_b4.font.name = 'Times New Roman'
    r_b4.font.size = Pt(12)
    r_b4.font.bold = True
    r_b4.font.color.rgb = RGBColor(0, 0, 0)

    # 4.1
    p_h41 = doc.add_paragraph()
    p_h41.paragraph_format.space_before = Pt(12)
    p_h41.paragraph_format.space_after = Pt(6)
    p_h41.paragraph_format.keep_with_next = True
    r_h41 = p_h41.add_run("4.1 Teknik Analisis Data Penelitian")
    r_h41.font.name = 'Times New Roman'
    r_h41.font.size = Pt(12)
    r_h41.font.bold = True

    p_k41 = doc.add_paragraph()
    p_k41.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_k41.paragraph_format.space_after = Pt(12)
    r_k41 = p_k41.add_run("- KOSONGKAN -\n(Sesuai petunjuk pengerjaan Tugas 1 STSI4310)")
    r_k41.font.name = 'Times New Roman'
    r_k41.font.size = Pt(11)
    r_k41.font.italic = True

    # 4.2
    p_h42 = doc.add_paragraph()
    p_h42.paragraph_format.space_before = Pt(12)
    p_h42.paragraph_format.space_after = Pt(6)
    p_h42.paragraph_format.keep_with_next = True
    r_h42 = p_h42.add_run("4.2 Kesimpulan Sementara dan Batasan Penelitian")
    r_h42.font.name = 'Times New Roman'
    r_h42.font.size = Pt(12)
    r_h42.font.bold = True

    p_k42 = doc.add_paragraph()
    p_k42.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_k42.paragraph_format.space_after = Pt(12)
    r_k42 = p_k42.add_run("- KOSONGKAN -\n(Sesuai petunjuk pengerjaan Tugas 1 STSI4310)")
    r_k42.font.name = 'Times New Roman'
    r_k42.font.size = Pt(11)
    r_k42.font.italic = True

    doc.add_page_break()

    # ==========================================
    # 9. DAFTAR PUSTAKA PENELITIAN (APA 7th STYLE)
    # ==========================================
    p_dp = doc.add_paragraph()
    p_dp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_dp.paragraph_format.space_before = Pt(0)
    p_dp.paragraph_format.space_after = Pt(18)
    r_dp = p_dp.add_run("DAFTAR PUSTAKA")
    r_dp.font.name = 'Times New Roman'
    r_dp.font.size = Pt(12)
    r_dp.font.bold = True
    r_dp.font.color.rgb = RGBColor(0, 0, 0)

    # APA Style References with Hanging Indent (1.27 cm / 0.5 in)
    references = [
        (
            "Lewis, P., Perez, E., Piktus, A., Petroni, F., Karpukhin, V., Goyal, N., Küttler, H., "
            "Lewis, M., Yih, W. T., Rocktäschel, T., Riedel, S., & Kiela, D. (2020). Retrieval-augmented "
            "generation for knowledge-intensive NLP tasks. ",
            "Advances in Neural Information Processing Systems",
            ", 33, 9459–9474."
        ),
        (
            "Team, G., Anil, R., Borgeaud, S., Wu, Y., Alayrac, J. B., Yu, J., Soricut, R., Schalkwyk, P., "
            "Dai, A. M., Hauth, A., & Dehghani, M. (2023). Gemini: A family of highly capable multimodal models. ",
            "arXiv preprint arXiv:2312.11805",
            "."
        ),
        (
            "Vaswani, A., Shazeer, N., Parmar, N., Uszkoreit, J., Jones, L., Gomez, A. N., Kaiser, Ł., "
            "& Polosukhin, I. (2017). Attention is all you need. ",
            "Advances in Neural Information Processing Systems",
            ", 30, 5998–6008."
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

    output_path = "Tugas_1_STSI4310_Proposal_Penelitian.docx"
    doc.save(output_path)
    print(f"Berhasil meng-generate dokumen profesional: {output_path}")

if __name__ == "__main__":
    create_proposal_document()
