#!/usr/bin/env python
import os
import re
import zipfile
import tempfile
import shutil
import threading
import subprocess
import traceback
import locale
import multiprocessing
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# 3. Parti kütüphaneler
import pdfkit
from PyPDF2 import PdfMerger
from bs4 import BeautifulSoup
import xml.etree.ElementTree as ET

# NOT: Loglama kütüphanesi yapılandırması tamamen kaldırıldı.
# Disk üzerinde .log dosyası oluşturulmayacak ve RAM'de log listesi tutulmayacak.

# ***** İş Mantığı Sınıfı: InvoiceProcessor *****
class InvoiceProcessor:
    def __init__(self, log_callback):
        """
        :param log_callback: Anlık durumu ekrana yazdırmak için kullanılan fonksiyon.
        """
        self.log_callback = log_callback
        # CPU sayısına göre iş parçacığı sayısı belirlenir
        self.max_workers = max(2, multiprocessing.cpu_count() - 1)

    def log_message(self, message):
        """
        Mesajı sadece arayüze (GUI) gönderir.
        Diske kayıt yapmaz, konsola basmaz.
        """
        if self.log_callback:
            self.log_callback(message)

    def extract_zip_recursively(self, zip_path, extract_path, depth=0, max_depth=5):
        """ZIP dosyalarını özyinelemeli olarak çıkarır"""
        if depth > max_depth:
            self.log_message(f"Maksimum derinliğe ulaşıldı, daha fazla açılmıyor: {zip_path}")
            return
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_path)
            self.log_message(f"Zip dosyası açıldı: {os.path.basename(zip_path)}")
        except Exception as e:
            self.log_message(f"Hata: {os.path.basename(zip_path)} açılamadı: {str(e)}")
            return

        # İç içe zip dosyalarını bul ve paralel olarak çıkar
        inner_zips = []
        for root_dir, _, files in os.walk(extract_path):
            for file in files:
                if file.lower().endswith('.zip'):
                    inner_zip_path = os.path.join(root_dir, file)
                    inner_zips.append(inner_zip_path)

        if inner_zips:
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                for inner_zip in inner_zips:
                    inner_extract_path = os.path.join(
                        extract_path,
                        f"extracted_{os.path.splitext(os.path.basename(inner_zip))[0]}"
                    )
                    os.makedirs(inner_extract_path, exist_ok=True)
                    executor.submit(self.extract_zip_recursively, inner_zip, inner_extract_path, depth + 1, max_depth)

    def extract_date_from_xml(self, xml_file):
        """XML dosyasından fatura tarihini çıkarır"""
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()
            namespaces = {
                'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2',
                'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2',
                'ubl': 'urn:oasis:names:specification:ubl:schema:xsd:Invoice-2'
            }
            issue_date = None

            for ns_prefix in [None, 'cbc', 'ubl']:
                tag = f"{{{namespaces[ns_prefix]}}}IssueDate" if ns_prefix else "IssueDate"
                elements = root.findall(f".//{tag}")
                if elements:
                    issue_date = elements[0].text
                    # Loglama kaldırıldı, sadece işlem yapılıyor
                    break

            if not issue_date:
                potential_date_tags = ["IssueDate", "DüzenlemeTarihi", "düzenlemetarihi", "BelgeTarihi", "belgetarihi"]
                for tag in potential_date_tags:
                    elements = root.findall(f".//{tag}")
                    if elements:
                        issue_date = elements[0].text
                        break

            if not issue_date:
                self.log_message(f"XML'de tarih bulunamadı: {os.path.basename(xml_file)}")
                return None

            try:
                if "-" in issue_date:
                    date_obj = datetime.strptime(issue_date, "%Y-%m-%d")
                elif "." in issue_date:
                    date_obj = datetime.strptime(issue_date, "%d.%m.%Y")
                else:
                    self.log_message(f"Geçersiz tarih formatı: {issue_date}")
                    return None
                # Başarılı işlem logu (ekrana bilgi için)
                self.log_message(f"✓ XML'den fatura tarihi: {date_obj.strftime('%d.%m.%Y')}")
                return date_obj
            except ValueError as e:
                self.log_message(f"Tarih ayrıştırma hatası: {str(e)}")
                return None
        except Exception as e:
            self.log_message(f"XML işleme hatası: {os.path.basename(xml_file)} - {str(e)}")
            return None

    def extract_evrak_id(self, xml_file):
        """XML dosyasından evrak ID'sini çıkarır"""
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()
            namespaces = {'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2'}
            id_elem = root.find('.//cbc:ID', namespaces)
            if id_elem is not None:
                evrak_id = id_elem.text.strip()
                if len(evrak_id) == 16:
                    # Ekrana bilgi basma (log dosyasına değil)
                    self.log_message(f"Evrak ID bulundu: {evrak_id}")
                    return evrak_id
                else:
                    self.log_message(f"Evrak ID uygun formatta değil: {evrak_id}")
                    return None
            else:
                self.log_message("XML'de Evrak ID bulunamadı.")
                return None
        except Exception as e:
            self.log_message(f"XML evrak ID işleme hatası: {str(e)}")
            return None

    def extract_invoice_dates(self, html_file):
        """HTML dosyasından fatura tarihini çıkarır"""
        try:
            with open(html_file, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            try:
                soup = BeautifulSoup(content, 'html.parser')
                text_content = soup.get_text()
            except Exception as e:
                self.log_message(f"HTML parse hatası: {str(e)}. Düz metin olarak devam ediliyor.")
                text_content = content

            date_formats = ["%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y"]
            primary_date_keywords = [
                "Düzenleme Tarihi", "Düzenleme tarihi", "düzenleme tarihi",
                "Belge Tarihi", "Belge tarihi", "belge tarihi",
                "Fatura Tarihi", "Fatura tarihi", "fatura tarihi",
                "Düzenlenme Tarihi", "e-Fatura Tarihi", "e-Arşiv Fatura Tarihi",
                "Tarih", "tarih", "TARİH"
            ]
            primary_dates = []
            all_other_dates = []

            for keyword in primary_date_keywords:
                patterns = [
                    rf"{re.escape(keyword)}\s*[:=\-]?\s*(\d{{1,2}}\.\d{{1,2}}\.\d{{4}})",
                    rf"{re.escape(keyword)}\s*[:=\-]?\s*(\d{{1,2}}/\d{{1,2}}/\d{{4}})",
                    rf"{re.escape(keyword)}\s*[:=\-]?\s*(\d{{1,2}}-\d{{1,2}}-\d{{4}})"
                ]
                for pattern in patterns:
                    matches = re.findall(pattern, text_content, re.IGNORECASE)
                    for date_str in matches:
                        for fmt in date_formats:
                            try:
                                date_obj = datetime.strptime(date_str, fmt)
                                if not any(d[0] == date_obj for d in primary_dates):
                                    primary_dates.append((date_obj, keyword))
                                    self.log_message(f"Öncelikli '{keyword}' bulundu: {date_str} ({os.path.basename(html_file)})")
                                break
                            except ValueError:
                                continue

            if primary_dates:
                invoice_date, _ = max(primary_dates, key=lambda x: x[0])
                return invoice_date

            date_patterns = [
                r'\b(\d{1,2}\.\d{1,2}\.\d{4})\b',
                r'\b(\d{1,2}/\d{1,2}/\d{4})\b',
                r'\b(\d{1,2}-\d{1,2}-\d{4})\b'
            ]
            for pattern in date_patterns:
                matches = re.findall(pattern, text_content)
                for date_str in matches:
                    for fmt in date_formats:
                        try:
                            date_obj = datetime.strptime(date_str, fmt)
                            if not any(d[0] == date_obj for d in all_other_dates):
                                all_other_dates.append((date_obj, "Genel Arama"))
                            break
                        except ValueError:
                            continue

            if not all_other_dates:
                self.log_message(f"⚠️ Hiçbir tarih bulunamadı: {os.path.basename(html_file)}")
                return None

            invoice_date, _ = max(all_other_dates, key=lambda x: x[0])
            return invoice_date
        except Exception as e:
            self.log_message(f"⚠️ Tarih çıkarma hatası: {os.path.basename(html_file)} - {str(e)}")
            return None

    def find_files(self, folder, extensions):
        """Belirtilen uzantılara sahip dosyaları bulur"""
        found_files = []
        for root_dir, _, files in os.walk(folder):
            for file in files:
                if os.path.splitext(file)[1].lower() in extensions:
                    found_files.append(os.path.join(root_dir, file))
        return found_files

    def match_html_with_xml(self, html_files, xml_files):
        """HTML ve XML dosyalarını eşleştirir ve tarihlerini çıkarır"""
        self.log_message("HTML ve XML dosyaları eşleştiriliyor...")
        files_with_dates = []
        html_without_xml = []
        
        xml_dict = {}
        for xml_file in xml_files:
            key = os.path.splitext(os.path.basename(xml_file))[0]
            xml_dict[key] = xml_file

        def process_html(html_file):
            base = os.path.splitext(os.path.basename(html_file))[0]
            if base in xml_dict:
                xml_file = xml_dict[base]
                date = self.extract_date_from_xml(xml_file)
                evrak_id = self.extract_evrak_id(xml_file)
                if not date:
                    self.log_message(f"⚠️ {base} için XML'de tarih bulunamadı. HTML'den çıkarılıyor.")
                    date = self.extract_invoice_dates(html_file)
                else:
                    self.log_message(f"✓ HTML-XML eşleşmesi: {base} - Tarih: {date.strftime('%d.%m.%Y') if date else 'Bilinmiyor'}")
                return (html_file, date, evrak_id, False)
            else:
                self.log_message(f"⚠️ {base} için eşleşen XML bulunamadı. HTML'den tarih çıkarılıyor.")
                date = self.extract_invoice_dates(html_file)
                return (html_file, date, None, True)

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            results = list(executor.map(process_html, html_files))
            
        for html_file, date, evrak_id, no_xml in results:
            files_with_dates.append((html_file, date, evrak_id))
            if no_xml:
                html_without_xml.append(html_file)
                
        self.log_message(f"Toplam {len(html_files)} HTML dosyasından:")
        self.log_message(f"- {len(html_files) - len(html_without_xml)} dosya XML ile eşleştirildi")
        self.log_message(f"- {len(html_without_xml)} dosya için XML bulunamadı")
        return files_with_dates

    def convert_html_to_pdf(self, html_file, output_path, config, pdf_options):
        """HTML dosyasını PDF'e dönüştürür. Hata durumunda alternatif yöntemleri dener."""
        base_name = os.path.basename(html_file)
        try:
            pdfkit.from_file(html_file, output_path, configuration=config, options=pdf_options)
            return True, ""
        except Exception as e1:
            error_detail = str(e1)
            self.log_message(f"⚠️ İlk deneme hatası: {base_name} - {error_detail}")
            try:
                self.log_message(f"Alternatif dönüştürme deneniyor: {base_name}")
                simplified_options = {"enable-local-file-access": ""}
                pdfkit.from_file(html_file, output_path, configuration=config, options=simplified_options)
                return True, ""
            except Exception as e2:
                error_detail = str(e2)
                self.log_message(f"⚠️ İkinci deneme başarısız: {base_name} - {error_detail}")
                try:
                    self.log_message(f"Son deneme: {base_name}")
                    with open(html_file, 'r', encoding='utf-8', errors='ignore') as f:
                        html_content = f.read()
                    pdfkit.from_string(html_content, output_path, configuration=config, options=simplified_options)
                    return True, ""
                except Exception as e3:
                    error_detail = str(e3)
                    self.log_message(f"✗ Tüm denemeler başarısız: {base_name} - {error_detail}")
                    return False, error_detail

    def convert_html_to_pdf_parallel(self, html_files_with_dates, temp_dir, config, pdf_options, update_status_callback=None):
        """Birden fazla HTML dosyasını paralel olarak PDF'e dönüştürür"""
        pdf_files_with_info = []
        error_list = []
        total_files = len(html_files_with_dates)

        def convert_one_file(idx, html_file, invoice_date, evrak_id):
            if update_status_callback:
                progress = 50 + (30 * (idx + 1) / total_files)
                update_status_callback(f"Dönüştürülüyor: {os.path.basename(html_file)}", progress)
            if invoice_date:
                dstr = invoice_date.strftime("%Y%m%d")
                pdf_name = f"fatura_{dstr}_{idx+1}.pdf"
            else:
                pdf_name = f"fatura_tarihsiz_{idx+1}.pdf"
            pdf_path = os.path.join(temp_dir, pdf_name)
            success, error = self.convert_html_to_pdf(html_file, pdf_path, config, pdf_options)
            if success:
                return (pdf_path, invoice_date, evrak_id, None)
            else:
                return (None, invoice_date, evrak_id, (evrak_id if evrak_id else "Bilinmiyor", f"Dönüştürme hatası: {error}"))

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = []
            for idx, (html_file, invoice_date, evrak_id) in enumerate(html_files_with_dates):
                futures.append(executor.submit(convert_one_file, idx, html_file, invoice_date, evrak_id))
            for future in futures:
                pdf_path, invoice_date, evrak_id, error = future.result()
                if pdf_path:
                    pdf_files_with_info.append((pdf_path, invoice_date, evrak_id))
                if error:
                    error_list.append(error)
        return pdf_files_with_info, error_list


# ***** Grafiksel Arayüz ve Uygulama: sKub *****
class SCubeTR:
    def __init__(self, root):
        self.root = root
        self.root.title("sKub")
        # Uygulama ikonunu ayarla
        try:
            self.root.iconbitmap("skub.ico")
        except Exception:
            pass # İkon yoksa hata vermeden devam et
            
        self.root.geometry("750x580")
        self.root.resizable(False, False)
        self.set_theme()

        # Türkçe tarih formatı ayarı
        try:
            locale.setlocale(locale.LC_TIME, 'tr_TR.UTF-8')
        except Exception:
            try:
                locale.setlocale(locale.LC_TIME, 'Turkish_Turkey.1254')
            except Exception:
                pass # Sessizce varsayılana dön

        self.zip_path = ""
        self.output_folder = ""
        self.temp_dir = tempfile.mkdtemp()
        
        # RAM OPTİMİZASYONU: self.logs listesi kaldırıldı.
        # Artık loglar bellekte tutulmuyor, sadece anlık ekrana basılıp unutuluyor.
        
        self.error_list = []
        self.process_running = False
        self.process_win = None
        self.max_workers = max(2, multiprocessing.cpu_count() - 1)

        self.create_widgets()

    def set_theme(self):
        """Modern görünüm için temayı ayarla"""
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("TFrame", background="#F0F7FF", padding=5)
        style.configure("TLabel", background="#F0F7FF", foreground="#333333", font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), foreground="#0063B1", background="#F0F7FF", padding=10, anchor="center")

        style.configure("TButton", font=("Segoe UI", 10), padding=8, relief="raised")
        style.map("TButton", 
                  background=[("pressed", "#DFE9F6"), ("active", "#E8F1FA")],
                  relief=[("pressed", "sunken"), ("!pressed", "raised")],
                  borderwidth=[("pressed", 2), ("!pressed", 1)])

        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=8)
        style.map("Primary.TButton",
                  background=[("pressed", "#0054A6"), ("active", "#0071D9"), ("!disabled", "#0063B1"), ("disabled", "#CCDDEE")],
                  foreground=[("pressed", "white"), ("active", "white"), ("!disabled", "white"), ("disabled", "#9CA3AF")],
                  relief=[("pressed", "sunken"), ("!pressed", "raised")],
                  borderwidth=[("pressed", 2), ("!pressed", 1)])

        style.configure("TCheckbutton", background="#F0F7FF", foreground="#333333", font=("Segoe UI", 10))
        style.map("TCheckbutton",
                  background=[("active", "#E8F1FA")],
                  indicatorcolor=[("selected", "#0063B1"), ("!selected", "#FFFFFF")])

        style.configure("TRadiobutton", background="#F0F7FF", foreground="#333333", font=("Segoe UI", 10))
        style.map("TRadiobutton",
                  background=[("active", "#E8F1FA")],
                  indicatorcolor=[("selected", "#0063B1"), ("!selected", "#FFFFFF")])

        style.configure("TProgressbar", thickness=12, background="#10B981", troughcolor="#E5E7EB", borderwidth=0)
        style.configure("Frame.TLabelframe", background="#F0F7FF", padding=10, borderwidth=1, relief="solid")
        style.configure("Frame.TLabelframe.Label", font=("Segoe UI", 11, "bold"), background="#F0F7FF", foreground="#333333")

        style.configure("TEntry", padding=6, font=("Segoe UI", 10), borderwidth=1)
        style.configure("Footer.TLabel", font=("Segoe UI", 9, "bold"), foreground="#0063B1", background="#E8F1FA", padding=5, anchor="center")
        style.configure("Footer.TFrame", background="#E8F1FA", padding=8, relief="solid", borderwidth=1)

        self.root.configure(background="#F0F7FF")

    def create_widgets(self):
        """Uygulama arayüzünü oluştur"""
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 15))
        title_label = ttk.Label(title_frame, text="sKub", style="Title.TLabel")
        title_label.pack(fill=tk.X)

        form_frame = ttk.Frame(main_frame, padding=5)
        form_frame.pack(fill=tk.X, pady=0)

        zip_frame = ttk.Frame(form_frame)
        zip_frame.pack(fill=tk.X, pady=5)
        zip_label = ttk.Label(zip_frame, text="Zip Dosyası:", width=12)
        zip_label.pack(side=tk.LEFT, padx=5)
        self.zip_entry = ttk.Entry(zip_frame, width=50)
        self.zip_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        zip_button = ttk.Button(zip_frame, text="Seç", command=self.select_zip, style="Primary.TButton", width=8)
        zip_button.pack(side=tk.RIGHT, padx=5)

        output_frame = ttk.Frame(form_frame)
        output_frame.pack(fill=tk.X, pady=5)
        output_label = ttk.Label(output_frame, text="Çıktı Klasörü:", width=12)
        output_label.pack(side=tk.LEFT, padx=5)
        self.output_entry = ttk.Entry(output_frame, width=50)
        self.output_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        output_button = ttk.Button(output_frame, text="Seç", command=self.select_output, style="Primary.TButton", width=8)
        output_button.pack(side=tk.RIGHT, padx=5)

        options_frame = ttk.LabelFrame(main_frame, text="PDF Seçenekleri", style="Frame.TLabelframe")
        options_frame.pack(fill=tk.X, pady=10, padx=5)

        self.merge_var = tk.BooleanVar(value=True)
        merge_check = ttk.Checkbutton(options_frame, text="PDF'leri Birleştir", variable=self.merge_var, command=self.toggle_sort_option)
        merge_check.pack(anchor=tk.W, padx=10, pady=5)

        sort_frame = ttk.Frame(options_frame)
        sort_frame.pack(fill=tk.X, padx=10, pady=5)
        self.sort_by_date_var = tk.BooleanVar(value=True)
        self.sort_check = ttk.Checkbutton(sort_frame, text="Faturaları Tarihe Göre Sırala", variable=self.sort_by_date_var, command=self.toggle_sort_option)
        self.sort_check.pack(anchor=tk.W)

        order_frame = ttk.Frame(options_frame)
        order_frame.pack(fill=tk.X, padx=10, pady=5)
        self.order_label = ttk.Label(order_frame, text="Sıralama Yönü:")
        self.order_label.pack(side=tk.LEFT)
        self.sort_order = tk.StringVar(value="asc")
        radio_frame = ttk.Frame(order_frame)
        radio_frame.pack(side=tk.LEFT, padx=10)
        self.asc_radio = ttk.Radiobutton(radio_frame, text="Eskiden Yeniye", variable=self.sort_order, value="asc")
        self.asc_radio.pack(side=tk.LEFT, padx=(0, 15))
        self.desc_radio = ttk.Radiobutton(radio_frame, text="Yeniden Eskiye", variable=self.sort_order, value="desc")
        self.desc_radio.pack(side=tk.LEFT)

        self.open_after_merge_var = tk.BooleanVar(value=False)
        open_after_merge_check = ttk.Checkbutton(options_frame, text="Birleştirilmiş PDF'i işlem bitince aç", variable=self.open_after_merge_var)
        open_after_merge_check.pack(anchor=tk.W, padx=10, pady=5)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        process_button = ttk.Button(button_frame, text="İşlemi Başlat", command=self.start_process_thread, style="Primary.TButton")
        process_button.pack(side=tk.RIGHT, padx=5)

        footer_frame = ttk.Frame(main_frame, style="Footer.TFrame")
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        footer_text = "Bu uygulama SMMM Arif SIRMACIOĞLU tarafından geliştirilmiştir"
        footer_contact = "Tüm hakları saklıdır. Soru ve görüşleriniz için arif@srmc.tr"
        footer_label = ttk.Label(footer_frame, text=footer_text, font=("Segoe UI", 10, "bold"), foreground="#0063B1", background="#E8F1FA", anchor="center")
        footer_label.pack(fill=tk.X)
        contact_label = ttk.Label(footer_frame, text=footer_contact, font=("Segoe UI", 9), foreground="#333333", background="#E8F1FA", anchor="center")
        contact_label.pack(fill=tk.X, pady=(0, 5))

    def show_result_in_process_window(self, result_msg, error_count):
        """İşlem sonucunu işlem penceresinde göster"""
        for widget in self.result_frame.winfo_children():
            widget.destroy()

        result_container = ttk.Frame(self.result_frame, padding=8)
        result_container.pack(fill=tk.X)
        res_label = ttk.Label(result_container, text=result_msg, font=("Segoe UI", 11, "bold"), wraplength=600, anchor="center")
        res_label.pack(pady=(5, 10), fill=tk.X)

        button_frame = ttk.Frame(result_container)
        button_frame.pack(pady=(0, 5), anchor="center")

        if error_count > 0:
            err_btn = ttk.Button(button_frame, text="Dönüştürülemeyenleri Göster",
                                 command=self.show_error_details_in_window, style="Primary.TButton")
            err_btn.pack(side=tk.LEFT, padx=(0, 10))
        done_btn = ttk.Button(button_frame, text="Tamam", command=self.close_process_window, style="Primary.TButton")
        done_btn.pack(side=tk.LEFT)

        self.process_running = False

    def show_error_details_in_window(self):
        """Hata detaylarını gösteren pencereyi oluştur"""
        details = ""
        for evrak_id, reason in self.error_list:
            details += f"Evrak No: {evrak_id}\nHata Nedeni: {reason}\n{'-'*50}\n"

        parent = self.process_win if self.process_win else self.root
        width = parent.winfo_width()
        height = parent.winfo_height()
        x = parent.winfo_x()
        y = parent.winfo_y()

        err_win = tk.Toplevel(parent)
        err_win.title("Hatalı Faturalar")
        try:
            err_win.iconbitmap("skub.ico")
        except:
            pass
        err_win.geometry(f"{width}x{height}+{x}+{y}")
        err_win.resizable(False, False)
        err_win.transient(parent)
        err_win.grab_set()

        frame = ttk.Frame(err_win, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        title_label = ttk.Label(frame, text="Hatalı Faturalar", style="Title.TLabel")
        title_label.pack(pady=(0, 10), fill=tk.X)

        text_frame = ttk.Frame(frame, padding=2, relief="solid", borderwidth=1)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        txt = tk.Text(text_frame, wrap=tk.WORD, font=("Consolas", 9))
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        txt.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=txt.yview)
        txt.insert(tk.END, details)
        txt.configure(state='disabled')

        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, pady=(0, 5))
        close_btn = ttk.Button(button_frame, text="Kapat", command=err_win.destroy, style="Primary.TButton")
        close_btn.pack(side=tk.RIGHT, pady=5, padx=5)

        footer_frame = ttk.Frame(frame, style="Footer.TFrame")
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        footer_text = "Bu uygulama SMMM Arif SIRMACIOĞLU tarafından geliştirilmiştir"
        footer_contact = "Tüm hakları saklıdır. Soru ve görüşleriniz için arif@srmc.tr"
        footer_label = ttk.Label(footer_frame, text=footer_text, font=("Segoe UI", 10, "bold"),
                                 foreground="#0063B1", background="#E8F1FA", anchor="center")
        footer_label.pack(fill=tk.X)
        contact_label = ttk.Label(footer_frame, text=footer_contact, font=("Segoe UI", 9),
                                  foreground="#333333", background="#E8F1FA", anchor="center")
        contact_label.pack(fill=tk.X, pady=(0, 5))

    def close_process_window(self):
        """İşlem penceresini kapat"""
        if self.process_win:
            self.process_win.grab_release()
            self.process_win.destroy()
            self.process_win = None

    def on_closing(self):
        """Uygulama kapatılırken çağrılır"""
        try:
            shutil.rmtree(self.temp_dir)
        except Exception:
            pass
        self.root.destroy()

    def process_files_thread(self):
        """Dosyaları işleyen ana iş parçacığı"""
        try:
            # RAM temizliği: Logs listesi olmadığı için temizlemeye gerek yok.
            self.error_list.clear()
            self.update_proc_status("İşlem başlatılıyor...", 0)

            # Temp klasörünü temizle
            for file in os.listdir(self.temp_dir):
                path = os.path.join(self.temp_dir, file)
                if os.path.isfile(path):
                    os.unlink(path)
                elif os.path.isdir(path):
                    shutil.rmtree(path)

            self.update_proc_status("Zip dosyası açılıyor...", 10)
            extract_dir = os.path.join(self.temp_dir, "extracted")
            os.makedirs(extract_dir, exist_ok=True)

            processor = InvoiceProcessor(self.log_message)
            processor.extract_zip_recursively(self.zip_path, extract_dir)

            self.update_proc_status("Dosyalar aranıyor...", 30)
            html_files = processor.find_files(extract_dir, ['.html', '.htm'])
            xml_files = processor.find_files(extract_dir, ['.xml'])

            if not html_files:
                self.root.after(0, lambda: messagebox.showerror("Hata", "Hiçbir HTML dosyası bulunamadı."))
                self.finish_process()
                return

            self.update_proc_status(f"Bulunan HTML dosyası sayısı: {len(html_files)}", 30)
            self.log_message(f"Bulunan HTML: {len(html_files)}  |  XML: {len(xml_files)}")

            self.update_proc_status("Fatura tarihleri ve evrak numaraları tespit ediliyor...", 40)
            html_files_with_dates = processor.match_html_with_xml(html_files, xml_files)

            wkhtmltopdf_path = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
            if not os.path.exists(wkhtmltopdf_path):
                self.update_proc_status("wkhtmltopdf bulunamadı. Sistem yolunu kontrol etme...", 40)
                try:
                    config = pdfkit.configuration()
                    self.log_message("wkhtmltopdf sistem yolunda bulundu.")
                except Exception:
                    self.root.after(0, lambda: messagebox.showerror("Hata", 
                        "wkhtmltopdf bulunamadı. Lütfen https://wkhtmltopdf.org/downloads.html adresinden indirip kurun."))
                    self.finish_process()
                    return
            else:
                config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)
                self.update_proc_status(f"wkhtmltopdf bulundu: {wkhtmltopdf_path}", 40)

            pdf_options = {
                "enable-local-file-access": "",
                "encoding": "UTF-8",
                "page-size": "A4",
                "margin-top": "10mm",
                "margin-right": "10mm",
                "margin-bottom": "10mm",
                "margin-left": "10mm"
            }

            self.update_proc_status("HTML dosyaları PDF'e dönüştürülüyor...", 50)
            pdf_files_with_info, conversion_errors = processor.convert_html_to_pdf_parallel(
                html_files_with_dates,
                self.temp_dir,
                config,
                pdf_options,
                self.update_proc_status
            )

            self.error_list.extend(conversion_errors)
            pdf_files = [p for (p, _, _) in pdf_files_with_info if p is not None]

            if not pdf_files:
                self.root.after(0, lambda: messagebox.showerror("Hata", "Hiçbir PDF dosyası oluşturulamadı."))
                self.finish_process()
                return

            if self.merge_var.get() and len(pdf_files) > 1:
                self.update_proc_status("PDF dosyaları birleştiriliyor...", 80)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                if self.sort_by_date_var.get():
                    order_str = "eskiden_yeniye" if self.sort_order.get() == "asc" else "yeniden_eskiye"
                    merged_name = f"birlesik_faturalar_{order_str}_{ts}.pdf"
                else:
                    merged_name = f"birlesik_faturalar_{ts}.pdf"
                merged_path = os.path.join(self.output_folder, merged_name)
                cnt = 1
                while os.path.exists(merged_path):
                    merged_name = f"birlesik_faturalar_{ts}_{cnt}.pdf"
                    merged_path = os.path.join(self.output_folder, merged_name)
                    cnt += 1

                merger = PdfMerger()
                merge_success_count = 0
                merge_error_count = 0

                if self.sort_by_date_var.get():
                    sorted_files = sorted(
                        [(pdf, date, eid) for pdf, date, eid in pdf_files_with_info if pdf is not None],
                        key=lambda x: x[1] if x[1] else datetime.min,
                        reverse=(self.sort_order.get() == "desc")
                    )
                    pdf_files_to_merge = [item[0] for item in sorted_files]
                else:
                    pdf_files_to_merge = pdf_files

                for pdf in pdf_files_to_merge:
                    try:
                        merger.append(pdf)
                        merge_success_count += 1
                    except Exception as e:
                        self.log_message(f"⚠️ Birleştirme hatası: {os.path.basename(pdf)} - {str(e)}")
                        merge_error_count += 1

                if merge_success_count > 0:
                    try:
                        merger.write(merged_path)
                        merger.close()
                        self.update_proc_status(f"Birleştirilmiş PDF kaydedildi: {os.path.basename(merged_path)}", 100)
                        self.log_message(f"Kayıt konumu: {merged_path}")
                        if self.open_after_merge_var.get():
                            try:
                                os.startfile(merged_path)
                            except Exception:
                                subprocess.Popen([merged_path], shell=True)
                        total_errors = len(self.error_list) + merge_error_count
                        result_msg = f"{merge_success_count} fatura birleştirildi ve kaydedildi."
                        if merge_error_count > 0:
                            result_msg += f" ({merge_error_count} fatura birleştirilemedi)"
                        if len(self.error_list) > 0:
                            result_msg += f" ({len(self.error_list)} fatura dönüştürülemedi)"
                        self.root.after(0, lambda: self.show_result_in_process_window(result_msg, total_errors))
                    except Exception as e:
                        self.root.after(0, lambda: messagebox.showerror("Hata", f"PDF birleştirilirken hata: {str(e)}"))
                        self.finish_process()
                else:
                    self.root.after(0, lambda: messagebox.showerror("Hata", "Hiçbir PDF birleştirilemedi."))
                    self.finish_process()
            else:
                self.update_proc_status("PDF dosyaları kopyalanıyor...", 80)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                success_count = 0
                output_sub = os.path.join(self.output_folder, f"faturalar_{ts}")
                cnt = 1
                while os.path.exists(output_sub):
                    output_sub = os.path.join(self.output_folder, f"faturalar_{ts}_{cnt}")
                    cnt += 1
                os.makedirs(output_sub, exist_ok=True)

                for i, (pdf, invoice_date, evrak_id) in enumerate(pdf_files_with_info):
                    if pdf is None:
                        continue
                    if evrak_id:
                        target_name = f"{evrak_id}.pdf"
                    else:
                        if invoice_date:
                            dstr = invoice_date.strftime("%Y%m%d")
                            target_name = f"fatura_{dstr}_{i+1}.pdf"
                        else:
                            target_name = f"fatura_{i+1}.pdf"
                    target_path = os.path.join(output_sub, target_name)
                    base, ext = os.path.splitext(target_name)
                    cdup = 1
                    while os.path.exists(target_path):
                        target_name = f"{base}_{cdup}{ext}"
                        target_path = os.path.join(output_sub, target_name)
                        cdup += 1
                    try:
                        shutil.copy2(pdf, target_path)
                        self.log_message(f"✓ Kaydedildi: {os.path.basename(target_path)}")
                        success_count += 1
                    except Exception as e:
                        self.log_message(f"✗ Kaydetme hatası: {os.path.basename(target_path)} - {str(e)}")
                        self.error_list.append((evrak_id if evrak_id else "Bilinmiyor", f"Dosya kopyalama hatası: {str(e)}"))
                self.update_proc_status(f"{success_count} PDF dosyası kaydedildi.", 100)
                self.log_message(f"Kayıt konumu: {output_sub}")
                result_msg = f"{success_count} fatura PDF'e dönüştürüldü ve kaydedildi."
                if len(self.error_list) > 0:
                    result_msg += f" ({len(self.error_list)} fatura dönüştürülemedi)"
                self.root.after(0, lambda: self.show_result_in_process_window(result_msg, len(self.error_list)))

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Hata", f"İşlem sırasında hata:\n{str(e)}"))
            self.update_proc_status("Hata oluştu!", 0)
            self.log_message(f"HATA: {str(e)}")
            self.log_message(traceback.format_exc())
            self.finish_process()

    def finish_process(self):
        """İşlem tamamlandığında çağrılır"""
        self.process_running = False

    def start_process_thread(self):
        """Dosya işleme iş parçacığını başlat"""
        if self.process_running:
            return

        if not self.zip_path:
            messagebox.showerror("Hata", "Lütfen bir zip dosyası seçin.")
            return

        if not self.output_folder:
            messagebox.showerror("Hata", "Lütfen bir çıktı klasörü seçin.")
            return

        if not os.path.exists(self.zip_path):
            messagebox.showerror("Hata", "Seçilen zip dosyası bulunamadı.")
            return

        if not os.path.exists(self.output_folder):
            try:
                os.makedirs(self.output_folder)
            except Exception as e:
                messagebox.showerror("Hata", f"Çıktı klasörü oluşturulamadı: {str(e)}")
                return

        self.process_running = True
        self.create_process_window()
        t = threading.Thread(target=self.process_files_thread)
        t.daemon = True
        t.start()

    def create_process_window(self):
        """İşlem durumunu gösteren pencereyi oluştur"""
        self.process_win = tk.Toplevel(self.root)
        self.process_win.title("İşlem Durumu")
        try:
            self.process_win.iconbitmap("skub.ico")
        except:
            pass
        self.process_win.geometry(self.root.geometry())
        x = self.root.winfo_x()
        y = self.root.winfo_y()
        self.process_win.geometry("+%d+%d" % (x, y))
        self.process_win.resizable(False, False)
        self.process_win.transient(self.root)
        self.process_win.grab_set()
        self.process_win.protocol("WM_DELETE_WINDOW", lambda: None)

        self.proc_frame = ttk.Frame(self.process_win, padding=10)
        self.proc_frame.pack(fill=tk.BOTH, expand=True)
        self.proc_top = ttk.Frame(self.proc_frame)
        self.proc_top.pack(fill=tk.X, padx=5)
        self.proc_progress = ttk.Progressbar(self.proc_top, maximum=100, style="TProgressbar")
        self.proc_progress.pack(fill=tk.X, pady=(5, 10))
        self.proc_status_label = ttk.Label(self.proc_top, text="İşlem Başlatılıyor...", font=("Segoe UI", 11, "bold"), anchor="center")
        self.proc_status_label.pack(pady=(0, 10), fill=tk.X)
        self.proc_middle = ttk.Frame(self.proc_frame)
        self.proc_middle.pack(fill=tk.BOTH, expand=True, padx=5)
        text_frame = ttk.Frame(self.proc_middle, padding=2, relief="solid", borderwidth=1)
        text_frame.pack(fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.proc_text = tk.Text(text_frame, wrap=tk.WORD, state='disabled', font=("Consolas", 9))
        self.proc_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.proc_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.proc_text.yview)
        self.proc_bottom = ttk.Frame(self.proc_frame)
        self.proc_bottom.pack(fill=tk.X, pady=(10, 0))
        self.result_frame = ttk.Frame(self.proc_bottom)
        self.result_frame.pack(fill=tk.X)
        footer_frame = ttk.Frame(self.proc_bottom, style="Footer.TFrame")
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        footer_text = "Bu uygulama SMMM Arif SIRMACIOĞLU tarafından geliştirilmiştir"
        footer_contact = "Tüm hakları saklıdır. Soru ve görüşleriniz için arif@srmc.tr"
        footer_label = ttk.Label(footer_frame, text=footer_text, font=("Segoe UI", 10, "bold"),
                                 foreground="#0063B1", background="#E8F1FA", anchor="center")
        footer_label.pack(fill=tk.X)
        contact_label = ttk.Label(footer_frame, text=footer_contact, font=("Segoe UI", 9),
                                  foreground="#333333", background="#E8F1FA", anchor="center")
        contact_label.pack(fill=tk.X, pady=(0, 5))

    def toggle_sort_option(self):
        """PDF birleştirme seçeneğine göre sıralama seçeneklerini etkinleştir/devre dışı bırak"""
        if self.merge_var.get():
            self.sort_check.configure(state="normal")
            if self.sort_by_date_var.get():
                self.order_label.configure(state="normal")
                self.asc_radio.configure(state="normal")
                self.desc_radio.configure(state="normal")
            else:
                self.order_label.configure(state="disabled")
                self.asc_radio.configure(state="disabled")
                self.desc_radio.configure(state="disabled")
        else:
            self.sort_check.configure(state="disabled")
            self.order_label.configure(state="disabled")
            self.asc_radio.configure(state="disabled")
            self.desc_radio.configure(state="disabled")

    def select_zip(self):
        """Zip dosyası seçimi için dosya iletişim kutusunu göster"""
        self.zip_path = filedialog.askopenfilename(
            title="Zip Dosyasını Seçin",
            filetypes=[("Zip Dosyaları", "*.zip")],
            initialdir=os.path.join(os.path.expanduser("~"), "Desktop")
        )
        if self.zip_path:
            self.zip_entry.delete(0, tk.END)
            self.zip_entry.insert(0, self.zip_path)

    def select_output(self):
        """Çıktı klasörü seçimi için klasör iletişim kutusunu göster"""
        self.output_folder = filedialog.askdirectory(
            title="Çıktı Klasörü Seçin",
            initialdir=os.path.join(os.path.expanduser("~"), "Desktop")
        )
        if self.output_folder:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, self.output_folder)

    def log_message(self, message):
        """
        Log mesajını sadece ekrana basar.
        RAM'de tutmaz (self.logs kaldırıldı).
        Diske yazmaz (Logger kaldırıldı).
        """
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {message}"
        
        # self.logs.append(line)  <-- RAM TASARRUFU İÇİN SİLİNDİ
        # logger.info(message)    <-- DİSK/CONSOLE LOGU İÇİN SİLİNDİ

        # Sadece kullanıcı arayüzünü güncelle (Geçici görselleştirme)
        if self.process_win and hasattr(self, 'proc_text'):
            try:
                self.proc_text.configure(state='normal')
                self.proc_text.insert(tk.END, line + "\n")
                self.proc_text.see(tk.END)
                self.proc_text.configure(state='disabled')
            except Exception:
                pass

    def update_proc_status(self, message, progress=None):
        """İşlem durumunu ve ilerlemeyi güncelle"""
        if self.process_win:
            if progress is not None:
                self.proc_progress['value'] = progress
            self.proc_status_label.config(text=message)
            self.log_message(message)


if __name__ == "__main__":
    root = tk.Tk()
    app = SCubeTR(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()