import sys
import time
import threading
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QLabel, QLineEdit, QSpinBox, QTextEdit, QPushButton, QTabWidget, 
                            QGroupBox, QFormLayout, QProgressBar, QMessageBox, QCheckBox,
                            QComboBox, QScrollArea, QListWidget, QListWidgetItem, QDialog,
                            QSplitter, QFrame, QSizePolicy, QToolButton, QGridLayout)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, pyqtSlot, QSize, QPropertyAnimation, QEasingCurve
from PyQt5.QtGui import QFont, QIcon, QPixmap, QColor, QPalette, QTextCursor, QTextCharFormat
import qdarkstyle

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

class CourseSelectionDialog(QDialog):
    def __init__(self, courses, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Pilih Mata Kuliah")
        self.setMinimumSize(600, 400)
        self.selected_courses = []
        
        # Apply light mode style
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f7;
            }
            QLabel {
                color: #333333;
            }
            QCheckBox {
                color: #333333;
                spacing: 8px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 4px;
                border: 1px solid #b3d4fc;
            }
            QCheckBox::indicator:unchecked {
                background-color: #ffffff;
            }
            QCheckBox::indicator:checked {
                background-color: #2c75b3;
            }
            QPushButton {
                border-radius: 8px;
                font-weight: bold;
                padding: 8px 12px;
                background-color: #2c75b3;
                color: white;
                border: none;
            }
            QPushButton:hover {
                background-color: #4a90e2;
            }
            QPushButton:pressed {
                background-color: #1b5493;
            }
            QScrollArea {
                border: 1px solid #b3d4fc;
                border-radius: 5px;
                background-color: #ffffff;
            }
        """)
        
        layout = QVBoxLayout(self)
        
        # Label instructions
        label = QLabel("Pilih mata kuliah yang ingin diisi EPBM-nya:")
        layout.addWidget(label)
        
        # Course list with checkboxes
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_area.setWidget(scroll_widget)
        scroll_layout = QVBoxLayout(scroll_widget)
        
        # Create checkboxes for each course
        self.checkboxes = []
        for i, course in enumerate(courses):
            # For completed courses, show them as disabled with info
            if course.get('is_completed', False):
                checkbox = QCheckBox(f"{course['title']}: {course['desc']} [Sudah Diisi]")
                checkbox.setChecked(False)
                checkbox.setEnabled(False)
                checkbox.setStyleSheet("color: #999999;")
                info_label = QLabel("  ↳ EPBM untuk mata kuliah ini sudah diisi sebelumnya.")
                info_label.setStyleSheet("color: #999999; font-style: italic;")
                scroll_layout.addWidget(checkbox)
                scroll_layout.addWidget(info_label)
            else:
                checkbox = QCheckBox(f"{course['title']}: {course['desc']}")
                checkbox.setChecked(True)  # Default select all unfinished courses
                checkbox.course_data = course
                scroll_layout.addWidget(checkbox)
                self.checkboxes.append(checkbox)
            
        layout.addWidget(scroll_area)
        
        # Quick selection buttons
        buttons_layout = QHBoxLayout()
        
        select_all_btn = QPushButton("Pilih Semua")
        select_all_btn.clicked.connect(self.select_all)
        buttons_layout.addWidget(select_all_btn)
        
        deselect_all_btn = QPushButton("Batalkan Semua")
        deselect_all_btn.clicked.connect(self.deselect_all)
        buttons_layout.addWidget(deselect_all_btn)
        
        layout.addLayout(buttons_layout)
        
        # OK and Cancel buttons
        button_box = QHBoxLayout()
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton("Batal")
        cancel_button.clicked.connect(self.reject)
        
        button_box.addWidget(ok_button)
        button_box.addWidget(cancel_button)
        layout.addLayout(button_box)
        
    def select_all(self):
        for checkbox in self.checkboxes:
            checkbox.setChecked(True)
            
    def deselect_all(self):
        for checkbox in self.checkboxes:
            checkbox.setChecked(False)
            
    def get_selected_courses(self):
        selected = []
        for checkbox in self.checkboxes:
            if checkbox.isChecked():
                selected.append(checkbox.course_data)
        return selected

class CourseFinderWorker(QThread):
    update_signal = pyqtSignal(str)
    courses_found_signal = pyqtSignal(list)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(self, credentials):
        super().__init__()
        self.credentials = credentials
        
    def run(self):
        try:
            courses = self.find_courses()
            if courses:
                self.courses_found_signal.emit(courses)
                self.finished_signal.emit(True, f"Ditemukan {len(courses)} mata kuliah.")
            else:
                self.finished_signal.emit(False, "Tidak ditemukan mata kuliah yang perlu diisi EPBM.")
        except Exception as e:
            self.finished_signal.emit(False, f"Error: {str(e)}")
            
    def log(self, message):
        self.update_signal.emit(message)
            
    def find_courses(self):
        username = self.credentials['username']
        password = self.credentials['password']
        
        # Setup Chrome driver
        chrome_options = Options()
        chrome_options.add_argument("--headless")  # Always use headless for scanning
        chrome_options.add_argument("--start-maximized")
        
        # Initialize the Chrome driver
        self.log("Menginisialisasi Chrome driver...")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        
        try:
            # Open the IPB student portal
            self.log("Membuka portal mahasiswa IPB...")
            driver.get("https://studentportal.ipb.ac.id/Akademik/EPBM/Detail")
            
            # Check if login page is displayed and handle login
            if "login" in driver.current_url.lower() or driver.find_elements(By.ID, "Username"):
                self.log("Halaman login terdeteksi, melakukan login...")
                
                # Wait for username field to be visible
                username_field = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "Username"))
                )
                username_field.clear()
                username_field.send_keys(username)
                
                # Enter password
                password_field = driver.find_element(By.ID, "Password")
                password_field.clear()
                password_field.send_keys(password)
                
                # Click login button
                login_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
                login_button.click()
                
                # Wait for login to complete
                self.log("Menunggu proses login selesai...")
                time.sleep(3)
                
                # Check if login failed
                error_alerts = driver.find_elements(By.CSS_SELECTOR, ".alert.alert-danger")
                for alert in error_alerts:
                    if "Login gagal" in alert.text or "password Anda salah" in alert.text:
                        self.log("Login gagal: Username atau password salah.")
                        self.finished_signal.emit(False, "Login gagal: Username atau password Anda salah. Silakan periksa kembali.")
                        return []
            
            # Check if we're still on the login page after clicking login button
            if "login" in driver.current_url.lower():
                self.log("Masih berada di halaman login. Kemungkinan username atau password salah.")
                self.finished_signal.emit(False, "Login gagal: Gagal masuk ke portal. Silakan periksa kembali username dan password.")
                return []
            
            # Wait for page to load after login
            self.log("Menunggu halaman EPBM dimuat...")
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".btn.card.small-box"))
            )
            
            # Get all EPBM cards
            epbm_cards = driver.find_elements(By.CSS_SELECTOR, ".btn.card.small-box")
            self.log(f"Ditemukan {len(epbm_cards)} kartu EPBM yang perlu diisi.")
            
            # Parse course information
            courses = []
            for i, card in enumerate(epbm_cards):
                try:
                    # Get card information
                    header = card.find_element(By.CSS_SELECTOR, ".card-header")
                    href = card.get_attribute("href")
                    
                    # Get course information
                    card_title = "Untitled"
                    card_desc = ""
                    
                    try:
                        card_title = header.find_element(By.TAG_NAME, "h4").text
                    except:
                        pass
                    
                    try:
                        card_texts = header.find_elements(By.TAG_NAME, "p")
                        if card_texts:
                            card_desc = card_texts[0].text
                    except:
                        pass
                    
                    # Detect if this is sarpras
                    is_sarpras = "Sarana dan Prasarana" in card_title or "sarpras" in href.lower()
                    
                    # Check if the EPBM form has already been filled (look for the check icon)
                    is_completed = False
                    try:
                        # Find check icon inside card
                        check_icons = card.find_elements(By.CSS_SELECTOR, ".fa-check-circle.text-success")
                        if check_icons:
                            is_completed = True
                            self.log(f"Terdeteksi {card_title}: {card_desc} sudah diisi sebelumnya.")
                    except:
                        pass
                    
                    # Add to courses list
                    courses.append({
                        'title': card_title,
                        'desc': card_desc,
                        'href': href,
                        'is_sarpras': is_sarpras,
                        'is_completed': is_completed,
                        'index': i
                    })
                    
                    self.log(f"Ditemukan: {card_title}: {card_desc}{' (Sudah diisi)' if is_completed else ''}")
                    
                except Exception as e:
                    self.log(f"Error parsing course card: {e}")
            
            return courses
            
        except Exception as e:
            self.log(f"Terjadi error: {e}")
            return []
        finally:
            # Close the browser
            driver.quit()

class EPBMAutomationWorker(QThread):
    update_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(self, credentials, settings, selected_courses):
        super().__init__()
        self.credentials = credentials
        self.settings = settings
        self.selected_courses = selected_courses
        
    def run(self):
        try:
            self.fill_epbm_portal()
            self.finished_signal.emit(True, "Otomasi selesai dengan sukses!")
        except Exception as e:
            self.finished_signal.emit(False, f"Error: {str(e)}")
            
    def log(self, message):
        self.update_signal.emit(message)
            
    def fill_epbm_portal(self):
        username = self.credentials['username']
        password = self.credentials['password']
        
        # Setup Chrome driver
        chrome_options = Options()
        if self.settings['headless']:
            chrome_options.add_argument("--headless")
        chrome_options.add_argument("--start-maximized")
        
        # Initialize the Chrome driver
        self.log("Menginisialisasi Chrome driver...")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        
        try:
            # Open the IPB student portal
            self.log("Membuka portal mahasiswa IPB...")
            driver.get("https://studentportal.ipb.ac.id/Akademik/EPBM/Detail")
            
            # Check if login page is displayed and handle login
            if "login" in driver.current_url.lower() or driver.find_elements(By.ID, "Username"):
                self.log("Halaman login terdeteksi, melakukan login...")
                
                # Wait for username field to be visible
                username_field = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "Username"))
                )
                username_field.clear()
                username_field.send_keys(username)
                
                # Enter password
                password_field = driver.find_element(By.ID, "Password")
                password_field.clear()
                password_field.send_keys(password)
                
                # Click login button
                login_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
                login_button.click()
                
                # Wait for login to complete
                self.log("Menunggu proses login selesai...")
                time.sleep(3)
                
                # Check if login failed
                error_alerts = driver.find_elements(By.CSS_SELECTOR, ".alert.alert-danger")
                for alert in error_alerts:
                    if "Login gagal" in alert.text or "password Anda salah" in alert.text:
                        self.log("Login gagal: Username atau password salah.")
                        self.finished_signal.emit(False, "Login gagal: Username atau password Anda salah. Silakan periksa kembali.")
                        return
                
                # Check if we're still on the login page after clicking login button
                if "login" in driver.current_url.lower():
                    self.log("Masih berada di halaman login. Kemungkinan username atau password salah.")
                    self.finished_signal.emit(False, "Login gagal: Gagal masuk ke portal. Silakan periksa kembali username dan password.")
                    return
            
            # Wait for page to load after login
            self.log("Menunggu halaman EPBM dimuat...")
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".btn.card.small-box"))
            )
            
            # Get all EPBM cards that match our selected courses
            epbm_cards = driver.find_elements(By.CSS_SELECTOR, ".btn.card.small-box")
            total_cards = len(self.selected_courses)
            
            self.log(f"Akan mengisi {total_cards} kartu EPBM yang dipilih.")
            
            # Initialize progress tracking
            total_steps = total_cards * 5  # Each course has about 5 steps (click, fill, checkbox, save, return)
            current_step = 0
            
            # Update initial progress
            self.progress_signal.emit(0)
            
            # Set page load timeout to prevent hanging and make it faster
            driver.set_page_load_timeout(15)  # Reduced from 30 to 15 seconds
            
            # Loop through each selected course
            for i, course in enumerate(self.selected_courses):
                try:
                    # Card found - update progress
                    current_step += 1
                    progress = min(int((current_step / total_steps) * 100), 99)  # Keep under 100% until fully complete
                    self.progress_signal.emit(progress)
                    
                    # Find the course card again (in case the page was refreshed)
                    epbm_cards = driver.find_elements(By.CSS_SELECTOR, ".btn.card.small-box")
                    
                    if 0 <= course['index'] < len(epbm_cards):
                        card = epbm_cards[course['index']]
                    else:
                        self.log(f"Error: Kartu dengan indeks {course['index']} tidak ditemukan. Melewati...")
                        continue
                    
                    self.log(f"\nMengisi EPBM untuk {course['title']}: {course['desc']} ({i+1}/{total_cards})")
                    
                    # Click on the card
                    try:
                        card.click()
                    except:
                        # Sometimes the click fails - use JavaScript if that happens
                        self.log("Menggunakan JavaScript untuk mengklik kartu...")
                        driver.execute_script("arguments[0].click();", card)
                        
                    # Wait with shorter timeout for better performance
                    time.sleep(0.5)  # Reduced from 1 second to 0.5 seconds
                    
                    # Clicked card - update progress
                    current_step += 1
                    progress = min(int((current_step / total_steps) * 100), 99)
                    self.progress_signal.emit(progress)
                    
                    # Check if this is the Sarana Prasarana form
                    is_sarpras = course['is_sarpras']
                    
                    if is_sarpras:
                        # Handle Sarana Prasarana form directly (single page)
                        self.log("Mengisi kuesioner Sarana dan Prasarana...")
                        # Fill star ratings
                        try:
                            star_ratings = WebDriverWait(driver, 3).until(  # Reduced from 5 to 3 seconds
                                EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".b-rating"))
                            )
                            
                            for j, rating in enumerate(star_ratings):
                                # Determine which rating to use based on the question (1-3)
                                if j == 0:  # Kenyamanan kelas
                                    star_value = self.settings['sarpras_kenyamanan']
                                elif j == 1:  # Fasilitas internet
                                    star_value = self.settings['sarpras_internet']
                                elif j == 2:  # Toilet
                                    star_value = self.settings['sarpras_toilet']
                                else:
                                    star_value = 4  # Default
                                    
                                # Get all stars in the current rating
                                stars = rating.find_elements(By.CSS_SELECTOR, ".b-rating-star")
                                
                                # Click on the star based on the rating (1-4)
                                if len(stars) >= star_value and star_value > 0:
                                    try:
                                        stars[star_value - 1].click()
                                    except:
                                        # Use JavaScript if regular click fails
                                        driver.execute_script("arguments[0].click();", stars[star_value - 1])
                                    
                                    # Reduce delay between clicks for faster processing
                                    time.sleep(0.1)  # Reduced from 0.2 to 0.1 seconds
                        except Exception as e:
                            self.log(f"Error pada pengisian rating: {str(e)}")
                        
                        # After filling stars - update progress
                        current_step += 1
                        progress = min(int((current_step / total_steps) * 100), 99)
                        self.progress_signal.emit(progress)
                        
                        # Find and click checkbox if it exists
                        try:
                            # Wait for checkboxes to be available
                            checkboxes = WebDriverWait(driver, 3).until(
                                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input[type='checkbox']"))
                            )
                            
                            for checkbox in checkboxes:
                                try:
                                    if not checkbox.is_selected():
                                        # Try direct click first
                                        try:
                                            checkbox.click()
                                        except:
                                            # If that fails, try JavaScript click
                                            driver.execute_script("arguments[0].click();", checkbox)
                                            
                                        self.log("Mengklik checkbox pernyataan...")
                                except Exception as e:
                                    self.log(f"Gagal mengklik checkbox: {e}")
                        except Exception as e:
                            self.log(f"Error pada checkbox: {str(e)}")
                        
                        # After clicking checkbox - update progress
                        current_step += 1
                        progress = min(int((current_step / total_steps) * 100), 99)
                        self.progress_signal.emit(progress)
                        
                        # Click simpan EPBM button
                        try:
                            # Find and click the save button
                            save_buttons = WebDriverWait(driver, 3).until(
                                EC.presence_of_all_elements_located((By.XPATH, "//button[contains(text(), 'Simpan EPBM')]"))
                            )
                            
                            if save_buttons:
                                self.log("Menyimpan EPBM Sarana dan Prasarana...")
                                # Try normal click first
                                try:
                                    save_buttons[0].click()
                                except:
                                    # If that fails, try JavaScript click
                                    driver.execute_script("arguments[0].click();", save_buttons[0])
                                
                                # Handling modals that may appear after saving
                                try:
                                    # Look for modal dialog or success message
                                    WebDriverWait(driver, 3).until(
                                        EC.presence_of_element_located((By.CLASS_NAME, "modal-dialog"))
                                    )
                                    self.log("Dialog modal terdeteksi setelah menyimpan")
                                    
                                    # Find and click OK/Close button on modal
                                    modal_buttons = driver.find_elements(By.CSS_SELECTOR, ".modal-footer button, .modal button.btn, .modal .close")
                                    if modal_buttons:
                                        try:
                                            modal_buttons[0].click()
                                            self.log("Mengklik tombol pada modal dialog")
                                        except:
                                            driver.execute_script("arguments[0].click();", modal_buttons[0])
                                except:
                                    # No modal found or timeout, that's fine
                                    pass
                                
                                # Navigate back to main page
                                try:
                                    driver.get("https://studentportal.ipb.ac.id/Akademik/EPBM/Detail")
                                    WebDriverWait(driver, 8).until(
                                        EC.presence_of_element_located((By.CSS_SELECTOR, ".btn.card.small-box"))
                                    )
                                    self.log("EPBM Sarana dan Prasarana berhasil disimpan!")
                                except:
                                    # If navigate fails, try once more
                                    driver.get("https://studentportal.ipb.ac.id/Akademik/EPBM/Detail")
                                    time.sleep(2)
                                    self.log("Kembali ke halaman utama setelah menyimpan")
                        except Exception as e:
                            self.log(f"Error ketika menyimpan data (akan mencoba kembali ke halaman utama): {str(e)}")
                            try:
                                driver.get("https://studentportal.ipb.ac.id/Akademik/EPBM/Detail")
                                time.sleep(2)
                            except:
                                self.log("Gagal kembali ke halaman utama setelah error")
                        
                        # After saving - update progress for this course completion
                        current_step += 1
                        progress = min(int((current_step / total_steps) * 100), 99)
                        self.progress_signal.emit(progress)
                    
                    else:
                        # Regular mata kuliah EPBM
                        # Wait for the form page to load with a shorter timeout
                        try:
                            WebDriverWait(driver, 3).until(  # Reduced from 5 to 3 seconds
                                EC.presence_of_element_located((By.CSS_SELECTOR, ".b-rating"))
                            )
                            self.log("Halaman form EPBM telah dimuat.")
                        except:
                            self.log("Timeout pada loading halaman, mencoba melanjutkan...")
                        
                        # Process all form pages with progress updates
                        page_count = 0
                        while True:
                            # Check the current page heading
                            try:
                                page_headings = driver.find_elements(By.TAG_NAME, "h5")
                                current_page = ""
                                for heading in page_headings:
                                    heading_text = heading.text
                                    current_page = heading_text
                                    self.log(f"Mengisi halaman: {heading_text}")
                                    break
                            except:
                                self.log("Tidak dapat menemukan heading halaman")
                            
                            # Fill star ratings (if available)
                            try:
                                star_ratings = driver.find_elements(By.CSS_SELECTOR, ".b-rating")
                                if star_ratings:
                                    self.log(f"Mengisi {len(star_ratings)} pertanyaan...")
                                    
                                    for j, rating in enumerate(star_ratings):
                                        # Determine which rating to use based on the page
                                        star_value = 4  # Default
                                        
                                        if "1. Pertanyaan terkait mata kuliah" in current_page:
                                            if j == 0:  # Proses pembelajaran sesuai harapan
                                                star_value = self.settings['matkul_sesuai_harapan']
                                            elif j == 1:  # Proses pembelajaran menyenangkan
                                                star_value = self.settings['matkul_menyenangkan']
                                            elif j == 2:  # Proses asesmen terbuka
                                                star_value = self.settings['matkul_asesmen']
                                            elif j == 3:  # Kesempatan hardskill/softskill
                                                star_value = self.settings['matkul_hardskill']
                                            elif j == 4:  # Dokumen ajar
                                                star_value = self.settings['matkul_dokumen']
                                        elif "2. Dosen memberikan kuliah dengan metode ceramah" in current_page:
                                            star_value = self.settings['dosen_ceramah']
                                        elif "3. Dosen menyampaikan kuliah dengan menjadi mentor" in current_page:
                                            star_value = self.settings['dosen_mentor']
                                        elif "4. Dosen memberikan contoh/ilustrasi" in current_page:
                                            star_value = self.settings['dosen_ilustrasi']
                                        elif "5. Dosen menfaatkan ketersediaan teknologi" in current_page:
                                            star_value = self.settings['dosen_teknologi']
                                        elif "6. Dosen memberikan umpan balik" in current_page:
                                            star_value = self.settings['dosen_feedback']
                                        
                                        # Get all stars in the current rating
                                        stars = rating.find_elements(By.CSS_SELECTOR, ".b-rating-star")
                                        
                                        # Click on the star based on the rating (1-4)
                                        if len(stars) >= star_value and star_value > 0:
                                            try:
                                                stars[star_value - 1].click()
                                            except:
                                                # Use JavaScript if click fails
                                                try:
                                                    driver.execute_script("arguments[0].click();", stars[star_value - 1])
                                                except:
                                                    self.log("Gagal mengklik rating")
                                                    
                                            # Reduce delay between clicks for faster processing
                                            time.sleep(0.1)  # Reduced from 0.2 to 0.1 seconds
                            except Exception as e:
                                self.log(f"Error pada pengisian rating: {str(e)}")
                            
                            # After filling a page - update progress
                            if page_count % 2 == 0:  # Update every other page to avoid too frequent updates
                                current_step += 1
                                progress = min(int((current_step / total_steps) * 100), 99)
                                self.progress_signal.emit(progress)
                            
                            page_count += 1
                            
                            # Check if we're on the saran page
                            if "7. Berikan saran untuk masing-masing dosen pengajar" in current_page:
                                # Fill suggestion text areas
                                try:
                                    text_areas = driver.find_elements(By.TAG_NAME, "textarea")
                                    for text_area in text_areas:
                                        try:
                                            text_area.clear()  # Clear first to ensure text is properly entered
                                            text_area.send_keys(self.settings['saran_dosen'])
                                            self.log("Mengisi saran untuk dosen...")
                                        except Exception as e:
                                            # Try JavaScript alternative if direct input fails
                                            try:
                                                js_text = self.settings['saran_dosen'].replace("'", "\\'")
                                                driver.execute_script(f"arguments[0].value = '{js_text}';", text_area)
                                                self.log("Menggunakan JavaScript untuk mengisi saran...")
                                            except:
                                                self.log(f"Gagal mengisi saran: {e}")
                                except Exception as e:
                                    self.log(f"Error mencari textarea: {str(e)}")
                            
                            # Check for the checkbox at the final page
                            try:
                                checkboxes = driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
                                for checkbox in checkboxes:
                                    try:
                                        if not checkbox.is_selected():
                                            try:
                                                checkbox.click()
                                            except:
                                                # Use JavaScript if click fails
                                                driver.execute_script("arguments[0].click();", checkbox)
                                                
                                            self.log("Mengklik checkbox pernyataan...")
                                    except Exception as e:
                                        self.log(f"Gagal mengklik checkbox: {e}")
                            except Exception as e:
                                self.log(f"Error pada checkbox: {str(e)}")
                            
                            # Check if there's a "Simpan EPBM" button (final page)
                            save_buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Simpan EPBM')]")
                            if save_buttons:
                                self.log("Halaman terakhir terdeteksi.")
                                try:
                                    self.log("Menyimpan EPBM...")
                                    # Try normal click first
                                    try:
                                        save_buttons[0].click()
                                    except:
                                        # If that fails, try JavaScript click
                                        driver.execute_script("arguments[0].click();", save_buttons[0])
                                    
                                    # Handling modals that may appear
                                    try:
                                        # Wait for any modal dialog
                                        WebDriverWait(driver, 3).until(
                                            EC.presence_of_element_located((By.CLASS_NAME, "modal-dialog"))
                                        )
                                        self.log("Dialog modal terdeteksi setelah menyimpan")
                                        
                                        # Find and click any button in the modal
                                        modal_buttons = driver.find_elements(By.CSS_SELECTOR, ".modal-footer button, .modal button.btn, .modal .close")
                                        if modal_buttons:
                                            try:
                                                modal_buttons[0].click()
                                                self.log("Mengklik tombol pada modal dialog")
                                            except:
                                                # Use JavaScript if click fails
                                                driver.execute_script("arguments[0].click();", modal_buttons[0])
                                    except:
                                        # No modal found or timeout, that's fine
                                        pass
                                    
                                    # Navigate back to main page
                                    try:
                                        # Wait for completion with shorter timeouts
                                        driver.get("https://studentportal.ipb.ac.id/Akademik/EPBM/Detail")
                                        WebDriverWait(driver, 8).until(
                                            EC.presence_of_element_located((By.CSS_SELECTOR, ".btn.card.small-box"))
                                        )
                                        self.log("EPBM berhasil disimpan!")
                                    except:
                                        # If wait fails, just wait a bit and continue
                                        driver.get("https://studentportal.ipb.ac.id/Akademik/EPBM/Detail")
                                        time.sleep(2)
                                        self.log("EPBM kemungkinan berhasil disimpan.")
                                except Exception as e:
                                    # If there's an error during save, try to recover
                                    self.log(f"Terjadi error saat simpan: {str(e)}")
                                    self.log("Mencoba kembali ke halaman utama...")
                                    
                                    # If error has stacktrace, don't show it in the log
                                    if "Stacktrace:" in str(e):
                                        self.log("Error tersebut umum terjadi dan biasanya tidak mempengaruhi hasil pengisian.")
                                    
                                    try:
                                        driver.get("https://studentportal.ipb.ac.id/Akademik/EPBM/Detail")
                                        time.sleep(2)
                                    except:
                                        self.log("Gagal kembali ke halaman utama")
                                break
                            
                            # Otherwise, click "Selanjutnya" to go to the next page
                            try:
                                next_buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Selanjutnya')]")
                                if next_buttons:
                                    self.log("Menuju halaman selanjutnya...")
                                    try:
                                        next_buttons[0].click()
                                    except:
                                        # Use JavaScript if click fails
                                        driver.execute_script("arguments[0].click();", next_buttons[0])
                                        
                                    # Shorter wait time for faster processing
                                    time.sleep(0.3)  # Reduced from 0.7 to 0.3 seconds
                                else:
                                    self.log("Tidak menemukan tombol Selanjutnya atau Simpan EPBM.")
                                    break
                            except Exception as e:
                                self.log(f"Error saat mencoba ke halaman selanjutnya: {str(e)}")
                                break
                    
                    self.log("Kembali ke halaman utama EPBM.")
                    
                except Exception as e:
                    self.log(f"Terjadi error saat mengisi EPBM: {str(e)}")
                    
                    # Don't show stacktrace in the log to keep it clean
                    if "Stacktrace:" in str(e):
                        self.log("Error tersebut umum terjadi dan biasanya tidak mempengaruhi hasil pengisian.")
                    
                    # Always try to go back to the main page
                    try:
                        driver.get("https://studentportal.ipb.ac.id/Akademik/EPBM/Detail")
                        time.sleep(2)
                    except:
                        self.log("Gagal kembali ke halaman utama, mencoba lagi...")
                        try:
                            driver.get("https://studentportal.ipb.ac.id/Akademik/EPBM/Detail")
                            time.sleep(3)
                        except:
                            self.log("Gagal kembali ke halaman utama setelah beberapa percobaan")
            
            # Set final progress to 100% when everything is done
            self.progress_signal.emit(100)
            self.log("\nProses pengisian EPBM selesai!")
            
        except Exception as e:
            self.log(f"Terjadi error pada proses keseluruhan: {e}")
        finally:
            # Close the browser
            driver.quit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AutoEPBM StudentPortal")
        self.setMinimumSize(900, 650)  # More compact size
        
        # Light mode design elements
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f7;
            }
            QWidget {
                color: #333333;
            }
            QGroupBox {
                border: 1px solid #b3d4fc;
                border-radius: 8px;
                margin-top: 10px;
                font-weight: bold;
                font-size: 12px;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
                color: #2c75b3;
            }
            QTabWidget::pane {
                border: 1px solid #b3d4fc;
                border-radius: 5px;
                top: -1px;
                background-color: #ffffff;
            }
            QTabBar::tab {
                background-color: #e6eef8;
                color: #2c75b3;
                border: 1px solid #b3d4fc;
                border-bottom-color: #b3d4fc;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                min-width: 8ex;
                padding: 6px 12px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background-color: #ffffff;
                color: #2c75b3;
                border-bottom-color: #ffffff;
                font-weight: bold;
            }
            QTabBar::tab:!selected {
                margin-top: 2px;
            }
            QLabel {
                color: #333333;
            }
            QTextEdit, QLineEdit, QSpinBox {
                background-color: #ffffff;
                color: #333333;
                border: 1px solid #b3d4fc;
            }
        """)
        
        # Available courses
        self.available_courses = []
        self.selected_courses = []
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # Create header with modern logo and title
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(5, 5, 5, 10)
        
        # Logo
        logo_label = QLabel()
        script_dir = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(script_dir, "assets", "logo.png")
        if os.path.exists(logo_path):
            logo_pixmap = QPixmap(logo_path)
            logo_label.setPixmap(logo_pixmap.scaled(40, 40, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        else:
            logo_label.setText("IPB")
            logo_label.setFont(QFont("Arial", 16, QFont.Bold))
            logo_label.setStyleSheet("color: #2c75b3; background-color: #e6eef8; padding: 5px; border-radius: 5px;")
        logo_label.setAlignment(Qt.AlignCenter)
        logo_label.setFixedSize(50, 40)
        header_layout.addWidget(logo_label)
        
        # Title with modern font
        title_label = QLabel("AutoEPBM StudentPortal")
        title_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        title_label.setFont(QFont("Segoe UI", 18, QFont.Bold))
        title_label.setStyleSheet("color: #2c75b3; margin-left: 10px;")
        header_layout.addWidget(title_label, 1)
        
        # Version info for professional touch
        version_label = QLabel("v1.0.0")
        version_label.setStyleSheet("color: #888888; font-size: 10px;")
        version_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        header_layout.addWidget(version_label)
        
        main_layout.addLayout(header_layout)
        
        # Add a subtle gradient separator instead of a plain line
        gradient_frame = QFrame()
        gradient_frame.setFixedHeight(2)
        gradient_frame.setStyleSheet("""
            QFrame {
                border: none;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                                          stop:0 #2c75b3, stop:0.5 #4a90e2, stop:1 #2c75b3);
            }
        """)
        main_layout.addWidget(gradient_frame)
        
        # Main content with splitter (two-panel layout)
        splitter = QSplitter(Qt.Horizontal)
        
        # Left panel (Login & Settings)
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(6)  # Reduce spacing
        
        # Create tabs for settings
        tabs = QTabWidget()
        
        # Login panel
        login_panel = QWidget()
        self.setup_login_panel(login_panel)
        
        # Settings tab
        self.settings_tab = QScrollArea()
        self.settings_tab.setWidgetResizable(True)
        self.setup_settings_tab()
        
        tabs.addTab(login_panel, "Login")
        tabs.addTab(self.settings_tab, "Pengaturan")
        
        left_layout.addWidget(tabs)
        
        # Status section
        status_group = QGroupBox("Status")
        status_layout = QVBoxLayout()
        
        # Course counter section
        counter_layout = QGridLayout()
        
        # Found courses
        counter_layout.addWidget(QLabel("Mata Kuliah Ditemukan:"), 0, 0)
        self.found_counter = QLabel("0")
        self.found_counter.setStyleSheet("font-size: 18px; color: #3498db; font-weight: bold;")
        counter_layout.addWidget(self.found_counter, 0, 1)
        
        # Selected courses
        counter_layout.addWidget(QLabel("Mata Kuliah Dipilih:"), 1, 0)
        self.selected_counter = QLabel("0")
        self.selected_counter.setStyleSheet("font-size: 18px; color: #2ecc71; font-weight: bold;")
        counter_layout.addWidget(self.selected_counter, 1, 1)
        
        # Completed courses
        counter_layout.addWidget(QLabel("Sudah Diisi:"), 2, 0)
        self.completed_counter = QLabel("0")
        self.completed_counter.setStyleSheet("font-size: 18px; color: #f39c12; font-weight: bold;")
        counter_layout.addWidget(self.completed_counter, 2, 1)
        
        status_layout.addLayout(counter_layout)
        
        # Course list
        self.course_list = QListWidget()
        self.course_list.setAlternatingRowColors(True)
        self.course_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #b3d4fc;
                border-radius: 5px;
                padding: 5px;
                background-color: #ffffff;
            }
            QListWidget::item {
                padding: 5px;
                margin: 2px;
                border-radius: 3px;
                color: #333333;
            }
            QListWidget::item:selected {
                background-color: #e6eef8;
                color: #2c75b3;
            }
            QListWidget::item:alternate {
                background-color: #f5f5f7;
            }
        """)
        status_layout.addWidget(self.course_list)
        
        # Progress section
        progress_layout = QVBoxLayout()
        
        self.status_label = QLabel("Belum dimulai")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("font-weight: bold;")
        progress_layout.addWidget(self.status_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p% (%v/%m)")
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #b3d4fc;
                border-radius: 8px;
                text-align: center;
                height: 22px;
                background-color: #ffffff;
                color: #333333;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                                          stop:0 #2c75b3, stop:1 #4a90e2);
                border-radius: 7px;
                margin: 0.5px;
            }
        """)
        progress_layout.addWidget(self.progress_bar)
        
        status_layout.addLayout(progress_layout)
        
        # Action buttons
        buttons_layout = QHBoxLayout()
        
        self.find_courses_button = QPushButton("Cari Mata Kuliah")
        self.find_courses_button.clicked.connect(self.find_courses)
        self.find_courses_button.setStyleSheet("""
            QPushButton {
                background-color: #2c75b3;
                border-radius: 8px;
                font-weight: bold;
                padding: 8px 12px;
                min-height: 30px;
                color: white;
                border: none;
            }
            QPushButton:hover {
                background-color: #4a90e2;
            }
            QPushButton:pressed {
                background-color: #1b5493;
            }
            QPushButton:disabled {
                background-color: #d0d8e0;
                color: #8a8a8a;
            }
        """)
        self.find_courses_button.setMinimumHeight(40)
        buttons_layout.addWidget(self.find_courses_button)
        
        self.edit_courses_button = QPushButton("Edit Pilihan")
        self.edit_courses_button.clicked.connect(self.edit_selected_courses)
        self.edit_courses_button.setEnabled(False)
        self.edit_courses_button.setStyleSheet("""
            QPushButton {
                background-color: #e67e22;
                border-radius: 8px;
                font-weight: bold;
                padding: 8px 12px;
                min-height: 30px;
                color: white;
                border: none;
            }
            QPushButton:hover {
                background-color: #d35400;
            }
            QPushButton:pressed {
                background-color: #a04000;
            }
            QPushButton:disabled {
                background-color: #d0d8e0;
                color: #8a8a8a;
            }
        """)
        self.edit_courses_button.setMinimumHeight(40)
        buttons_layout.addWidget(self.edit_courses_button)
        
        status_layout.addLayout(buttons_layout)
        
        # Start automation button
        self.start_button = QPushButton("Mulai Otomasi EPBM")
        self.start_button.setEnabled(False)
        self.start_button.clicked.connect(self.start_automation)
        self.start_button.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                border-radius: 8px;
                font-weight: bold;
                padding: 10px 12px;
                min-height: 30px;
                color: white;
                border: none;
            }
            QPushButton:hover {
                background-color: #2ecc71;
            }
            QPushButton:pressed {
                background-color: #1e8449;
            }
            QPushButton:disabled {
                background-color: #d0d8e0;
                color: #8a8a8a;
            }
        """)
        self.start_button.setMinimumHeight(50)
        status_layout.addWidget(self.start_button)
        
        status_group.setLayout(status_layout)
        left_layout.addWidget(status_group)
        
        splitter.addWidget(left_panel)
        
        # Right panel (Live Log)
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        log_group = QGroupBox("Log Real-Time")
        log_layout = QVBoxLayout()
        
        # Log text with colors and formatting
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("""
            QTextEdit {
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 11pt;
                background-color: #ffffff;
                color: #333333;
                border: 1px solid #b3d4fc;
                border-radius: 5px;
                padding: 5px;
            }
        """)
        log_layout.addWidget(self.log_text)
        
        # Log control buttons
        log_control_layout = QHBoxLayout()
        
        clear_log_button = QPushButton("Bersihkan Log")
        clear_log_button.clicked.connect(self.clear_log)
        clear_log_button.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                padding: 5px;
                border-radius: 5px;
                border: none;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        log_control_layout.addWidget(clear_log_button)
        
        self.autoscroll_checkbox = QCheckBox("Auto-scroll")
        self.autoscroll_checkbox.setChecked(True)
        self.autoscroll_checkbox.setStyleSheet("""
            QCheckBox {
                color: #333333;
                spacing: 8px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 4px;
                border: 1px solid #b3d4fc;
            }
            QCheckBox::indicator:unchecked {
                background-color: #ffffff;
            }
            QCheckBox::indicator:checked {
                background-color: #2c75b3;
            }
        """)
        self.autoscroll_enabled = True
        self.autoscroll_checkbox.toggled.connect(self.toggle_autoscroll)
        log_control_layout.addWidget(self.autoscroll_checkbox)
        
        log_control_layout.addStretch()
        
        log_layout.addLayout(log_control_layout)
        
        log_group.setLayout(log_layout)
        right_layout.addWidget(log_group)
        
        splitter.addWidget(right_panel)
        
        # Set initial sizes of splitter
        splitter.setSizes([350, 550])
        
        main_layout.addWidget(splitter)
        
        # Status bar
        self.statusBar().showMessage("Siap memulai")
        
        # Worker threads
        self.finder_worker = None
        self.automation_worker = None
        
        # Add welcome message
        self.update_log("Selamat datang di AutoEPBM StudentPortal!", "success")
        self.update_log("Aplikasi ini dapat mengotomatisasi pengisian EPBM di portal mahasiswa IPB.", "info")
        self.update_log("Untuk memulai, masukkan username dan password IPB Anda, lalu klik 'Cari Mata Kuliah'.", "info")
        
        # Add credits footer
        self.add_credits_footer(main_layout)
    
    def add_credits_footer(self, main_layout):
        # Create credits section
        credits_frame = QFrame()
        credits_frame.setFrameShape(QFrame.StyledPanel)
        credits_frame.setStyleSheet("""
            QFrame {
                background-color: #f0f4f8;
                border: 1px solid #b3d4fc;
                border-radius: 8px;
                margin-top: 5px;
            }
        """)
        
        credits_layout = QHBoxLayout(credits_frame)
        credits_layout.setContentsMargins(10, 5, 10, 5)
        
        # Developer info
        dev_label = QLabel("Developed by: Dzakwan Alifi (2025)")
        dev_label.setStyleSheet("color: #2c75b3; font-weight: bold;")
        credits_layout.addWidget(dev_label)
        
        # Social media links
        github_button = QPushButton("GitHub")
        github_button.setIcon(QIcon.fromTheme("github", QIcon("assets/github.png")))
        github_button.setStyleSheet("""
            QPushButton {
                background-color: #333333;
                color: white;
                border-radius: 4px;
                padding: 5px 10px;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #2c2c2c;
            }
        """)
        github_button.setCursor(Qt.PointingHandCursor)
        github_button.clicked.connect(lambda: self.open_url("https://github.com/dzakwanalifi/"))
        
        instagram_button = QPushButton("Instagram")
        instagram_button.setIcon(QIcon.fromTheme("instagram", QIcon("assets/instagram.png")))
        instagram_button.setStyleSheet("""
            QPushButton {
                background-color: #c13584;
                color: white;
                border-radius: 4px;
                padding: 5px 10px;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #a5306f;
            }
        """)
        instagram_button.setCursor(Qt.PointingHandCursor)
        instagram_button.clicked.connect(lambda: self.open_url("https://www.instagram.com/dzakwanalifi"))
        
        credits_layout.addStretch()
        credits_layout.addWidget(github_button)
        credits_layout.addWidget(instagram_button)
        
        main_layout.addWidget(credits_frame)
    
    def open_url(self, url):
        """Open URL in default browser"""
        import webbrowser
        webbrowser.open(url)
    
    def setup_login_panel(self, panel):
        layout = QVBoxLayout(panel)
        layout.setSpacing(8)  # More compact spacing
        
        # Login credentials group
        credentials_group = QGroupBox("Masukkan Akun IPB")
        credentials_layout = QFormLayout()
        credentials_layout.setSpacing(6)  # Reduce spacing
        credentials_layout.setContentsMargins(8, 8, 8, 8)  # Reduce margins
        
        username_layout = QHBoxLayout()
        username_icon = QLabel()
        username_icon.setPixmap(QIcon.fromTheme("user-info").pixmap(16, 16))
        username_layout.addWidget(username_icon)
        
        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("Username IPB")
        self.username_input.setStyleSheet("""
            QLineEdit {
                padding: 8px;
                border: 1px solid #b3d4fc;
                border-radius: 5px;
                background-color: #ffffff;
                color: #333333;
            }
            QLineEdit:focus {
                border: 1px solid #2c75b3;
            }
        """)
        username_layout.addWidget(self.username_input)
        
        credentials_layout.addRow("Username:", username_layout)
        
        password_layout = QHBoxLayout()
        password_icon = QLabel()
        password_icon.setPixmap(QIcon.fromTheme("dialog-password").pixmap(16, 16))
        password_layout.addWidget(password_icon)
        
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Password IPB")
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setStyleSheet("""
            QLineEdit {
                padding: 8px;
                border: 1px solid #b3d4fc;
                border-radius: 5px;
                background-color: #ffffff;
                color: #333333;
            }
            QLineEdit:focus {
                border: 1px solid #2c75b3;
            }
        """)
        password_layout.addWidget(self.password_input)
        
        credentials_layout.addRow("Password:", password_layout)
        
        credentials_group.setLayout(credentials_layout)
        layout.addWidget(credentials_group)
        
        # Options group
        options_group = QGroupBox("Opsi Eksekusi")
        options_layout = QVBoxLayout()
        options_layout.setSpacing(3)  # Reduce spacing between items
        options_layout.setContentsMargins(8, 8, 8, 8)  # Reduce margins
        
        # Checkboxes with descriptions - Only leave the headless option and make it checked by default
        headless_layout = QHBoxLayout()
        self.headless_checkbox = QCheckBox("Jalankan di latar belakang (headless mode)")
        self.headless_checkbox.setChecked(True)  # Changed to true by default
        headless_layout.addWidget(self.headless_checkbox)
        headless_info = QLabel("Browser tidak akan tampil saat proses berjalan")
        headless_info.setStyleSheet("color: #7f8c8d; font-style: italic;")
        headless_layout.addWidget(headless_info, 1)
        options_layout.addLayout(headless_layout)
        
        options_group.setLayout(options_layout)
        layout.addWidget(options_group)
        
        layout.addStretch()
    
    def setup_settings_tab(self):
        settings_widget = QWidget()
        self.settings_tab.setWidget(settings_widget)
        layout = QVBoxLayout(settings_widget)
        layout.setSpacing(8)  # More compact spacing
        
        # Mata Kuliah settings group
        matkul_group = QGroupBox("Pengaturan Pertanyaan Mata Kuliah")
        matkul_layout = QFormLayout()
        
        self.matkul_sesuai_harapan = QSpinBox()
        self.matkul_sesuai_harapan.setRange(1, 4)
        self.matkul_sesuai_harapan.setValue(4)
        self.matkul_sesuai_harapan.setStyleSheet("padding: 5px;")
        matkul_layout.addRow("Proses pembelajaran sesuai dengan yang diharapkan:", self.matkul_sesuai_harapan)
        
        self.matkul_menyenangkan = QSpinBox()
        self.matkul_menyenangkan.setRange(1, 4)
        self.matkul_menyenangkan.setValue(4)
        self.matkul_menyenangkan.setStyleSheet("padding: 5px;")
        matkul_layout.addRow("Proses pembelajaran menyenangkan dan menginspirasi:", self.matkul_menyenangkan)
        
        self.matkul_asesmen = QSpinBox()
        self.matkul_asesmen.setRange(1, 4)
        self.matkul_asesmen.setValue(4)
        self.matkul_asesmen.setStyleSheet("padding: 5px;")
        matkul_layout.addRow("Mahasiswa mengetahui proses asesmen secara terbuka:", self.matkul_asesmen)
        
        self.matkul_hardskill = QSpinBox()
        self.matkul_hardskill.setRange(1, 4)
        self.matkul_hardskill.setValue(4)
        self.matkul_hardskill.setStyleSheet("padding: 5px;")
        matkul_layout.addRow("Mahasiswa mendapatkan kesempatan meningkatkan hardskill/sofskill:", self.matkul_hardskill)
        
        self.matkul_dokumen = QSpinBox()
        self.matkul_dokumen.setRange(1, 4)
        self.matkul_dokumen.setValue(4)
        self.matkul_dokumen.setStyleSheet("padding: 5px;")
        matkul_layout.addRow("Dokumen ajar dilengkapi dan bisa diakses:", self.matkul_dokumen)
        
        matkul_group.setLayout(matkul_layout)
        layout.addWidget(matkul_group)
        
        # Dosen settings group
        dosen_group = QGroupBox("Pengaturan Pertanyaan Dosen")
        dosen_layout = QFormLayout()
        
        self.dosen_ceramah = QSpinBox()
        self.dosen_ceramah.setRange(1, 4)
        self.dosen_ceramah.setValue(4)
        self.dosen_ceramah.setStyleSheet("padding: 5px;")
        dosen_layout.addRow("Dosen memberikan kuliah dengan metode ceramah:", self.dosen_ceramah)
        
        self.dosen_mentor = QSpinBox()
        self.dosen_mentor.setRange(1, 4)
        self.dosen_mentor.setValue(4)
        self.dosen_mentor.setStyleSheet("padding: 5px;")
        dosen_layout.addRow("Dosen menyampaikan kuliah dengan menjadi mentor:", self.dosen_mentor)
        
        self.dosen_ilustrasi = QSpinBox()
        self.dosen_ilustrasi.setRange(1, 4)
        self.dosen_ilustrasi.setValue(4)
        self.dosen_ilustrasi.setStyleSheet("padding: 5px;")
        dosen_layout.addRow("Dosen memberikan contoh/ilustrasi dalam kehidupan nyata:", self.dosen_ilustrasi)
        
        self.dosen_teknologi = QSpinBox()
        self.dosen_teknologi.setRange(1, 4)
        self.dosen_teknologi.setValue(4)
        self.dosen_teknologi.setStyleSheet("padding: 5px;")
        dosen_layout.addRow("Dosen menfaatkan ketersediaan teknologi:", self.dosen_teknologi)
        
        self.dosen_feedback = QSpinBox()
        self.dosen_feedback.setRange(1, 4)
        self.dosen_feedback.setValue(4)
        self.dosen_feedback.setStyleSheet("padding: 5px;")
        dosen_layout.addRow("Dosen memberikan umpan balik:", self.dosen_feedback)
        
        self.saran_dosen = QTextEdit()
        self.saran_dosen.setPlaceholderText("Masukkan saran untuk dosen...")
        self.saran_dosen.setText("Terima kasih atas ilmu yang diberikan. Semoga pembelajaran ke depannya semakin baik.")
        self.saran_dosen.setStyleSheet("padding: 5px;")
        dosen_layout.addRow("Saran untuk dosen:", self.saran_dosen)
        
        dosen_group.setLayout(dosen_layout)
        layout.addWidget(dosen_group)
        
        # Sarana Prasarana settings group
        sarpras_group = QGroupBox("Pengaturan Sarana Prasarana")
        sarpras_layout = QFormLayout()
        
        self.sarpras_kenyamanan = QSpinBox()
        self.sarpras_kenyamanan.setRange(1, 4)
        self.sarpras_kenyamanan.setValue(4)
        self.sarpras_kenyamanan.setStyleSheet("padding: 5px;")
        sarpras_layout.addRow("Kenyamanan kelas dan flexibility learning:", self.sarpras_kenyamanan)
        
        self.sarpras_internet = QSpinBox()
        self.sarpras_internet.setRange(1, 4)
        self.sarpras_internet.setValue(4)
        self.sarpras_internet.setStyleSheet("padding: 5px;")
        sarpras_layout.addRow("Fasilitas internet memadai:", self.sarpras_internet)
        
        self.sarpras_toilet = QSpinBox()
        self.sarpras_toilet.setRange(1, 4)
        self.sarpras_toilet.setValue(4)
        self.sarpras_toilet.setStyleSheet("padding: 5px;")
        sarpras_layout.addRow("Kualitas toilet yang sehat:", self.sarpras_toilet)
        
        sarpras_group.setLayout(sarpras_layout)
        layout.addWidget(sarpras_group)
        
        # Add preset buttons
        preset_group = QGroupBox("Preset Penilaian")
        preset_layout = QHBoxLayout()
        
        all_max_button = QPushButton("Semua Nilai 4")
        all_max_button.clicked.connect(self.set_all_max)
        all_max_button.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                color: white;
                padding: 5px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
        """)
        preset_layout.addWidget(all_max_button)
        
        all_mid_button = QPushButton("Semua Nilai 3")
        all_mid_button.clicked.connect(self.set_all_mid)
        all_mid_button.setStyleSheet("""
            QPushButton {
                background-color: #f39c12;
                color: white;
                padding: 5px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #d35400;
            }
        """)
        preset_layout.addWidget(all_mid_button)
        
        preset_group.setLayout(preset_layout)
        layout.addWidget(preset_group)
        
        # Light mode spinboxes
        spinbox_style = """
            QSpinBox {
                border: 1px solid #b3d4fc;
                border-radius: 5px;
                padding: 4px;
                background-color: #ffffff;
                color: #333333;
                selection-background-color: #e6eef8;
            }
            QSpinBox::up-button, QSpinBox::down-button {
                width: 20px;
                border-radius: 3px;
                background-color: #e6eef8;
                color: #2c75b3;
            }
            QSpinBox::up-button:hover, QSpinBox::down-button:hover {
                background-color: #b3d4fc;
            }
        """
        
        # Apply style to all spinboxes
        for spinbox in [
            self.matkul_sesuai_harapan, self.matkul_menyenangkan, self.matkul_asesmen,
            self.matkul_hardskill, self.matkul_dokumen, self.dosen_ceramah,
            self.dosen_mentor, self.dosen_ilustrasi, self.dosen_teknologi,
            self.dosen_feedback, self.sarpras_kenyamanan, self.sarpras_internet,
            self.sarpras_toilet
        ]:
            spinbox.setStyleSheet(spinbox_style)
        
        # Light mode textarea for suggestions
        self.saran_dosen.setStyleSheet("""
            QTextEdit {
                border: 1px solid #b3d4fc;
                border-radius: 8px;
                padding: 8px;
                background-color: #ffffff;
                color: #333333;
                selection-background-color: #e6eef8;
            }
            QTextEdit:focus {
                border: 1px solid #2c75b3;
            }
        """)
    
    def toggle_autoscroll(self, enabled):
        self.autoscroll_enabled = enabled
    
    def clear_log(self):
        self.log_text.clear()
        # Update style for log text to match light theme
        self.log_text.setStyleSheet("""
            QTextEdit {
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 11pt;
                background-color: #ffffff;
                color: #333333;
                border: 1px solid #b3d4fc;
                border-radius: 5px;
                padding: 5px;
            }
        """)
        self.update_log("Log dibersihkan.", "info")
    
    def set_all_max(self):
        # Set all rating spinboxes to 4
        for spinbox in [
            self.matkul_sesuai_harapan, self.matkul_menyenangkan, self.matkul_asesmen,
            self.matkul_hardskill, self.matkul_dokumen, self.dosen_ceramah,
            self.dosen_mentor, self.dosen_ilustrasi, self.dosen_teknologi,
            self.dosen_feedback, self.sarpras_kenyamanan, self.sarpras_internet,
            self.sarpras_toilet
        ]:
            spinbox.setValue(4)
        self.update_log("Semua nilai diatur ke 4 (maksimum).", "info")
            
    def set_all_mid(self):
        # Set all rating spinboxes to 3
        for spinbox in [
            self.matkul_sesuai_harapan, self.matkul_menyenangkan, self.matkul_asesmen,
            self.matkul_hardskill, self.matkul_dokumen, self.dosen_ceramah,
            self.dosen_mentor, self.dosen_ilustrasi, self.dosen_teknologi,
            self.dosen_feedback, self.sarpras_kenyamanan, self.sarpras_internet,
            self.sarpras_toilet
        ]:
            spinbox.setValue(3)
        self.update_log("Semua nilai diatur ke 3 (sedang).", "info")
    
    @pyqtSlot(str)
    def update_log(self, message, level="normal"):
        # Format based on message level - update colors for light theme
        cursor = self.log_text.textCursor()
        format = QTextCharFormat()
        
        if level == "success":
            format.setForeground(QColor("#27ae60"))  # Darker green for light background
            format.setFontWeight(QFont.Bold)
        elif level == "info":
            format.setForeground(QColor("#2980b9"))  # Darker blue for light background
        elif level == "warning":
            format.setForeground(QColor("#d35400"))  # Darker orange for light background
        elif level == "error":
            format.setForeground(QColor("#c0392b"))  # Darker red for light background
            format.setFontWeight(QFont.Bold)
        else:
            format.setForeground(QColor("#333333"))  # Default text color for light background
        
        cursor.movePosition(QTextCursor.End)
        cursor.insertText(time.strftime("[%H:%M:%S] "))
        
        cursor.insertText(message + "\n", format)
        
        # Auto-scroll if enabled
        if self.autoscroll_enabled:
            scrollbar = self.log_text.verticalScrollBar()
            scrollbar.setValue(scrollbar.maximum())

    # Modified worker callback methods
    def find_courses(self):
        username = self.username_input.text()
        password = self.password_input.text()
        
        if not username or not password:
            QMessageBox.warning(self, "Input Error", "Mohon masukkan username dan password!")
            return
            
        # Disable buttons
        self.find_courses_button.setEnabled(False)
        self.find_courses_button.setText("Mencari Mata Kuliah...")
        self.status_label.setText("Mencari mata kuliah...")
        
        # Prepare credentials
        credentials = {
            'username': username,
            'password': password
        }
        
        # Clear log
        self.clear_log()
        self.update_log("Memulai pencarian mata kuliah...", "info")
        
        # Create and start finder worker thread
        self.finder_worker = CourseFinderWorker(credentials)
        self.finder_worker.update_signal.connect(lambda msg: self.update_log(msg))
        self.finder_worker.courses_found_signal.connect(self.show_course_selection)
        self.finder_worker.finished_signal.connect(self.finder_finished)
        self.finder_worker.start()
    
    def show_course_selection(self, courses):
        self.available_courses = courses
        
        # Update counters
        total_courses = len(courses)
        completed_courses = sum(1 for c in courses if c.get('is_completed', False))
        
        self.found_counter.setText(str(total_courses))
        self.completed_counter.setText(str(completed_courses))
        
        # Clear course list
        self.course_list.clear()
        
        # Add courses to list with status indicators
        for course in courses:
            title = course['title']
            desc = course['desc']
            status = "[Sudah Diisi]" if course.get('is_completed', False) else ""
            
            item_text = f"{title}: {desc} {status}"
            item = QListWidgetItem(item_text)
            
            if course.get('is_completed', False):
                item.setForeground(QColor("#f39c12"))  # Orange for completed
            else:
                item.setForeground(QColor("#2ecc71"))  # Green for available
                
            self.course_list.addItem(item)
        
        # Create course selection dialog
        dialog = CourseSelectionDialog(courses, self)
        if dialog.exec_():
            self.selected_courses = dialog.get_selected_courses()
            self.update_selected_courses_counter()
            
    def update_selected_courses_counter(self):
        if self.selected_courses:
            self.selected_counter.setText(str(len(self.selected_courses)))
            self.start_button.setEnabled(True)
            self.edit_courses_button.setEnabled(True)
            self.status_label.setText(f"Siap mengisi {len(self.selected_courses)} mata kuliah")
        else:
            self.selected_counter.setText("0")
            self.start_button.setEnabled(False)
            self.edit_courses_button.setEnabled(False)
            self.status_label.setText("Belum ada mata kuliah yang dipilih")
    
    def finder_finished(self, success, message):
        # Re-enable finder button
        self.find_courses_button.setEnabled(True)
        self.find_courses_button.setText("Cari Mata Kuliah")
        
        # Show message if error
        if not success:
            self.update_log(message, "error")
            self.status_label.setText("Pencarian gagal")
            QMessageBox.critical(self, "Error", message)
        else:
            self.update_log(message, "success")
            self.status_label.setText("Pencarian selesai")
    
    def start_automation(self):
        username = self.username_input.text()
        password = self.password_input.text()
        
        if not username or not password:
            QMessageBox.warning(self, "Input Error", "Mohon masukkan username dan password!")
            return
            
        if not self.selected_courses:
            QMessageBox.warning(self, "Input Error", "Tidak ada mata kuliah yang dipilih!")
            return
        
        # Disable start button
        self.start_button.setEnabled(False)
        self.start_button.setText("Sedang Berjalan...")
        
        # Prepare credentials and settings
        credentials = {
            'username': username,
            'password': password
        }
        
        settings = {
            'headless': self.headless_checkbox.isChecked(),
            # Remove test_mode setting, make it always save
            'matkul_sesuai_harapan': self.matkul_sesuai_harapan.value(),
            'matkul_menyenangkan': self.matkul_menyenangkan.value(),
            'matkul_asesmen': self.matkul_asesmen.value(),
            'matkul_hardskill': self.matkul_hardskill.value(),
            'matkul_dokumen': self.matkul_dokumen.value(),
            'dosen_ceramah': self.dosen_ceramah.value(),
            'dosen_mentor': self.dosen_mentor.value(),
            'dosen_ilustrasi': self.dosen_ilustrasi.value(),
            'dosen_teknologi': self.dosen_teknologi.value(),
            'dosen_feedback': self.dosen_feedback.value(),
            'saran_dosen': self.saran_dosen.toPlainText(),
            'sarpras_kenyamanan': self.sarpras_kenyamanan.value(),
            'sarpras_internet': self.sarpras_internet.value(),
            'sarpras_toilet': self.sarpras_toilet.value()
        }
        
        # Switch to log tab
        self.centralWidget().findChild(QTabWidget).setCurrentIndex(2)
        
        # Clear log
        self.clear_log()
        
        # Create and start worker thread
        self.automation_worker = EPBMAutomationWorker(credentials, settings, self.selected_courses)
        self.automation_worker.update_signal.connect(self.update_log)
        self.automation_worker.progress_signal.connect(self.update_progress)
        self.automation_worker.finished_signal.connect(self.automation_finished)
        self.automation_worker.start()
        
    @pyqtSlot(int)
    def update_progress(self, value):
        # Update the progress bar with a smooth animation
        self.progress_bar.setValue(value)
        
        # Also update the status label based on progress
        if value < 10:
            self.status_label.setText("Memulai proses...")
        elif value < 30:
            self.status_label.setText("Mengisi data mata kuliah...")
        elif value < 60:
            self.status_label.setText("Mengisi data dosen...")
        elif value < 90:
            self.status_label.setText("Menyelesaikan pengisian...")
        elif value < 100:
            self.status_label.setText("Hampir selesai...")
        else:
            self.status_label.setText("Pengisian selesai!")
            
        # Update the status bar as well
        self.statusBar().showMessage(f"Pengisian EPBM: {value}% selesai")
        
    @pyqtSlot(bool, str)
    def automation_finished(self, success, message):
        # Re-enable start button
        self.start_button.setEnabled(True)
        self.start_button.setText("Mulai Otomasi")
        
        # Show message
        if success:
            QMessageBox.information(self, "Selesai", message)
        else:
            QMessageBox.critical(self, "Error", message)
            
    def edit_selected_courses(self):
        """Open course selection dialog with current courses to allow editing."""
        if not self.available_courses:
            QMessageBox.warning(self, "Perhatian", "Daftar mata kuliah tidak tersedia. Silakan cari mata kuliah terlebih dahulu.")
            return
            
        # Create course selection dialog with pre-selected courses
        dialog = CourseSelectionDialog(self.available_courses, self)
        
        # Pre-select the currently selected courses
        selected_ids = [course['index'] for course in self.selected_courses]
        for checkbox in dialog.checkboxes:
            if 'course_data' in dir(checkbox) and checkbox.course_data.get('index') in selected_ids:
                checkbox.setChecked(True)
            else:
                checkbox.setChecked(False)
                
        if dialog.exec_():
            self.selected_courses = dialog.get_selected_courses()
            self.update_selected_courses_counter()
            self.update_log(f"Pilihan mata kuliah diperbarui: {len(self.selected_courses)} mata kuliah terpilih.", "info")

# Main application entry point
def main():
    app = QApplication(sys.argv)
    
    # Create assets directory if it doesn't exist
    script_dir = os.path.dirname(os.path.abspath(__file__))
    assets_dir = os.path.join(script_dir, "assets")
    if not os.path.exists(assets_dir):
        os.makedirs(assets_dir)
    
    # Set application style more compact
    app.setStyle("Fusion")
    
    # Consistent message box and checkbox styling
    app.setStyleSheet("""
        QMessageBox {
            background-color: #ffffff;
        }
        QMessageBox QLabel {
            color: #333333;
        }
        QMessageBox QPushButton {
            border-radius: 5px;
            font-weight: bold;
            padding: 8px 12px;
            background-color: #2c75b3;
            color: white;
            border: none;
            min-width: 80px;
        }
        QMessageBox QPushButton:hover {
            background-color: #4a90e2;
        }
        QMessageBox QPushButton:disabled {
            background-color: #d0d8e0;
            color: #8a8a8a;
        }
        QCheckBox {
            color: #333333;
            spacing: 8px;
        }
        QCheckBox::indicator {
            width: 18px;
            height: 18px;
            border-radius: 4px;
            border: 1px solid #b3d4fc;
        }
        QCheckBox::indicator:unchecked {
            background-color: #ffffff;
        }
        QCheckBox::indicator:checked {
            background-color: #2c75b3;
        }
        QCheckBox::indicator:disabled {
            background-color: #f0f0f0;
            border: 1px solid #d0d0d0;
        }
    """)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()