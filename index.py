import sys
import logging
import os
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QLineEdit, QHBoxLayout, QFileDialog, QSpacerItem, QSizePolicy, QGroupBox, QProgressBar
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from pdf2image import convert_from_path
import pytesseract
from PyPDF2 import PdfWriter, PdfReader
from io import BytesIO

# Configuração de logging
log_path = os.path.join(os.path.expanduser('~'), 'app.log')
logging.basicConfig(level=logging.DEBUG, filename=log_path, filemode='w',
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# Configure o caminho do Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class ConversionThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)

    def __init__(self, input_pdf, output_file, operation):
        super().__init__()
        self.input_pdf = input_pdf
        self.output_file = output_file
        self.operation = operation

    def run(self):
        try:
            logging.info(f"Iniciando conversão: {self.operation}")
            if self.operation == "text":
                self.convert_pdf_to_text(self.input_pdf, self.output_file)
            elif self.operation == "searchable_pdf":
                self.convert_pdf_to_searchable(self.input_pdf, self.output_file)
            self.finished.emit(self.output_file)
        except Exception as e:
            logging.error(f"Erro durante a conversão: {e}")

    def convert_pdf_to_searchable(self, input_pdf, output_pdf):
        try:
            logging.info(f"Convertendo PDF para PDF selecionável: {input_pdf}")
            images = convert_from_path(input_pdf, dpi=300)
            pdf_writer = PdfWriter()

            total_pages = len(images)
            for i, image in enumerate(images):
                gray_image = image.convert('L')
                enhanced_image = gray_image.point(lambda x: 0 if x < 140 else 255, '1')
                custom_config = r'--oem 3 --psm 6'
                pdf_bytes = pytesseract.image_to_pdf_or_hocr(enhanced_image, extension='pdf', config=custom_config)
                pdf_page = PdfReader(BytesIO(pdf_bytes)).pages[0]
                pdf_writer.add_page(pdf_page)

                # Emit progress
                progress = int((i + 1) / total_pages * 100)
                self.progress.emit(progress)
                self.msleep(50)  # Pequena pausa para permitir a atualização da GUI

            with open(output_pdf, 'wb') as f:
                pdf_writer.write(f)
            logging.info(f"Conversão para PDF selecionável concluída: {output_pdf}")
        except Exception as e:
            logging.error(f"Erro durante a conversão para PDF selecionável: {e}")

    def convert_pdf_to_text(self, input_pdf, output_txt):
        try:
            logging.info(f"Convertendo PDF para texto: {input_pdf}")
            images = convert_from_path(input_pdf, dpi=300)

            total_pages = len(images)
            with open(output_txt, 'w', encoding='utf-8') as txt_file:
                for i, image in enumerate(images):
                    text = pytesseract.image_to_string(image, lang='por')
                    txt_file.write(text + '\n')

                    # Emit progress
                    progress = int((i + 1) / total_pages * 100)
                    self.progress.emit(progress)
                    self.msleep(50)  # Pequena pausa para permitir a atualização da GUI
            logging.info(f"Conversão para texto concluída: {output_txt}")
        except Exception as e:
            logging.error(f"Erro durante a conversão para texto: {e}")

class PDFConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('OFFICE OCR')
        self.setGeometry(100, 100, 800, 600)
        self.setWindowFlags(Qt.Window | Qt.WindowCloseButtonHint | Qt.WindowMinimizeButtonHint)
        self.setContentsMargins(10, 10, 10, 0)

        # Defina o ícone da janela
        self.setWindowIcon(QIcon('./office.ico'))

        # Layout principal
        layout = QVBoxLayout()
        
        # Espaçador
        layout.addItem(QSpacerItem(10, 10, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Título
        title = QLabel('Office OCR')
        title.setFont(QFont('Italic', 45, QFont.Bold))
        title.setStyleSheet('color: #FF820E; padding: 10px; font-size: 95px;')
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # Espaçador
        layout.addItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))
        
        # Grupo para a barra de pesquisa e botão
        file_group = QGroupBox("Selecionar Arquivo PDF")
        file_layout = QHBoxLayout()

        # Barra de pesquisa
        self.search_bar = QLineEdit(self)
        self.search_bar.setPlaceholderText('Selecione um arquivo PDF...')
        file_layout.addWidget(self.search_bar)
        self.search_bar.setStyleSheet("font-size: 18px;")
        
        # Botão para abrir o diálogo de arquivo
        file_button = QPushButton('📂', self)
        file_button.setFixedSize(42, 42)
        file_button.setStyleSheet("font-size: 18px;")
        file_button.clicked.connect(self.open_file_dialog)
        file_layout.addWidget(file_button)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # Botões para selecionar a operação
        self.text_button = QPushButton('Texto', self)
        self.text_button.setCheckable(True)
        self.text_button.setStyleSheet('font-size: 20px; padding: 10px;')
        self.text_button.clicked.connect(self.select_text_conversion)

        self.pdf_button = QPushButton('PDF Selecionável', self)
        self.pdf_button.setCheckable(True)
        self.pdf_button.setStyleSheet('font-size: 20px; padding: 10px;')
        self.pdf_button.clicked.connect(self.select_pdf_conversion)

        # Layout para os botões de seleção
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.text_button)
        button_layout.addWidget(self.pdf_button)
        layout.addLayout(button_layout)

        # Barra de progresso
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Botão de conversão
        self.convert_button = QPushButton('CONVERTER')
        self.convert_button.setFont(QFont('Calibri', 16, QFont.Bold))
        self.convert_button.setStyleSheet('background-color: #FF820E; color: white; margin-top: 20px;')
        self.convert_button.setFixedSize(400, 100)
        self.convert_button.clicked.connect(self.perform_conversion)
        layout.addWidget(self.convert_button, alignment=Qt.AlignCenter)

        # Espaçador
        layout.addItem(QSpacerItem(20, 100, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Crédito com link apenas no nome
        credit = QLabel('Software desenvolvido por <a href="https://github.com/Alexandre3000z" style="color: blue; text-decoration: underline;">João Alexandre</a> para Assessoria Contábil Office.')
        credit.setOpenExternalLinks(True)
        credit.setFont(QFont('Calibri', 9, QFont.Bold))
        credit.setAlignment(Qt.AlignLeft | Qt.AlignBottom)
        layout.addWidget(credit)

        # Configura o layout
        self.setLayout(layout)
    
    def open_file_dialog(self):
        try:
            options = QFileDialog.Options()
            file_path, _ = QFileDialog.getOpenFileName(self, "Selecione um Documento PDF", "", "PDF Files (*.pdf);;Todos os Arquivos (*)", options=options)
            if file_path:
                self.search_bar.setText(file_path)
                logging.info(f"Arquivo selecionado: {file_path}")
        except Exception as e:
            logging.error(f"Erro ao selecionar arquivo: {e}")

    def select_text_conversion(self):
        self.text_button.setChecked(True)
        self.pdf_button.setChecked(False)

    def select_pdf_conversion(self):
        self.pdf_button.setChecked(True)
        self.text_button.setChecked(False)

    def perform_conversion(self):
        try:
            input_pdf = self.search_bar.text()
            if input_pdf:
                logging.info(f"Iniciando conversão para: {input_pdf}")
                # Desabilitar widgets durante a conversão
                self.set_widgets_enabled(False)
                self.progress_bar.setVisible(True)
                self.progress_bar.setValue(0)
                if self.text_button.isChecked():
                    output_txt, _ = QFileDialog.getSaveFileName(self, "Salvar Texto Como", "", "Text Files (*.txt);;Todos os Arquivos (*)")
                    if output_txt:
                        logging.info(f"Arquivo de saída para texto: {output_txt}")
                        self.thread = ConversionThread(input_pdf, output_txt, "text")
                        self.thread.progress.connect(self.progress_bar.setValue)
                        self.thread.finished.connect(self.on_conversion_finished)
                        self.thread.start()
                elif self.pdf_button.isChecked():
                    output_pdf, _ = QFileDialog.getSaveFileName(self, "Salvar PDF Como", "", "PDF Files (*.pdf);;Todos os Arquivos (*)")
                    if output_pdf:
                        logging.info(f"Arquivo de saída para PDF: {output_pdf}")
                        self.thread = ConversionThread(input_pdf, output_pdf, "searchable_pdf")
                        self.thread.progress.connect(self.progress_bar.setValue)
                        self.thread.finished.connect(self.on_conversion_finished)
                        self.thread.start()
        except Exception as e:
            logging.error(f"Erro durante a conversão: {e}")

    def set_widgets_enabled(self, enabled):
        self.search_bar.setEnabled(enabled)
        self.text_button.setEnabled(enabled)
        self.pdf_button.setEnabled(enabled)
        self.convert_button.setEnabled(enabled)

    def on_conversion_finished(self, output_file):
        self.progress_bar.setVisible(False)
        self.set_widgets_enabled(True)
        logging.info(f"Conversão concluída: {output_file}")

def main():
    app = QApplication(sys.argv)
    ex = PDFConverterApp()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()