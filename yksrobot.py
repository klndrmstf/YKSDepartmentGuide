import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView
from PyQt5.QtGui import QColor, QFont
from PyQt5.QtCore import Qt

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Kullanıcı Sıralama Girişi")
        self.setGeometry(100, 100, 300, 200)

        # Central widget
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        # Layout
        self.layout = QVBoxLayout(self.central_widget)

        # Label
        self.label = QLabel("Kullanıcı Sıralaması (4000-40000):", self)
        self.layout.addWidget(self.label)

        # Line edit
        self.line_edit = QLineEdit(self)
        self.layout.addWidget(self.line_edit)

        # Button
        self.button = QPushButton("Hesapla", self)
        self.button.clicked.connect(self.calculate_and_show_results)
        self.layout.addWidget(self.button)

    def calculate_and_show_results(self):
        try: 
            ks_value = int(self.line_edit.text())
            

            # Dosya yolunu belirtin
            file_path = 'sondeneme.xlsx'

            # Excel dosyasını yükleyelim
            df = pd.read_excel(file_path)

            # Sütun adlarını belirleyelim
            df.columns = ['Üniversite', 'Bölüm', 'KONTENJAN 2024', 'KO 2023', 'SIRALAMA 2023', 'SIRALAMA 2022', 'SIRALAMA 2021']

            # Sütunları sayısal verilere çevirelim
            df['SIRALAMA 2023'] = df['SIRALAMA 2023'].astype(float)
            df['SIRALAMA 2022'] = df['SIRALAMA 2022'].astype(float)
            df['SIRALAMA 2021'] = df['SIRALAMA 2021'].astype(float)

            # ST24 hesaplaması
            df['ST24'] = df.apply(lambda row: max(row['SIRALAMA 2023'] + (row['SIRALAMA 2023'] - row['SIRALAMA 2022'] + row['SIRALAMA 2022'] - row['SIRALAMA 2021']) / 2, row['SIRALAMA 2023'] * 0.5), axis=1)

            # Kontenjan Değişim Yüzdesi Hesaplaması
            df['KO 2023'] = df['KO 2023'].astype(str).str.split('+').str[0].astype(int)
            df['KONTENJAN 2024'] = df['KONTENJAN 2024'].astype(str).str.split('+').str[0].astype(int)
            df['Kontenjan_Değişim_Yüzdesi'] = ((df['KONTENJAN 2024'] - df['KO 2023']) / df['KO 2023']) * 100

            # Kontenjan ve Sıralama Değişimine Göre Tahmini 2024 Sıralaması (TAHMİNİ SIRALAMA 2024)
            df['TAHMİNİ SIRALAMA 2024'] = df['ST24'] * (1 + df['Kontenjan_Değişim_Yüzdesi'] / 100)

            # Kullanıcı sıralaması
            KS = ks_value

            # Hesaplamalar
            df['Fark'] = abs(df['TAHMİNİ SIRALAMA 2024'] - KS)
            df["Ondeyiz"] = df['TAHMİNİ SIRALAMA 2024'] > KS
            df['SIR 2023_30_Yüzde'] = df['SIRALAMA 2023'] * 0.3
            df["DELTA30"] = df['Fark'] > df['SIR 2023_30_Yüzde']
            df["2024>2023"] = df["TAHMİNİ SIRALAMA 2024"] > df["SIRALAMA 2023"]

            # Not atama fonksiyonu
            def not_atama(row):
                if row["Ondeyiz"] and row["DELTA30"] and row['2024>2023']:
                    return 'A'
                elif (row["Ondeyiz"] and row["DELTA30"] and not row['2024>2023']
                    or row["Ondeyiz"] and not row["DELTA30"] and row['2024>2023']): 
                    return 'B'
                elif (row["Ondeyiz"] and not row["DELTA30"] and not row['2024>2023']
                    or not row["Ondeyiz"] and not row["DELTA30"] and row['2024>2023']):
                    return 'C'
                elif (not row["Ondeyiz"] and not row["DELTA30"] and not row['2024>2023']
                    or not row["Ondeyiz"] and row["DELTA30"] and row['2024>2023']):
                    return 'D'
                else:
                    return "E"

            df['Not'] = df.apply(not_atama, axis=1)

            df = df.drop(columns=['Fark', 'KO 2023',"SIR 2023_30_Yüzde","Kontenjan_Değişim_Yüzdesi","ST24","Ondeyiz",'2024>2023',"DELTA30"])
            # Renk kodları
            colors = {'A': QColor('darkgreen'), 'B': QColor('lightgreen'), 'C': QColor('yellow'), 'D': QColor('orange'),'E': QColor('red')}

            # Sonuçları görselleştiren pencereyi aç
            self.results_window = ResultsWindow(df, colors)
            self.results_window.show()
        
        except Exception:
            self.label.setText("Lütfen geçerli bir değer girin.")


class ResultsWindow(QMainWindow):
    def __init__(self, df, colors):
        super().__init__()
        self.setWindowTitle("Sonuçlar")
        self.setGeometry(150, 150, 1200, 800)

        # Central widget
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        # Layout
        self.layout = QVBoxLayout(self.central_widget)

        # QTableWidget
        self.table_widget = QTableWidget()
        self.table_widget.setRowCount(len(df))
        self.table_widget.setColumnCount(len(df.columns))

        # Sütun başlıklarını ayarlayalım
        self.table_widget.setHorizontalHeaderLabels(df.columns)

        # Verileri tabloya ekleyelim
        for i in range(len(df)):
            for j in range(len(df.columns)):
                value = df.iat[i, j]
                item = QTableWidgetItem()
                if pd.isna(value):
                    item.setText("")
                else:
                    item.setText(str(int(value)) if isinstance(value, (int, float)) and not pd.isna(value) else str(value))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Düzenlemeyi devre dışı bırak
                self.table_widget.setItem(i, j, item)

        # TAHMİNİ SIRALAMA 2024 sütununu özel olarak vurgulayalım
        for i in range(len(df)):
            self.table_widget.item(i, df.columns.get_loc('TAHMİNİ SIRALAMA 2024')).setBackground(QColor('lightblue'))

        # Renk kodlarını tabloya ekleyelim
        for i in range(len(df)):
            for j in range(len(df.columns)):
                self.table_widget.item(i, j).setBackground(colors[df['Not'].iloc[i]])

        # Sütun genişliklerini içeriklere göre ayarlayalım
        column_widths = [20, 20, 10, 10, 10, 10, 10]  # Burada genişlikleri ayarlayın
        for col, width in enumerate(column_widths):
            self.table_widget.setColumnWidth(col, width)

        # Sütun genişliklerini pencere boyutuna göre dinamik olarak ayarlayalım
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # Yazıları ortalama ve büyütme
        
        font = QFont("Calibri", 20)  # Yazı tipi ve boyutunu belirleyin
        uni_font = QFont("Calibri", 10)  # Yazı tipi ve boyutunu belirleyin
        dept_font = QFont("Calibri", 10)  # Yazı tipi ve boyutunu belirleyin


        uni_col = df.columns.get_loc('Üniversite')
        dept_col = df.columns.get_loc('Bölüm')
        numeratic_col = [2,3,4,5,6,7]



        for row in range(len(df)):
            self.table_widget.item(row, uni_col).setFont(uni_font)
            self.table_widget.item(row, dept_col).setFont(dept_font)

        # Diğer sütunların fontunu küçültelim
        for col in range(len(df.columns)):
            if col in numeratic_col:
                for row in range(len(df)):
                    self.table_widget.item(row, col).setFont(font)
                    self.table_widget.item(row, col).setTextAlignment(Qt.AlignCenter)
        
        
        # Tabloda kaydırma çubuklarını etkinleştirelim
        self.table_widget.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.table_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)

        # Table widget'ı layout'a ekleyelim
        self.layout.addWidget(self.table_widget)

# Uygulama başlat
app = QApplication(sys.argv)
main_window = MainWindow()
main_window.show()
sys.exit(app.exec_())
