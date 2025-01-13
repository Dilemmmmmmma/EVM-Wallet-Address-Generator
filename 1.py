from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                           QMessageBox, QFileDialog, QGridLayout,
                           QScrollArea, QFrame, QRadioButton, QButtonGroup)
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QClipboard
from web3 import Web3
from mnemonic import Mnemonic
import secrets
import pandas as pd
import sys

# 支持的链和它们的配置
SUPPORTED_CHAINS = {
    'ETH': {
        'chain_id': 1,
        'path': "m/44'/60'/0'/0/0"
    },
    'BSC': {
        'chain_id': 56,
        'path': "m/44'/60'/0'/0/0"
    },
    'HECO': {
        'chain_id': 128,
        'path': "m/44'/60'/0'/0/0"
    },
    'MATIC': {
        'chain_id': 137,
        'path': "m/44'/60'/0'/0/0"
    },
    'FANTOM': {
        'chain_id': 250,
        'path': "m/44'/60'/0'/0/0"
    }
}

MNEMONIC_STRENGTHS = {
    "12位": 128,
    "15位": 160,
    "18位": 192,
    "21位": 224,
    "24位": 256
}

class WalletGenerator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("EVM钱包生成器")
        self.setFixedSize(1000, 600)
        self.wallet_data = []
        self.w3 = Web3()
        self.init_ui()

    def init_ui(self):
        # 创建中心部件和主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setAlignment(Qt.AlignTop)
        
        # 标题
        title = QLabel("EVM钱包生成器")
        title.setStyleSheet("font-size: 24px; margin: 20px;")
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)
        
        # 创建表单布局
        form_layout = QGridLayout()
        form_layout.setSpacing(20)
        
        # 链选择（单选按钮组）
        chain_label = QLabel("选择链:")
        chain_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        form_layout.addWidget(chain_label, 0, 0)
        
        chain_widget = QWidget()
        chain_layout = QHBoxLayout(chain_widget)
        chain_layout.setSpacing(15)
        
        self.chain_group = QButtonGroup(self)
        for i, chain in enumerate(SUPPORTED_CHAINS.keys()):
            radio = QRadioButton(chain)
            radio.setStyleSheet("""
                QRadioButton {
                    font-size: 14px;
                    spacing: 8px;
                }
                QRadioButton::indicator {
                    width: 15px;
                    height: 15px;
                }
            """)
            if i == 0:  # 默认选中第一个
                radio.setChecked(True)
            chain_layout.addWidget(radio)
            self.chain_group.addButton(radio)
        
        form_layout.addWidget(chain_widget, 0, 1)
        
        # 助记词长度选择（单选按钮组）
        mnemonic_label = QLabel("助记词长度:")
        mnemonic_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        form_layout.addWidget(mnemonic_label, 1, 0)
        
        mnemonic_widget = QWidget()
        mnemonic_layout = QHBoxLayout(mnemonic_widget)
        mnemonic_layout.setSpacing(15)
        
        self.mnemonic_group = QButtonGroup(self)
        for i, length in enumerate(MNEMONIC_STRENGTHS.keys()):
            radio = QRadioButton(length)
            radio.setStyleSheet("""
                QRadioButton {
                    font-size: 14px;
                    spacing: 8px;
                }
                QRadioButton::indicator {
                    width: 15px;
                    height: 15px;
                }
            """)
            if i == 0:  # 默认选中第一个
                radio.setChecked(True)
            mnemonic_layout.addWidget(radio)
            self.mnemonic_group.addButton(radio)
        
        form_layout.addWidget(mnemonic_widget, 1, 1)
        
        # 数量输入
        amount_label = QLabel("生成数量:")
        amount_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        self.amount_input = QLineEdit()
        self.amount_input.setFixedWidth(200)
        form_layout.addWidget(amount_label, 2, 0)
        form_layout.addWidget(self.amount_input, 2, 1)
        
        main_layout.addLayout(form_layout)
        
        # 按钮布局
        button_layout = QHBoxLayout()
        
        # 生成按钮
        self.generate_btn = QPushButton("生成钱包")
        self.generate_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 8px 16px;
                font-size: 14px;
                border: none;
                border-radius: 4px;
                min-width: 120px;
                min-height: 40px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        self.generate_btn.clicked.connect(self.generate_wallets)
        button_layout.addWidget(self.generate_btn)
        
        # 重新生成按钮
        self.regenerate_btn = QPushButton("重新生成")
        self.regenerate_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                padding: 8px 16px;
                font-size: 14px;
                border: none;
                border-radius: 4px;
                min-width: 120px;
                min-height: 40px;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
        """)
        self.regenerate_btn.clicked.connect(self.generate_wallets)
        button_layout.addWidget(self.regenerate_btn)
        
        # 下载按钮
        self.download_btn = QPushButton("下载表格")
        self.download_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                padding: 8px 16px;
                font-size: 14px;
                border: none;
                border-radius: 4px;
                min-width: 120px;
                min-height: 40px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)
        self.download_btn.clicked.connect(self.save_to_excel)
        button_layout.addWidget(self.download_btn)
        
        main_layout.addLayout(button_layout)

        # 在按钮布局之后添加滚动区域
        # 创建滚动区域
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("""
            QScrollArea {
                border: 1px solid #ccc;
                background: white;
            }
        """)
        
        # 创建滚动区域的内容容器
        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        scroll_area.setWidget(self.scroll_content)
        
        main_layout.addWidget(scroll_area)

    def create_copy_button(self, text):
        """创建复制按钮"""
        copy_btn = QPushButton("复制")
        copy_btn.setFixedSize(QSize(60, 25))
        copy_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                border-radius: 3px;
                padding: 5px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)
        copy_btn.clicked.connect(lambda: self.copy_to_clipboard(text))
        return copy_btn

    def copy_to_clipboard(self, text):
        """复制文本到剪贴板"""
        clipboard = QApplication.clipboard()
        clipboard.setText(text)
        QMessageBox.information(self, "成功", "已复制到剪贴板")

    def update_wallet_display(self):
        """更新钱包显示区域"""
        # 清除现有内容
        while self.scroll_layout.count():
            child = self.scroll_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        # 添加新的钱包信息
        for i, wallet in enumerate(self.wallet_data, 1):
            # 钱包标题
            wallet_title = QLabel(f"钱包 {i}")
            wallet_title.setStyleSheet("font-weight: bold; color: #2196F3; margin-top: 10px;")
            self.scroll_layout.addWidget(wallet_title)

            # 地址行
            addr_widget = QWidget()
            addr_layout = QHBoxLayout(addr_widget)
            addr_layout.setContentsMargins(0, 0, 0, 0)
            addr_layout.setSpacing(10)  # 设置间距
            
            # 创建一个容器来包含地址文本
            addr_text_container = QWidget()
            addr_text_layout = QHBoxLayout(addr_text_container)
            addr_text_layout.setContentsMargins(0, 0, 0, 0)
            
            addr_label = QLabel(f"<span style='font-weight: bold;'>地址: {wallet['address']}</span>")
            addr_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
            addr_label.setWordWrap(True)  # 允许文本换行
            addr_text_layout.addWidget(addr_label)
            
            # 将地址文本容器添加到主布局
            addr_layout.addWidget(addr_text_container, stretch=1)  # stretch=1 使其占据所有可用空间
            
            # 添加复制按钮，不设置stretch，保持固定大小
            copy_btn = self.create_copy_button(wallet['address'])
            addr_layout.addWidget(copy_btn, alignment=Qt.AlignRight | Qt.AlignVCenter)
            
            self.scroll_layout.addWidget(addr_widget)

            # 私钥行
            key_widget = QWidget()
            key_layout = QHBoxLayout(key_widget)
            key_layout.setContentsMargins(0, 0, 0, 0)
            key_layout.setSpacing(10)
            
            # 创建一个容器来包含私钥文本
            key_text_container = QWidget()
            key_text_layout = QHBoxLayout(key_text_container)
            key_text_layout.setContentsMargins(0, 0, 0, 0)
            
            key_label = QLabel(f"<span style='font-weight: bold;'>私钥: {wallet['private_key']}</span>")
            key_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
            key_label.setWordWrap(True)
            key_text_layout.addWidget(key_label)
            
            key_layout.addWidget(key_text_container, stretch=1)
            key_layout.addWidget(self.create_copy_button(wallet['private_key']), 
                               alignment=Qt.AlignRight | Qt.AlignVCenter)
            
            self.scroll_layout.addWidget(key_widget)

            # 助记词行
            mnemonic_widget = QWidget()
            mnemonic_layout = QHBoxLayout(mnemonic_widget)
            mnemonic_layout.setContentsMargins(0, 0, 0, 0)
            mnemonic_layout.setSpacing(10)
            
            # 创建一个容器来包含助记词文本
            mnemonic_text_container = QWidget()
            mnemonic_text_layout = QHBoxLayout(mnemonic_text_container)
            mnemonic_text_layout.setContentsMargins(0, 0, 0, 0)
            
            mnemonic_label = QLabel(f"<span style='font-weight: bold;'>助记词: {wallet['mnemonic']}</span>")
            mnemonic_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
            mnemonic_label.setWordWrap(True)
            mnemonic_text_layout.addWidget(mnemonic_label)
            
            mnemonic_layout.addWidget(mnemonic_text_container, stretch=1)
            mnemonic_layout.addWidget(self.create_copy_button(wallet['mnemonic']), 
                                    alignment=Qt.AlignRight | Qt.AlignVCenter)
            
            self.scroll_layout.addWidget(mnemonic_widget)

            # 添加分隔线
            if i < len(self.wallet_data):
                line = QFrame()
                line.setFrameShape(QFrame.HLine)
                line.setFrameShadow(QFrame.Sunken)
                line.setStyleSheet("background-color: #ccc;")
                self.scroll_layout.addWidget(line)

        # 添加弹性空间
        self.scroll_layout.addStretch()

    def generate_wallets(self):
        try:
            num_wallets = int(self.amount_input.text())
            if num_wallets <= 0:
                QMessageBox.critical(self, "错误", "请输入大于0的数字")
                return
            
            self.wallet_data = []
            # 获取选中的链
            chain = self.chain_group.checkedButton().text()
            # 获取选中的助记词长度
            mnemonic_length = self.mnemonic_group.checkedButton().text()
            strength = MNEMONIC_STRENGTHS[mnemonic_length]
            
            for _ in range(num_wallets):
                # 生成助记词
                mnemo = Mnemonic("english")
                mnemonic = mnemo.generate(strength=strength)
                
                # 生成钱包
                account = self.w3.eth.account.create()
                
                self.wallet_data.append({
                    'address': account.address,
                    'private_key': account.key.hex(),
                    'mnemonic': mnemonic
                })
            
            # 在生成完成后更新显示
            self.update_wallet_display()
            
            QMessageBox.information(self, "成功", 
                f"成功生成 {num_wallets} 个钱包!\n请点击下载表格保存。")
            
        except ValueError:
            QMessageBox.critical(self, "错误", "请输入有效的数字")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成钱包时发生错误: {str(e)}")

    def save_to_excel(self):
        if not self.wallet_data:
            QMessageBox.critical(self, "错误", "请先生成钱包")
            return
        
        try:
            # 获取保存路径
            filename, _ = QFileDialog.getSaveFileName(
                self,
                "保存钱包信息",
                "wallets.xlsx",
                "Excel Files (*.xlsx)"
            )
            
            if filename:
                df = pd.DataFrame(self.wallet_data)
                df.to_excel(filename, index=False)
                QMessageBox.information(self, "成功", f"文件已保存至:\n{filename}")
                
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存文件时发生错误: {str(e)}")

if __name__ == '__main__':
    # 确保安装了必要的库
    try:
        from web3 import Web3
        import pandas as pd
        from mnemonic import Mnemonic
    except ImportError:
        print("请先安装必要的库:")
        print("pip install web3 pandas mnemonic PyQt5 openpyxl")
        sys.exit(1)
        
    app = QApplication(sys.argv)
    window = WalletGenerator()
    window.show()
    sys.exit(app.exec_())
