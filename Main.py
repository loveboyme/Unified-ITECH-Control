import sys
import time
import pyvisa
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QGridLayout, QLabel, QLineEdit, QPushButton, QTextEdit,
                             QDoubleSpinBox, QMessageBox, QGroupBox, QSizePolicy, QFrame)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject, pyqtSlot, QTimer # 导入 QTimer
from PyQt5.QtGui import QPalette, QColor, QFont, QIcon

# --- 设备工作器 (在独立线程中执行VISA操作) ---
class DeviceWorker(QObject):
    # 工作器 -> GUI 的信号
    connected = pyqtSignal(str, str) # 设备名称, IDN字符串
    disconnected = pyqtSignal(str) # 设备名称
    error = pyqtSignal(str, str, str) # 设备名称, 错误标题, 错误信息
    log_message_signal = pyqtSignal(str) # 日志消息

    # 电源特定更新信号
    ps_settings_updated = pyqtSignal(float, float) # 设定电压, 设定电流上限
    ps_measurements_updated = pyqtSignal(float, float, float) # 测量电压, 测量电流, 测量功率
    ps_output_status_updated = pyqtSignal(str) # "ON", "OFF" 或其他状态

    # 电子负载特定更新信号
    el_settings_updated = pyqtSignal(float) # 设定电流
    el_measurements_updated = pyqtSignal(float, float, float) # 测量电压, 测量电流, 测量功率
    el_input_status_updated = pyqtSignal(str) # "ON", "OFF" 或其他状态

    def __init__(self, rm, resource_string, device_name):
        super().__init__()
        self.rm = rm
        self.resource_string = resource_string
        self.device_name = device_name
        self.instrument = None
        self._is_connected = False

    @pyqtSlot()
    def connect_device(self):
        """尝试连接VISA设备。此方法将在工作线程中执行。"""
        if self._is_connected:
            self.log_message_signal.emit(f"<font color='#FFC107'>警告:</font> {self.device_name} 已经连接。") # 警告色
            return

        self.log_message_signal.emit(f"正在连接到 <font color='#17A2B8'>{self.device_name}</font> (<font color='#17A2B8'>{self.resource_string}</font>)...") # 信息色
        try:
            self.instrument = self.rm.open_resource(self.resource_string)
            self.instrument.timeout = 5000  # 5秒超时
            self.instrument.read_termination = '\n'
            self.instrument.write_termination = '\n'
            self.instrument.write("*CLS") # 清除设备错误
            idn = self.instrument.query("*IDN?").strip()
            self.log_message_signal.emit(f"<font color='#28A745'>成功:</font> {self.device_name} 已连接: {idn}") # 成功色
            
            try:
                # 尝试进入远程模式，某些设备可能不需要或不支持
                self.instrument.write("SYST:REM")
                self.log_message_signal.emit(f"{self.device_name}: 已发送 <font color='#17A2B8'>SYST:REM</font> (进入远程模式) 命令。") # 信息色
            except pyvisa.errors.VisaIOError as e:
                self.log_message_signal.emit(f"<font color='#FFC107'>警告:</font> {self.device_name}: 发送 SYST:REM 命令失败 (可能不需要): {e}") # 警告色

            self._is_connected = True
            self.connected.emit(self.device_name, idn)
        except pyvisa.errors.VisaIOError as e:
            err_msg = f"连接 {self.device_name} 失败: {e}"
            self.log_message_signal.emit(f"<font color='#DC3545'>错误:</font> {err_msg}") # 错误色
            self.error.emit(self.device_name, f"连接 {self.device_name} 错误", err_msg)
            if self.instrument: self.instrument.close()
            self.instrument = None
            self._is_connected = False
        except Exception as e:
            err_msg = f"连接 {self.device_name} 时发生未知错误: {e}"
            self.log_message_signal.emit(f"<font color='#DC3545'>错误:</font> {err_msg}") # 错误色
            self.error.emit(self.device_name, f"{self.device_name} 未知错误", err_msg)
            if self.instrument: self.instrument.close()
            self.instrument = None
            self._is_connected = False

    @pyqtSlot()
    def disconnect_device(self):
        """断开VISA设备连接。此方法将在工作线程中执行。"""
        if not self._is_connected or not self.instrument:
            self.log_message_signal.emit(f"<font color='#FFC107'>警告:</font> {self.device_name} 未连接，无需断开。") # 警告色
            return

        self.log_message_signal.emit(f"正在断开 <font color='#17A2B8'>{self.device_name}</font>...") # 信息色
        try:
            try:
                # 尝试返回本地模式
                self.instrument.write("SYST:LOC")
                self.log_message_signal.emit(f"{self.device_name}: 已发送 <font color='#17A2B8'>SYST:LOC</font> (返回本地模式) 命令。") # 信息色
            except pyvisa.errors.VisaIOError as e:
                self.log_message_signal.emit(f"<font color='#FFC107'>警告:</font> {self.device_name}: 发送 SYST:LOC 命令失败 (可能设备已不响应或无此命令): {e}") # 警告色
            self.instrument.close()
            self.instrument = None
            self._is_connected = False
            self.log_message_signal.emit(f"<font color='#28A745'>成功:</font> {self.device_name} 设备已断开。</font>") # 成功色
            self.disconnected.emit(self.device_name)
        except Exception as e:
            err_msg = f"断开 {self.device_name} 时发生错误: {e}"
            self.log_message_signal.emit(f"<font color='#DC3545'>错误:</font> {err_msg}") # 错误色
            self.error.emit(self.device_name, f"{self.device_name} 断开错误", err_msg)

    @pyqtSlot(str, str, bool, str)
    def process_command(self, command, param, is_query, caller_id=""):
        """
        在工作线程中处理SCPI命令。
        :param command: SCPI命令 (例如 "VOLT")
        :param param: 命令参数 (例如 "1.0")
        :param is_query: True表示查询命令 (例如 "VOLT?")
        :param caller_id: 用于区分不同命令类型，以便精确地发射更新信号
        """
        if not self._is_connected or not self.instrument:
            self.log_message_signal.emit(f"<font color='#DC3545'>错误:</font> {self.device_name} 未连接，无法执行命令: {command} {param}") # 错误色
            self.error.emit(self.device_name, f"{self.device_name} 命令错误", f"{self.device_name} 未连接。")
            return

        full_command = f"{command} {param}" if param else command
        try:
            self.log_message_signal.emit(f"{self.device_name} 发送: <font color='#17A2B8'>{full_command}</font>") # 信息色
            if is_query:
                response = self.instrument.query(full_command).strip()
                self.log_message_signal.emit(f"{self.device_name} 收到: <font color='#20B2AA'>{response}</font>") # 成功色 (偏青)
                
                # 根据设备类型和命令发射特定信号
                if self.device_name == "电源":
                    if command == "VOLT?":
                        self.ps_settings_updated.emit(float(response), -1.0) # -1.0表示不更新电流
                    elif command == "CURR?":
                        self.ps_settings_updated.emit(-1.0, float(response)) # -1.0表示不更新电压
                    elif command == "OUTP?":
                        self.ps_output_status_updated.emit(response)
                    elif command == "MEAS:VOLT?":
                         # 测量值查询通常在刷新时批量处理，这里可以单独处理，但为了避免重复刷新，只在需要时更新UI
                         pass 
                elif self.device_name == "电子负载":
                    if command == "CURR?":
                        self.el_settings_updated.emit(float(response))
                    elif command == "INP?":
                        self.el_input_status_updated.emit(response)
                    elif command == "MEAS:VOLT?":
                         pass
            else: # 写入命令
                self.instrument.write(full_command)
                self.log_message_signal.emit(f"<font color='#28A745'>成功:</font> {self.device_name}: 命令 '{full_command}' 已发送成功。</font>") # 成功色
                # 写入命令后，立即查询状态以更新UI
                if self.device_name == "电源" and command == "OUTP":
                    # 短暂延迟以确保设备处理完写入命令
                    time.sleep(0.1) 
                    self.process_command("OUTP?", "", True, "query_output_status")
                elif self.device_name == "电子负载" and command == "INP":
                    time.sleep(0.1)
                    self.process_command("INP?", "", True, "query_input_status")

        except pyvisa.errors.VisaIOError as e:
            err_msg = f"执行 '{full_command}' 失败: {e}"
            self.log_message_signal.emit(f"<font color='#DC3545'>错误:</font> {self.device_name} 命令 ({full_command}) VISA I/O 错误: {e}</font>") # 错误色
            self.error.emit(self.device_name, f"{self.device_name} 命令错误", err_msg)
        except ValueError:
            self.log_message_signal.emit(f"<font color='#DC3545'>错误:</font> {self.device_name} 响应 '{response}' 数据解析失败。</font>") # 错误色
            self.error.emit(self.device_name, f"{self.device_name} 数据解析错误", f"无法解析设备响应: {response}")
        except Exception as e:
            err_msg = f"执行 '{full_command}' 时发生未知错误: {e}"
            self.log_message_signal.emit(f"<font color='#DC3545'>错误:</font> {self.device_name} 命令 ({full_command}) 未知错误: {e}</font>") # 错误色
            self.error.emit(self.device_name, f"{self.device_name} 未知错误", err_msg)

    @pyqtSlot()
    def refresh_status_and_measurements(self):
        """
        在工作线程中刷新设备的设定值、状态和测量值。
        """
        if not self._is_connected or not self.instrument:
            self.log_message_signal.emit(f"<font color='#FFC107'>警告:</font> {self.device_name} 未连接，无法刷新。") # 警告色
            return

        self.log_message_signal.emit(f"正在刷新 <font color='#17A2B8'>{self.device_name}</font> 所有状态和测量值...") # 信息色
        try:
            if self.device_name == "电源":
                v_set_str = self.instrument.query("VOLT?").strip()
                time.sleep(0.05)
                i_set_str = self.instrument.query("CURR?").strip()
                self.ps_settings_updated.emit(float(v_set_str), float(i_set_str))
                
                time.sleep(0.05)
                outp_status_str = self.instrument.query("OUTP?").strip()
                self.ps_output_status_updated.emit(outp_status_str)

                time.sleep(0.05)
                v_meas = float(self.instrument.query("MEAS:VOLT?").strip())
                time.sleep(0.05)
                i_meas = float(self.instrument.query("MEAS:CURR?").strip())
                time.sleep(0.05)
                p_meas = float(self.instrument.query("MEAS:POW?").strip())
                self.ps_measurements_updated.emit(v_meas, i_meas, p_meas)
                self.log_message_signal.emit(f"<font color='#28A745'>成功:</font> {self.device_name}: 刷新完成。</font>") # 成功色

            elif self.device_name == "电子负载":
                i_set_str = self.instrument.query("CURR?").strip()
                self.el_settings_updated.emit(float(i_set_str))
                
                time.sleep(0.05)
                inp_status_str = self.instrument.query("INP?").strip()
                self.el_input_status_updated.emit(inp_status_str)

                time.sleep(0.05)
                v_meas = float(self.instrument.query("MEAS:VOLT?").strip())
                time.sleep(0.05)
                i_meas = float(self.instrument.query("MEAS:CURR?").strip())
                time.sleep(0.05)
                p_meas = float(self.instrument.query("MEAS:POW?").strip())
                self.el_measurements_updated.emit(v_meas, i_meas, p_meas)
                self.log_message_signal.emit(f"<font color='#28A745'>成功:</font> {self.device_name}: 刷新完成。</font>") # 成功色
        except pyvisa.errors.VisaIOError as e:
            err_msg = f"刷新 {self.device_name} 状态失败: {e}"
            self.log_message_signal.emit(f"<font color='#DC3545'>错误:</font> {err_msg}") # 错误色
            self.error.emit(self.device_name, f"{self.device_name} 刷新错误", err_msg)
        except ValueError:
            self.log_message_signal.emit(f"<font color='#DC3545'>错误:</font> {self.device_name} 刷新数据解析失败。</font>") # 错误色
            self.error.emit(self.device_name, f"{self.device_name} 刷新错误", "无法解析刷新数据。")
        except Exception as e:
            err_msg = f"刷新 {self.device_name} 状态时发生未知错误: {e}"
            self.log_message_signal.emit(f"<font color='#DC3545'>错误:</font> {err_msg}") # 错误色
            self.error.emit(self.device_name, f"{self.device_name} 刷新错误", err_msg)

# --- 主GUI窗口 ---
class UnifiedControllerGUI(QMainWindow):
    # GUI -> 工作器 的信号
    ps_connect_request = pyqtSignal()
    ps_disconnect_request = pyqtSignal()
    ps_command_request = pyqtSignal(str, str, bool, str) # command, param, is_query, caller_id
    ps_refresh_request = pyqtSignal()

    el_connect_request = pyqtSignal()
    el_disconnect_request = pyqtSignal()
    el_command_request = pyqtSignal(str, str, bool, str)
    el_refresh_request = pyqtSignal()

    DEFAULT_PS_VISA_RESOURCE = 'USB0::0x2EC7::0x6000::803982200797740009::INSTR' 
    DEFAULT_EL_VISA_RESOURCE = 'USB0::0x2EC7::0x8900::803280023806740001::INSTR' 

    def __init__(self):
        super().__init__()
        self.setWindowTitle("ITECH 设备统一控制器 (IT6000 & IT8902E)")
        self.setGeometry(100, 100, 1050, 850) # 调整窗口大小，留出更多空间

        self.rm = pyvisa.ResourceManager()
        self.ps_worker = None
        self.ps_thread = None
        self.el_worker = None
        self.el_thread = None

        self.ps_voltage_default_on_connect = 1.0
        self.ps_current_limit_default_on_connect = 0.1
        self.el_current_default_on_connect = 1.0

        # Store original class properties for measurement labels for flashing
        # 这些将用于在闪烁后恢复标签的原始QSS类样式（包括警告色）
        self._ps_voltage_label_original_class = "measurement_value"
        self._ps_current_label_original_class = "measurement_value"
        self._ps_power_label_original_class = "measurement_value"
        self._el_voltage_label_original_class = "measurement_value"
        self._el_current_label_original_class = "measurement_value"
        self._el_power_label_original_class = "measurement_value"

        self.init_ui()
        self.apply_stylesheet()
        self.log_message("<font color='#666666'>应用程序已启动。</font> 请连接设备。") # 普通日志颜色

    def apply_stylesheet(self):
        """应用新的浅色主题和现代化UI样式。"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #F0F4F8; /* 浅灰色/蓝灰色主背景 */
                color: #333333; /* 默认文字颜色 */
            }
            QGroupBox {
                background-color: #E0E6F0; /* 组框背景色，比主背景稍暗 */
                color: #333333;
                border: 1px solid #C0D0E0; /* 组框边框 */
                border-radius: 8px;
                margin-top: 2.5ex; /* 顶部留出标题空间 */
                font-weight: bold;
                font-size: 11pt;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 15px; /* 增加标题内边距 */
                background-color: #607B8B; /* 标题背景色 ( muted blue-gray) */
                color: #FFFFFF; /* 标题文字颜色 */
                border-radius: 6px;
                font-size: 12pt; /* 增大标题字体 */
            }
            QLabel {
                color: #333333; /* 普通标签文字颜色 */
                font-size: 10pt;
            }
            QLineEdit, QDoubleSpinBox {
                background-color: #FFFFFF; /* 输入框背景色 */
                color: #333333;
                border: 1px solid #B0C4DE; /* 输入框边框 (Light Steel Blue) */
                border-radius: 5px;
                padding: 10px; /* 增加内边距 */
                font-size: 11pt; /* 增大字体 */
            }
            QLineEdit:disabled, QDoubleSpinBox:disabled {
                background-color: #E8E8E8;
                color: #888888;
                border: 1px dashed #D0D0D0; /* Disabled border style */
            }
            QPushButton {
                background-color: #4682B4; /* 按钮默认蓝色 (Steel Blue) */
                color: white;
                border: none;
                border-radius: 5px;
                padding: 12px 24px; /* 增加按钮内边距 */
                font-size: 11pt; /* 增大字体 */
                font-weight: bold;
                min-width: 120px; /* 增加最小宽度 */
                min-height: 35px; /* 增加最小高度 */
                /* QSS的'transition'属性在PyQt中可能不完全支持或导致警告，
                   故为避免“Unknown property transition”警告，此处已移除。
                   按钮的颜色变化将为即时切换。 */
            }
            QPushButton:hover {
                background-color: #5A9BD6; /* 鼠标悬停时更亮的蓝色 */
            }
            QPushButton:pressed {
                background-color: #3A6F9B; /* 鼠标按下时更暗的蓝色 */
            }
            QPushButton:disabled {
                background-color: #AAAAAA;
                color: #666666;
            }
            QTextEdit {
                background-color: #FFFFFF; /* 日志区背景色 */
                color: #333333; /* 日志区文字颜色 */
                border: 1px solid #B0C4DE;
                border-radius: 5px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 9pt;
                padding: 10px; /* 增加内边距 */
            }
            /* 状态指示器标签 */
            QLabel.status_indicator {
                min-width: 24px; /* 增大指示器尺寸 */
                min-height: 24px;
                max-width: 24px;
                max-height: 24px;
                border-radius: 12px; /* 圆形 */
                background-color: #6C757D; /* 默认灰色 (未知) */
                border: 1px solid #999999;
            }
            QLabel.status_indicator.on {
                background-color: #28A745; /* 绿色 */
            }
            QLabel.status_indicator.off {
                background-color: #DC3545; /* 红色 */
            }
            QLabel.status_indicator.connecting {
                background-color: #FFC107; /* 琥珀色 (连接中) */
            }
            QLabel.status_indicator.unknown {
                background-color: #6C757D; /* 灰色 */
            }
            /* 测量值标签样式 - 醒目化 */
            QLabel.measurement_value {
                font-weight: bold;
                color: #007BFF; /* 醒目蓝色 */
                font-size: 14pt; /* 进一步增大字体 */
                padding: 4px 0; /* 增加垂直内边距 */
            }
            QLabel.measurement_value.warning { /* 零值或异常值 */
                color: #DC3545; /* 红色 */
            }
            /* 分隔线 */
            QFrame#horizontalLine {
                background-color: #A9B1BB; /* 浅色主题下的分隔线颜色 */
                height: 1px;
                margin: 10px 0; /* 增加上下间距 */
            }
        """)

    def init_ui(self):
        """初始化用户界面布局和控件。"""
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(20, 20, 20, 20) # 增加整体边距
        main_layout.setSpacing(25) # 增加主要布局间距

        # --- 连接面板 ---
        connection_panel = QGroupBox("设备连接")
        connection_layout = QGridLayout()
        connection_layout.setContentsMargins(20, 30, 20, 20)
        connection_layout.setVerticalSpacing(15) # 增加垂直间距
        connection_layout.setHorizontalSpacing(20) # 增加水平间距

        # 电源连接区域
        connection_layout.addWidget(QLabel("电源 VISA 地址:"), 0, 0, Qt.AlignRight)
        self.ps_visa_entry = QLineEdit(self.DEFAULT_PS_VISA_RESOURCE)
        connection_layout.addWidget(self.ps_visa_entry, 0, 1)
        self.ps_connect_button = QPushButton("连接电源")
        self.ps_connect_button.clicked.connect(self.toggle_ps_connection)
        connection_layout.addWidget(self.ps_connect_button, 0, 2)
        
        self.ps_idn_label = QLabel("电源 IDN: 未连接")
        connection_layout.addWidget(self.ps_idn_label, 1, 0, 1, 3) # 占据三列

        # 电子负载连接区域
        connection_layout.addWidget(QLabel("负载 VISA 地址:"), 2, 0, Qt.AlignRight)
        self.el_visa_entry = QLineEdit(self.DEFAULT_EL_VISA_RESOURCE)
        connection_layout.addWidget(self.el_visa_entry, 2, 1)
        self.el_connect_button = QPushButton("连接负载")
        self.el_connect_button.clicked.connect(self.toggle_el_connection)
        connection_layout.addWidget(self.el_connect_button, 2, 2)
        
        self.el_idn_label = QLabel("电子负载 IDN: 未连接")
        connection_layout.addWidget(self.el_idn_label, 3, 0, 1, 3)
        
        connection_panel.setLayout(connection_layout)
        main_layout.addWidget(connection_panel)

        # --- 控制面板 (电源和负载并排) ---
        controls_layout = QHBoxLayout()
        controls_layout.setSpacing(30) # 增加组框之间的间距

        # 电源控制组
        ps_controls_group = QGroupBox("电源控制 (IT6000)")
        ps_grid = QGridLayout()
        ps_grid.setContentsMargins(20, 30, 20, 20)
        ps_grid.setVerticalSpacing(12) # 增加垂直间距
        ps_grid.setHorizontalSpacing(20) # 增加水平间距

        ps_grid.addWidget(QLabel("设定电压 (V):"), 0, 0, Qt.AlignRight)
        self.ps_voltage_spinbox = QDoubleSpinBox()
        self.ps_voltage_spinbox.setRange(0, 60.0) # 根据IT6000型号调整
        self.ps_voltage_spinbox.setDecimals(3)
        self.ps_voltage_spinbox.setValue(self.ps_voltage_default_on_connect)
        ps_grid.addWidget(self.ps_voltage_spinbox, 0, 1)
        self.ps_set_voltage_button = QPushButton("设定电压")
        self.ps_set_voltage_button.clicked.connect(self.set_ps_voltage)
        ps_grid.addWidget(self.ps_set_voltage_button, 0, 2)

        ps_grid.addWidget(QLabel("设定电流上限 (A):"), 1, 0, Qt.AlignRight)
        self.ps_current_limit_spinbox = QDoubleSpinBox()
        self.ps_current_limit_spinbox.setRange(0, 180.0) # 根据IT6000型号调整 (例如IT6018B是180A)
        self.ps_current_limit_spinbox.setDecimals(3)
        self.ps_current_limit_spinbox.setValue(self.ps_current_limit_default_on_connect)
        ps_grid.addWidget(self.ps_current_limit_spinbox, 1, 1)
        self.ps_set_current_limit_button = QPushButton("设定电流上限")
        self.ps_set_current_limit_button.clicked.connect(self.set_ps_current_limit)
        ps_grid.addWidget(self.ps_set_current_limit_button, 1, 2)
        
        # 输出控制和状态
        ps_grid.addWidget(QLabel("电源输出:"), 2, 0, Qt.AlignRight)
        output_control_layout = QHBoxLayout()
        output_control_layout.setSpacing(10) # 按钮间距
        self.ps_output_on_button = QPushButton("打开输出")
        self.ps_output_on_button.clicked.connect(lambda: self.set_ps_output_state(True))
        output_control_layout.addWidget(self.ps_output_on_button)
        self.ps_output_off_button = QPushButton("关闭输出")
        self.ps_output_off_button.clicked.connect(lambda: self.set_ps_output_state(False))
        output_control_layout.addWidget(self.ps_output_off_button)
        ps_grid.addLayout(output_control_layout, 2, 1)

        output_status_layout = QHBoxLayout()
        output_status_layout.setSpacing(8) # 指示器和标签间距
        self.ps_output_status_indicator = QLabel()
        self.ps_output_status_indicator.setProperty("class", "status_indicator unknown")
        self.ps_output_status_indicator._is_pulsing = False # Custom flag for pulsing animation
        output_status_layout.addWidget(self.ps_output_status_indicator, alignment=Qt.AlignRight)
        self.ps_output_status_label = QLabel("未知")
        output_status_layout.addWidget(self.ps_output_status_label)
        output_status_layout.addStretch(1) # 确保标签不会被挤压
        ps_grid.addLayout(output_status_layout, 2, 2, Qt.AlignLeft)

        # 测量值显示
        ps_grid.addWidget(self.create_horizontal_line(), 3, 0, 1, 3) # 分隔线占据三列

        ps_grid.addWidget(QLabel("实际电压 (V):"), 4, 0, Qt.AlignRight)
        self.ps_measured_voltage_label = QLabel("---")
        self.ps_measured_voltage_label.setProperty("class", self._ps_voltage_label_original_class)
        ps_grid.addWidget(self.ps_measured_voltage_label, 4, 1, 1, 2)

        ps_grid.addWidget(QLabel("实际电流 (A):"), 5, 0, Qt.AlignRight)
        self.ps_measured_current_label = QLabel("---")
        self.ps_measured_current_label.setProperty("class", self._ps_current_label_original_class)
        ps_grid.addWidget(self.ps_measured_current_label, 5, 1, 1, 2)
        
        ps_grid.addWidget(QLabel("实际功率 (W):"), 6, 0, Qt.AlignRight)
        self.ps_measured_power_label = QLabel("---")
        self.ps_measured_power_label.setProperty("class", self._ps_power_label_original_class)
        ps_grid.addWidget(self.ps_measured_power_label, 6, 1, 1, 2)
        
        self.ps_refresh_button = QPushButton("刷新电源状态")
        self.ps_refresh_button.clicked.connect(self.refresh_ps_status)
        ps_grid.addWidget(self.ps_refresh_button, 7, 0, 1, 3)

        ps_controls_group.setLayout(ps_grid)
        controls_layout.addWidget(ps_controls_group)
        self.set_ps_controls_enabled(False) # 初始禁用

        # 电子负载控制组
        el_controls_group = QGroupBox("电子负载控制 (IT8902E)")
        el_grid = QGridLayout()
        el_grid.setContentsMargins(20, 30, 20, 20)
        el_grid.setVerticalSpacing(12)
        el_grid.setHorizontalSpacing(20)

        el_grid.addWidget(QLabel("设定电流 (A):"), 0, 0, Qt.AlignRight)
        self.el_current_spinbox = QDoubleSpinBox()
        self.el_current_spinbox.setRange(0, 240.0)  # 根据IT8902E型号调整 (例如240A)
        self.el_current_spinbox.setDecimals(3)
        self.el_current_spinbox.setValue(self.el_current_default_on_connect)
        el_grid.addWidget(self.el_current_spinbox, 0, 1)
        self.el_set_current_button = QPushButton("设定负载电流")
        self.el_set_current_button.clicked.connect(self.set_el_current)
        el_grid.addWidget(self.el_set_current_button, 0, 2)

        # 负载输入控制和状态
        el_grid.addWidget(QLabel("负载输入:"), 1, 0, Qt.AlignRight)
        input_control_layout = QHBoxLayout()
        input_control_layout.setSpacing(10)
        self.el_input_on_button = QPushButton("打开负载")
        self.el_input_on_button.clicked.connect(lambda: self.set_el_input_state(True))
        input_control_layout.addWidget(self.el_input_on_button)
        self.el_input_off_button = QPushButton("关闭负载")
        self.el_input_off_button.clicked.connect(lambda: self.set_el_input_state(False))
        input_control_layout.addWidget(self.el_input_off_button)
        el_grid.addLayout(input_control_layout, 1, 1)

        input_status_layout = QHBoxLayout()
        input_status_layout.setSpacing(8)
        self.el_input_status_indicator = QLabel()
        self.el_input_status_indicator.setProperty("class", "status_indicator unknown")
        self.el_input_status_indicator._is_pulsing = False # Custom flag for pulsing animation
        input_status_layout.addWidget(self.el_input_status_indicator, alignment=Qt.AlignRight)
        self.el_input_status_label = QLabel("未知")
        input_status_layout.addWidget(self.el_input_status_label)
        input_status_layout.addStretch(1)
        el_grid.addLayout(input_status_layout, 1, 2, Qt.AlignLeft)

        # 负载测量值显示
        el_grid.addWidget(self.create_horizontal_line(), 2, 0, 1, 3) # 分隔线

        el_grid.addWidget(QLabel("测量电压 (V):"), 3, 0, Qt.AlignRight)
        self.el_measured_voltage_label = QLabel("---")
        self.el_measured_voltage_label.setProperty("class", self._el_voltage_label_original_class)
        el_grid.addWidget(self.el_measured_voltage_label, 3, 1, 1, 2)

        el_grid.addWidget(QLabel("测量电流 (A):"), 4, 0, Qt.AlignRight)
        self.el_measured_current_label = QLabel("---")
        self.el_measured_current_label.setProperty("class", self._el_current_label_original_class)
        el_grid.addWidget(self.el_measured_current_label, 4, 1, 1, 2)

        el_grid.addWidget(QLabel("测量功率 (W):"), 5, 0, Qt.AlignRight)
        self.el_measured_power_label = QLabel("---")
        self.el_measured_power_label.setProperty("class", self._el_power_label_original_class)
        el_grid.addWidget(self.el_measured_power_label, 5, 1, 1, 2)

        self.el_refresh_button = QPushButton("刷新负载状态")
        self.el_refresh_button.clicked.connect(self.refresh_el_status)
        el_grid.addWidget(self.el_refresh_button, 6, 0, 1, 3)

        el_controls_group.setLayout(el_grid)
        controls_layout.addWidget(el_controls_group)
        self.set_el_controls_enabled(False) # 初始禁用

        main_layout.addLayout(controls_layout)

        # --- 状态日志面板 ---
        log_group = QGroupBox("状态日志")
        log_layout = QVBoxLayout()
        log_layout.setContentsMargins(20, 30, 20, 20)
        self.status_log_edit = QTextEdit()
        self.status_log_edit.setReadOnly(True)
        log_layout.addWidget(self.status_log_edit)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)
        
        main_layout.setStretchFactor(log_group, 1) # 使日志区域自动扩展

    def create_horizontal_line(self):
        """创建水平分隔线。"""
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setObjectName("horizontalLine")
        return line

    def log_message(self, message):
        """在状态日志文本框中添加时间戳消息。"""
        timestamp = time.strftime('<font color="#666666">%H:%M:%S</font>') # 时间戳颜色
        self.status_log_edit.append(f"{timestamp} - {message}")
        self.status_log_edit.ensureCursorVisible() # 自动滚动到最新消息

    def update_status_indicator(self, indicator_label: QLabel, status: str):
        """更新状态指示器的颜色和文本，并处理连接中的脉动效果。"""
        # 在设置新状态前，如果当前正在脉动且新状态不是连接中，则停止脉动
        if status.upper() != "CONNECTING" and hasattr(indicator_label, '_is_pulsing') and indicator_label._is_pulsing:
            indicator_label._is_pulsing = False # 通知脉动函数停止
            indicator_label.setStyleSheet("") # 清除直接样式表，让QSS类样式生效
            # 确保类属性已正确设置到新状态，以便QSS应用正确的颜色
            indicator_label.setProperty("class", f"status_indicator {status.lower()}")
            indicator_label.style().polish(indicator_label) # 强制刷新样式
            
        indicator_label.setText("") # 清空文本
        if status.upper() == "ON" or status == "1":
            indicator_label.setProperty("class", "status_indicator on")
        elif status.upper() == "OFF" or status == "0":
            indicator_label.setProperty("class", "status_indicator off")
        elif status.upper() == "CONNECTING":
            indicator_label.setProperty("class", "status_indicator connecting")
            self.pulse_indicator(indicator_label, "#FFC107") # 启动脉动 (琥珀色)
        else: # UNKNOWN
            indicator_label.setProperty("class", "status_indicator unknown")
        
        indicator_label.style().polish(indicator_label) # 强制样式刷新

    def pulse_indicator(self, indicator_label: QLabel, target_color_str: str):
        """使指示器在连接状态下脉动。"""
        # 只有当指示器确实处于“connecting”状态且未在脉动时才启动
        if not indicator_label.property("class") == "status_indicator connecting":
            indicator_label._is_pulsing = False # 确保标志正确
            return

        if not hasattr(indicator_label, '_is_pulsing') or not indicator_label._is_pulsing:
            indicator_label._is_pulsing = True # 标记为正在脉动
            
        pulse_color_start = QColor(target_color_str).darker(150) # 较暗的颜色
        pulse_color_end = QColor(target_color_str) # 原始颜色

        def toggle_pulse_color():
            if not indicator_label._is_pulsing: # 检查标志，如果为False则停止脉动
                indicator_label.setStyleSheet("") # 清除直接样式，恢复到QSS类样式
                indicator_label.style().polish(indicator_label)
                return

            current_bg_style = indicator_label.styleSheet()
            # 检查当前背景色是否为“结束色”，如果是，则切换到“开始色”
            # 注意：这里通过字符串匹配，可能会有格式问题，但通常有效
            if f"background-color: {pulse_color_end.name()}" in current_bg_style.replace(";", "").replace(" ", ""): 
                indicator_label.setStyleSheet(f"background-color: {pulse_color_start.name()}; border-radius: 12px; border: 1px solid #999999;") # 保持边框
            else: # 否则（包括没有直接样式或处于开始色），切换到“结束色”
                indicator_label.setStyleSheet(f"background-color: {pulse_color_end.name()}; border-radius: 12px; border: 1px solid #999999;") # 保持边框
            
            QTimer.singleShot(500, toggle_pulse_color) # 0.5秒后再次切换颜色

        # 启动脉动（如果尚未启动）
        if indicator_label._is_pulsing:
            toggle_pulse_color()


    def flash_measurement_label(self, label: QLabel, original_class: str):
        """使测量值标签短暂闪烁以突出更新。"""
        # 应用临时的强调样式，通过直接设置样式表
        highlight_color = "#FFFF00" # 亮黄色用于闪烁
        # 从原始样式中获取 font-weight 和 font-size，并应用于闪烁样式，确保字体样式在闪烁时也保持
        font_weight = label.font().weight()
        font_size = label.font().pointSize()
        label.setStyleSheet(f"color: {highlight_color}; font-weight: {font_weight}; font-size: {font_size}pt;")
        
        # 短暂延迟后，恢复到原始样式
        QTimer.singleShot(200, lambda: self.revert_label_style(label, original_class))

    def revert_label_style(self, label: QLabel, original_class: str):
        """将测量值标签的样式恢复到其原始的QSS类状态。"""
        label.setStyleSheet("") # 清除任何直接样式表，以便QSS类规则能够生效
        label.setProperty("class", original_class) # 重新应用原始的类属性
        label.style().polish(label) # 强制重新计算和应用QSS规则


    # --- 电源 (PS) 特定功能 ---
    def toggle_ps_connection(self):
        """根据当前连接状态切换电源的连接/断开。"""
        if self.ps_worker and self.ps_worker._is_connected:
            self.disconnect_ps()
        else:
            self.connect_ps()

    def connect_ps(self):
        """启动线程连接电源设备。"""
        visa_resource = self.ps_visa_entry.text().strip()
        if not visa_resource or "USB0::" not in visa_resource: # 简单校验格式
            QMessageBox.critical(self, "连接错误", "请输入有效的电源 VISA 资源名称 (例如 'USB0::0x...::INSTR')。")
            return

        self.ps_connect_button.setEnabled(False)
        self.ps_connect_button.setText("连接中...")
        self.update_status_indicator(self.ps_output_status_indicator, "CONNECTING")
        self.ps_output_status_label.setText("连接中...")

        # 初始化并启动工作线程
        self.ps_thread = QThread()
        self.ps_worker = DeviceWorker(self.rm, visa_resource, "电源")
        self.ps_worker.moveToThread(self.ps_thread)

        # 连接工作器信号到GUI槽
        self.ps_worker.connected.connect(self.on_ps_connected_threaded)
        self.ps_worker.disconnected.connect(self.on_ps_disconnected_threaded)
        self.ps_worker.error.connect(self.show_error_message)
        self.ps_worker.log_message_signal.connect(self.log_message)
        
        # 连接GUI信号到工作器槽 (用于发送命令)
        self.ps_connect_request.connect(self.ps_worker.connect_device)
        self.ps_disconnect_request.connect(self.ps_worker.disconnect_device)
        self.ps_command_request.connect(self.ps_worker.process_command)
        self.ps_refresh_request.connect(self.ps_worker.refresh_status_and_measurements)

        # 连接工作器更新信号到GUI更新槽
        self.ps_worker.ps_settings_updated.connect(self.update_ps_settings_ui)
        self.ps_worker.ps_measurements_updated.connect(self.update_ps_measurements_ui)
        self.ps_worker.ps_output_status_updated.connect(self.update_ps_output_status_ui)

        self.ps_thread.started.connect(self.ps_connect_request) # 线程启动时发出连接请求
        self.ps_thread.start()

    @pyqtSlot(str, str)
    def on_ps_connected_threaded(self, device_name, idn):
        """电源连接成功后的回调函数 (在GUI线程中执行)。"""
        self.ps_idn_label.setText(f"电源 IDN: {idn}")
        self.ps_connect_button.setText("断开电源")
        self.ps_visa_entry.setEnabled(False)
        self.set_ps_controls_enabled(True)
        # 初始设定并刷新状态
        # 这些操作现在通过信号发送到工作线程，避免GUI阻塞
        self.ps_command_request.emit("VOLT", str(self.ps_voltage_default_on_connect), False, "set_initial_volt")
        self.ps_command_request.emit("CURR", str(self.ps_current_limit_default_on_connect), False, "set_initial_curr")
        # 触发一次全面刷新以获取所有当前状态
        self.ps_refresh_request.emit()
        self.ps_connect_button.setEnabled(True)
        # 状态指示器将在刷新后更新

    def disconnect_ps(self):
        """断开电源设备连接。"""
        if not self.ps_worker or not self.ps_worker._is_connected:
            self.log_message("<font color='#FFC107'>警告:</font> 电源已断开或未连接，无需操作。")
            return

        self.set_ps_controls_enabled(False) # 立即禁用控件
        self.ps_connect_button.setEnabled(False)
        self.ps_connect_button.setText("断开中...")
        self.update_status_indicator(self.ps_output_status_indicator, "OFF") # 立即显示为关闭状态

        # 先发送关闭输出命令，再发送断开连接请求
        self.ps_command_request.emit("OUTP", "OFF", False, "disconnect_cleanup")
        # 稍微等待确保命令被工作器接收，然后触发断开
        QThread.msleep(100) # 非阻塞等待
        self.ps_disconnect_request.emit() # 触发工作线程的断开槽

    @pyqtSlot(str)
    def on_ps_disconnected_threaded(self, device_name):
        """电源断开后的回调函数 (在GUI线程中执行)。"""
        self.ps_idn_label.setText("电源 IDN: 未连接")
        self.ps_connect_button.setText("连接电源")
        self.ps_connect_button.setEnabled(True)
        self.ps_visa_entry.setEnabled(True)
        self.set_ps_controls_enabled(False)
        self.update_status_indicator(self.ps_output_status_indicator, "UNKNOWN")
        self.ps_output_status_label.setText("未知")
        self.ps_measured_voltage_label.setText("---")
        self.ps_measured_current_label.setText("---")
        self.ps_measured_power_label.setText("---")
        # 确保样式重置为默认 (非警告，正常 measurement_value)
        self.ps_measured_voltage_label.setProperty("class", self._ps_voltage_label_original_class)
        self.ps_measured_current_label.setProperty("class", self._ps_current_label_original_class)
        self.ps_measured_power_label.setProperty("class", self._ps_power_label_original_class)
        self.ps_measured_voltage_label.style().polish(self.ps_measured_voltage_label)
        self.ps_measured_current_label.style().polish(self.ps_measured_current_label)
        self.ps_measured_power_label.style().polish(self.ps_measured_power_label)

        # 确保线程完全退出和清理
        if self.ps_thread:
            self.ps_thread.quit()
            self.ps_thread.wait(1000) # 等待线程结束，最多1秒
            self.ps_thread.deleteLater()
            self.ps_worker.deleteLater()
            self.ps_thread = None
            self.ps_worker = None

    def set_ps_controls_enabled(self, enabled):
        """设置电源控制面板中的控件启用/禁用状态。"""
        self.ps_voltage_spinbox.setEnabled(enabled)
        self.ps_set_voltage_button.setEnabled(enabled)
        self.ps_current_limit_spinbox.setEnabled(enabled)
        self.ps_set_current_limit_button.setEnabled(enabled)
        self.ps_output_on_button.setEnabled(enabled)
        self.ps_output_off_button.setEnabled(enabled)
        self.ps_refresh_button.setEnabled(enabled)

    def set_ps_voltage(self):
        """设定电源输出电压 (通过信号发送到工作线程)。"""
        if not self.ps_worker or not self.ps_worker._is_connected:
            QMessageBox.warning(self, "操作失败", "电源未连接，无法设定电压。")
            return
        voltage = self.ps_voltage_spinbox.value()
        # 校验输入是否合理
        if voltage > 50: # 假设安全电压阈值，此值可根据实际设备调整
            reply = QMessageBox.warning(self, "高电压设定警告", 
                                        f"您设定的电压 (<font color='#DC3545'><b>{voltage} V</b></font>) 较高，请确认操作安全！", # 警告色
                                        QMessageBox.Ok | QMessageBox.Cancel)
            if reply == QMessageBox.Cancel:
                self.log_message(f"<font color='#FFC107'>警告:</font> 电源: 高电压设定操作已取消。") # 警告色
                return

        self.ps_command_request.emit("VOLT", str(voltage), False, "set_voltage")

    def set_ps_current_limit(self):
        """设定电源输出电流上限 (通过信号发送到工作线程)。"""
        if not self.ps_worker or not self.ps_worker._is_connected:
            QMessageBox.warning(self, "操作失败", "电源未连接，无法设定电流上限。")
            return
        current = self.ps_current_limit_spinbox.value()
        # 校验输入是否合理
        if current > 100: # 假设大电流阈值，此值可根据实际设备调整
            reply = QMessageBox.warning(self, "高电流上限设定警告", 
                                        f"您设定的电流上限 (<font color='#DC3545'><b>{current} A</b></font>) 较高，请确认操作安全！", # 警告色
                                        QMessageBox.Ok | QMessageBox.Cancel)
            if reply == QMessageBox.Cancel:
                self.log_message(f"<font color='#FFC107'>警告:</font> 电源: 高电流上限设定操作已取消。") # 警告色
                return
        self.ps_command_request.emit("CURR", str(current), False, "set_current_limit")

    @pyqtSlot(float, float)
    def update_ps_settings_ui(self, voltage_set, current_limit_set):
        """更新电源的设定电压和电流上限UI显示。"""
        if voltage_set != -1.0: # -1.0表示此值未更新
            self.ps_voltage_spinbox.setValue(voltage_set)
        if current_limit_set != -1.0:
            self.ps_current_limit_spinbox.setValue(current_limit_set)

    def set_ps_output_state(self, state_on):
        """
        设置电源输出状态 (ON/OFF)，并提供安全确认。
        :param state_on: True为ON，False为OFF
        """
        if not self.ps_worker or not self.ps_worker._is_connected:
            QMessageBox.warning(self, "操作失败", "电源设备未连接。")
            return

        if state_on: # 打开输出时才进行风险确认
            voltage = self.ps_voltage_spinbox.value()
            current = self.ps_current_limit_spinbox.value()
            
            # 零值或低值设定下的风险提示
            if voltage <= 0.001 or current <= 0.001: # 设定值为零或接近零
                reply = QMessageBox.warning(self, "重要警告：输出设置过低", 
                                    f"当前设定电压为 <font color='#DC3545'><b>{voltage:.3f} V</b></font>，电流上限为 <font color='#DC3545'><b>{current:.3f} A</b></font>。<br>" # 警告色
                                    "在零值或过低设置下打开输出可能导致设备无法正常工作或输出异常，请确认操作无误！<br>"
                                    "是否仍要打开输出？",
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if reply == QMessageBox.No:
                    self.log_message("<font color='#FFC107'>警告:</font> 电源: 用户取消了在低设置下打开输出操作。") # 警告色
                    return

            # 打开输出的最终确认
            reply = QMessageBox.question(self, "确认操作：打开电源输出", 
                                         f"您即将打开电源输出。<br>"
                                         f"当前设定：电压 <font color='#007BFF'><b>{voltage:.3f} V</b></font>，电流上限 <font color='#007BFF'><b>{current:.3f} A</b></font>。<br>" # 醒目色
                                         f"<font color='#DC3545'><b>强烈建议在打开输出前，务必确认负载连接正确且安全！</b></font><br>" # 警告色
                                         f"是否继续打开输出？",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.No:
                self.log_message("<font color='#FFC107'>警告:</font> 电源: 用户取消了打开输出操作。") # 警告色
                return
        # 关闭输出不再确认，直接执行
        
        command_param = "ON" if state_on else "OFF"
        self.ps_command_request.emit("OUTP", command_param, False, "set_output_state")

    @pyqtSlot(str)
    def update_ps_output_status_ui(self, status):
        """更新电源输出状态的UI显示。"""
        if status == "1" or status.upper() == "ON": 
            self.update_status_indicator(self.ps_output_status_indicator, "ON")
            self.ps_output_status_label.setText("ON")
        elif status == "0" or status.upper() == "OFF": 
            self.update_status_indicator(self.ps_output_status_indicator, "OFF")
            self.ps_output_status_label.setText("OFF")
        else: 
            self.update_status_indicator(self.ps_output_status_indicator, "UNKNOWN")
            self.ps_output_status_label.setText(f"未知 ({status})")

    @pyqtSlot(float, float, float)
    def update_ps_measurements_ui(self, voltage, current, power):
        """更新电源实际输出的电压、电流和功率UI显示，并加入闪烁动效。"""
        # 检查值是否实际发生了变化，以避免不必要的闪烁
        voltage_changed = self.ps_measured_voltage_label.text() != f"{voltage:.3f}"
        current_changed = self.ps_measured_current_label.text() != f"{current:.3f}"
        power_changed = self.ps_measured_power_label.text() != f"{power:.3f}"

        self.ps_measured_voltage_label.setText(f"{voltage:.3f}")
        self.ps_measured_current_label.setText(f"{current:.3f}")
        self.ps_measured_power_label.setText(f"{power:.3f}")
        
        # 根据测量值是否过低来应用警告类
        if voltage < 0.01 and current < 0.01 and power < 0.01: 
            self.ps_measured_voltage_label.setProperty("class", "measurement_value warning")
            self.ps_measured_current_label.setProperty("class", "measurement_value warning")
            self.ps_measured_power_label.setProperty("class", "measurement_value warning")
        else:
            # 恢复到正常的measurement_value类，这将应用QSS中定义的蓝色
            self.ps_measured_voltage_label.setProperty("class", self._ps_voltage_label_original_class)
            self.ps_measured_current_label.setProperty("class", self._ps_current_label_original_class)
            self.ps_measured_power_label.setProperty("class", self._ps_power_label_original_class)

        # 强制刷新样式以应用警告/正常颜色（在潜在闪烁之前）
        self.ps_measured_voltage_label.style().polish(self.ps_measured_voltage_label)
        self.ps_measured_current_label.style().polish(self.ps_measured_current_label)
        self.ps_measured_power_label.style().polish(self.ps_measured_power_label)

        # 如果值发生变化，则应用闪烁效果（闪烁会暂时覆盖，然后恢复）
        if voltage_changed:
            # 传递当前的类属性，以便闪烁完成后能够正确恢复到警告或正常状态
            self.flash_measurement_label(self.ps_measured_voltage_label, self.ps_measured_voltage_label.property("class"))
        if current_changed:
            self.flash_measurement_label(self.ps_measured_current_label, self.ps_measured_current_label.property("class"))
        if power_changed:
            self.flash_measurement_label(self.ps_measured_power_label, self.ps_measured_power_label.property("class"))


    def refresh_ps_status(self):
        """刷新电源的所有状态和测量值 (通过信号发送到工作线程)。"""
        if not self.ps_worker or not self.ps_worker._is_connected:
            QMessageBox.warning(self, "操作失败", "电源设备未连接，无法刷新。")
            return
        self.ps_refresh_request.emit()

    # --- 电子负载 (EL) 特定功能 ---
    def toggle_el_connection(self):
        """根据当前连接状态切换电子负载的连接/断开。"""
        if self.el_worker and self.el_worker._is_connected:
            self.disconnect_el()
        else:
            self.connect_el()

    def connect_el(self):
        """启动线程连接电子负载设备。"""
        visa_resource = self.el_visa_entry.text().strip()
        if not visa_resource or "USB0::" not in visa_resource:
            QMessageBox.critical(self, "连接错误", "请输入有效的电子负载 VISA 资源名称 (例如 'USB0::0x...::INSTR')。")
            return
        
        self.el_connect_button.setEnabled(False)
        self.el_connect_button.setText("连接中...")
        self.update_status_indicator(self.el_input_status_indicator, "CONNECTING")
        self.el_input_status_label.setText("连接中...")

        self.el_thread = QThread()
        self.el_worker = DeviceWorker(self.rm, visa_resource, "电子负载")
        self.el_worker.moveToThread(self.el_thread)

        # 连接工作器信号到GUI槽
        self.el_worker.connected.connect(self.on_el_connected_threaded)
        self.el_worker.disconnected.connect(self.on_el_disconnected_threaded)
        self.el_worker.error.connect(self.show_error_message)
        self.el_worker.log_message_signal.connect(self.log_message)

        # 连接GUI信号到工作器槽 (用于发送命令)
        self.el_connect_request.connect(self.el_worker.connect_device)
        self.el_disconnect_request.connect(self.el_worker.disconnect_device)
        self.el_command_request.connect(self.el_worker.process_command)
        self.el_refresh_request.connect(self.el_worker.refresh_status_and_measurements)

        # 连接工作器更新信号到GUI更新槽
        self.el_worker.el_settings_updated.connect(self.update_el_settings_ui)
        self.el_worker.el_measurements_updated.connect(self.update_el_measurements_ui)
        self.el_worker.el_input_status_updated.connect(self.update_el_input_status_ui)

        self.el_thread.started.connect(self.el_connect_request)
        self.el_thread.start()

    @pyqtSlot(str, str)
    def on_el_connected_threaded(self, device_name, idn):
        """电子负载连接成功后的回调函数 (在GUI线程中执行)。"""
        self.el_idn_label.setText(f"电子负载 IDN: {idn}")
        self.el_connect_button.setText("断开负载")
        self.el_visa_entry.setEnabled(False)
        self.set_el_controls_enabled(True)
        # 初始设定并刷新状态
        self.el_command_request.emit("CURR", str(self.el_current_default_on_connect), False, "set_initial_curr")
        self.el_refresh_request.emit()
        self.el_connect_button.setEnabled(True)
        # 状态指示器将在刷新后更新

    def disconnect_el(self):
        """断开电子负载设备连接。"""
        if not self.el_worker or not self.el_worker._is_connected:
            self.log_message("<font color='#FFC107'>警告:</font> 电子负载已断开或未连接，无需操作。")
            return
        
        self.set_el_controls_enabled(False) # 立即禁用控件
        self.el_connect_button.setEnabled(False)
        self.el_connect_button.setText("断开中...")
        self.update_status_indicator(self.el_input_status_indicator, "OFF") # 立即显示为关闭状态

        # 先发送关闭输入命令，再发送断开连接请求
        self.el_command_request.emit("INP", "OFF", False, "disconnect_cleanup")
        QThread.msleep(100)
        self.el_disconnect_request.emit() # 触发工作线程的断开槽

    @pyqtSlot(str)
    def on_el_disconnected_threaded(self, device_name):
        """电子负载断开后的回调函数 (在GUI线程中执行)。"""
        self.el_idn_label.setText("电子负载 IDN: 未连接")
        self.el_connect_button.setText("连接负载")
        self.el_connect_button.setEnabled(True)
        self.el_visa_entry.setEnabled(True)
        self.set_el_controls_enabled(False)
        self.update_status_indicator(self.el_input_status_indicator, "UNKNOWN")
        self.el_input_status_label.setText("未知")
        self.el_measured_voltage_label.setText("---")
        self.el_measured_current_label.setText("---")
        self.el_measured_power_label.setText("---")
        # 确保样式重置为默认 (非警告，正常 measurement_value)
        self.el_measured_voltage_label.setProperty("class", self._el_voltage_label_original_class)
        self.el_measured_current_label.setProperty("class", self._el_current_label_original_class)
        self.el_measured_power_label.setProperty("class", self._el_power_label_original_class)
        self.el_measured_voltage_label.style().polish(self.el_measured_voltage_label)
        self.el_measured_current_label.style().polish(self.el_measured_current_label)
        self.el_measured_power_label.style().polish(self.el_measured_power_label)

        # 确保线程完全退出和清理
        if self.el_thread:
            self.el_thread.quit()
            self.el_thread.wait(1000)
            self.el_thread.deleteLater()
            self.el_worker.deleteLater()
            self.el_thread = None
            self.el_worker = None

    def set_el_controls_enabled(self, enabled):
        """设置电子负载控制面板中的控件启用/禁用状态。"""
        self.el_current_spinbox.setEnabled(enabled)
        self.el_set_current_button.setEnabled(enabled)
        self.el_input_on_button.setEnabled(enabled)
        self.el_input_off_button.setEnabled(enabled)
        self.el_refresh_button.setEnabled(enabled)

    def set_el_current(self):
        """设定电子负载吸收电流 (通过信号发送到工作线程)。"""
        if not self.el_worker or not self.el_worker._is_connected:
            QMessageBox.warning(self, "操作失败", "电子负载未连接，无法设定电流。")
            return
        current = self.el_current_spinbox.value()
        # 校验输入是否合理
        if current > 200: # 假设大电流阈值，此值可根据实际设备调整
            reply = QMessageBox.warning(self, "高电流设定警告", 
                                        f"您设定的电流 (<font color='#DC3545'><b>{current} A</b></font>) 较高，请确认操作安全！", # 警告色
                                        QMessageBox.Ok | QMessageBox.Cancel)
            if reply == QMessageBox.Cancel:
                self.log_message(f"<font color='#FFC107'>警告:</font> 电子负载: 高电流设定操作已取消。") # 警告色
                return
        self.el_command_request.emit("CURR", str(current), False, "set_load_current")

    @pyqtSlot(float)
    def update_el_settings_ui(self, current_set):
        """更新电子负载的设定电流UI显示。"""
        self.el_current_spinbox.setValue(current_set)

    def set_el_input_state(self, state_on):
        """
        设置电子负载输入状态 (ON/OFF)，并提供安全确认。
        :param state_on: True为ON，False为OFF
        """
        if not self.el_worker or not self.el_worker._is_connected:
            QMessageBox.warning(self, "操作失败", "电子负载设备未连接。")
            return
        
        if state_on: # 打开输入时才进行风险确认
            current = self.el_current_spinbox.value()
            
            # 零值或低值设定下的风险提示
            if current <= 0.001: # 设定电流为零或接近零
                reply = QMessageBox.warning(self, "重要警告：输入设置过低", 
                                    f"当前设定吸收电流为 <font color='#DC3545'><b>{current:.3f} A</b></font>。<br>" # 警告色
                                    "在零值或过低设置下打开负载输入可能导致无法正常加载或操作异常，请确认操作无误！<br>"
                                    "是否仍要打开输入？",
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if reply == QMessageBox.No:
                    self.log_message("<font color='#FFC107'>警告:</font> 电子负载: 用户取消了在低设置下打开输入操作。") # 警告色
                    return

            # 打开输入的最终确认
            reply = QMessageBox.question(self, "确认操作：打开电子负载输入", 
                                         f"您即将打开电子负载输入。<br>"
                                         f"当前设定吸收电流为 <font color='#007BFF'><b>{current:.3f} A</b></font>。<br>" # 醒目色
                                         f"<font color='#DC3545'><b>强烈建议在打开输入前，务必确认连接的电源已开启并稳定输出！</b></font><br>" # 警告色
                                         f"是否继续打开输入？",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.No:
                self.log_message("<font color='#FFC107'>警告:</font> 电子负载: 用户取消了打开输入操作。") # 警告色
                return
        # 关闭输入不再确认，直接执行

        command_param = "ON" if state_on else "OFF"
        self.el_command_request.emit("INP", command_param, False, "set_input_state")

    @pyqtSlot(str)
    def update_el_input_status_ui(self, status):
        """更新电子负载输入状态的UI显示。"""
        if status == "1" or status.upper() == "ON": 
            self.update_status_indicator(self.el_input_status_indicator, "ON")
            self.el_input_status_label.setText("ON")
        elif status == "0" or status.upper() == "OFF": 
            self.update_status_indicator(self.el_input_status_indicator, "OFF")
            self.el_input_status_label.setText("OFF")
        else: 
            self.update_status_indicator(self.el_input_status_indicator, "UNKNOWN")
            self.el_input_status_label.setText(f"未知 ({status})")

    @pyqtSlot(float, float, float)
    def update_el_measurements_ui(self, voltage, current, power):
        """更新电子负载实际吸收的电压、电流和功率UI显示，并加入闪烁动效。"""
        # 检查值是否实际发生了变化，以避免不必要的闪烁
        voltage_changed = self.el_measured_voltage_label.text() != f"{voltage:.3f}"
        current_changed = self.el_measured_current_label.text() != f"{current:.3f}"
        power_changed = self.el_measured_power_label.text() != f"{power:.3f}"

        self.el_measured_voltage_label.setText(f"{voltage:.3f}")
        self.el_measured_current_label.setText(f"{current:.3f}")
        self.el_measured_power_label.setText(f"{power:.3f}")
        
        # 根据测量值是否过低来应用警告类
        if voltage < 0.01 and current < 0.01 and power < 0.01: 
            self.el_measured_voltage_label.setProperty("class", "measurement_value warning")
            self.el_measured_current_label.setProperty("class", "measurement_value warning")
            self.el_measured_power_label.setProperty("class", "measurement_value warning")
        else:
            # 恢复到正常的measurement_value类
            self.el_measured_voltage_label.setProperty("class", self._el_voltage_label_original_class)
            self.el_measured_current_label.setProperty("class", self._el_current_label_original_class)
            self.el_measured_power_label.setProperty("class", self._el_power_label_original_class)

        # 强制刷新样式以应用警告/正常颜色（在潜在闪烁之前）
        self.el_measured_voltage_label.style().polish(self.el_measured_voltage_label)
        self.el_measured_current_label.style().polish(self.el_measured_current_label)
        self.el_measured_power_label.style().polish(self.el_measured_power_label)

        # 如果值发生变化，则应用闪烁效果（闪烁会暂时覆盖，然后恢复）
        if voltage_changed:
            self.flash_measurement_label(self.el_measured_voltage_label, self.el_measured_voltage_label.property("class"))
        if current_changed:
            self.flash_measurement_label(self.el_measured_current_label, self.el_measured_current_label.property("class"))
        if power_changed:
            self.flash_measurement_label(self.el_measured_power_label, self.el_measured_power_label.property("class"))

    def refresh_el_status(self):
        """刷新电子负载的所有状态和测量值 (通过信号发送到工作线程)。"""
        if not self.el_worker or not self.el_worker._is_connected:
            QMessageBox.warning(self, "操作失败", "电子负载设备未连接，无法刷新。")
            return
        self.el_refresh_request.emit()

    @pyqtSlot(str, str, str)
    def show_error_message(self, device_name, title, message):
        """显示来自工作线程的错误消息。"""
        QMessageBox.critical(self, f"{device_name} - {title}", message)

    def closeEvent(self, event):
        """主窗口关闭事件处理函数，确保设备安全断开。"""
        self.log_message("<font color='#666666'>正在关闭应用程序...</font>") # 普通日志颜色
        
        # 尝试安全断开所有连接，无二次确认
        # 在关闭前停止任何激活的脉动动画
        if hasattr(self.ps_output_status_indicator, '_is_pulsing'):
            self.ps_output_status_indicator._is_pulsing = False
        if hasattr(self.el_input_status_indicator, '_is_pulsing'):
            self.el_input_status_indicator._is_pulsing = False

        if self.ps_worker and self.ps_worker._is_connected:
            self.log_message("<font color='#666666'>尝试安全关闭电源输出并断开连接...</font>") # 普通日志颜色
            self.ps_command_request.emit("OUTP", "OFF", False, "app_exit_cleanup")
            QThread.msleep(100) # 短暂等待
            self.ps_disconnect_request.emit() 
            if self.ps_thread and self.ps_thread.isRunning():
                self.ps_thread.quit()
                self.ps_thread.wait(2000) # 给予更多时间确保线程退出

        if self.el_worker and self.el_worker._is_connected:
            self.log_message("<font color='#666666'>尝试安全关闭电子负载输入并断开连接...</font>") # 普通日志颜色
            self.el_command_request.emit("INP", "OFF", False, "app_exit_cleanup")
            QThread.msleep(100)
            self.el_disconnect_request.emit()
            if self.el_thread and self.el_thread.isRunning():
                self.el_thread.quit()
                self.el_thread.wait(2000)

        if self.rm:
            try:
                self.rm.close()
                self.log_message("<font color='#28A745'>成功:</font> VISA Resource Manager 已关闭。") # 成功色
            except Exception as e:
                self.log_message(f"<font color='#DC3545'>错误:</font> 关闭 VISA Resource Manager 时出错: {e}") # 错误色
        
        super().closeEvent(event)
        self.log_message("<font color='#666666'>应用程序已退出。</font>") # 普通日志颜色

if __name__ == '__main__':
    app = QApplication(sys.argv)
 
    app.setAttribute(Qt.AA_EnableHighDpiScaling, True) 
    app.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    
    font = QFont("Microsoft YaHei UI", 10)
    app.setFont(font)

    main_win = UnifiedControllerGUI()
    main_win.show()
    sys.exit(app.exec_())