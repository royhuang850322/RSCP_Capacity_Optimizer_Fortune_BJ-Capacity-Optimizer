"""PySide6 desktop UI for the Fortune BJ scheduling optimizer."""
from __future__ import annotations

import json
import re
import sys
import traceback
from datetime import datetime
from pathlib import Path

from app.fortune_bj import (
    DATA_DIR,
    DEPLOY_ROOT,
    PROJECT_ROOT,
    REPORT_DIR,
    DEMAND_TEMPLATE_NAME,
    OPS_TEMPLATE_NAME,
    WC_TEMPLATE_NAME,
    OPTIONAL_OPS_TEMPLATE_NAME,
    CALENDAR_TEMPLATE_NAME,
    FORECAST_TEMPLATE_NAME,
    FortuneBjConfig,
    run_fortune_bj_schedule,
)
from app.license_validator import LicenseInfo, LicenseValidationError, validate_license_with_fallback
from app.machine_fingerprint import build_machine_identity_payload, sanitize_machine_label
from app.version import APP_VERSION

try:
    from PySide6.QtCore import QDate, QSettings, QThread, Signal, Qt, QUrl
    from PySide6.QtGui import QDesktopServices, QPalette
    from PySide6.QtWidgets import (
        QApplication,
        QCheckBox,
        QComboBox,
        QDateEdit,
        QFileDialog,
        QFormLayout,
        QFrame,
        QGridLayout,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QListWidget,
        QListWidgetItem,
        QMainWindow,
        QMessageBox,
        QProgressBar,
        QPushButton,
        QScrollArea,
        QSpinBox,
        QStackedWidget,
        QTextEdit,
        QVBoxLayout,
        QWidget,
    )
except ModuleNotFoundError:  # pragma: no cover
    PYSIDE6_AVAILABLE = False
    QSettings = object  # type: ignore[assignment]
    QThread = object  # type: ignore[assignment]
    QMainWindow = object  # type: ignore[assignment]

    def Signal(*_args, **_kwargs):  # type: ignore[no-redef]
        return None
else:
    PYSIDE6_AVAILABLE = True


if PYSIDE6_AVAILABLE:

    class NoWheelComboBox(QComboBox):
        def wheelEvent(self, event) -> None:  # pragma: no cover - UI event behavior
            event.ignore()


    class NoWheelDateEdit(QDateEdit):
        def wheelEvent(self, event) -> None:  # pragma: no cover - UI event behavior
            event.ignore()


    class NoWheelSpinBox(QSpinBox):
        def wheelEvent(self, event) -> None:  # pragma: no cover - UI event behavior
            event.ignore()

else:  # pragma: no cover
    NoWheelComboBox = object  # type: ignore[assignment]
    NoWheelDateEdit = object  # type: ignore[assignment]
    NoWheelSpinBox = object  # type: ignore[assignment]


APP_TITLE = f"Fortune BJ 产能分析工具 v{APP_VERSION}"
SETTINGS_ORGANIZATION = "RSCP"
SETTINGS_APPLICATION = "FortuneBJCapacityOptimizer"
LICENSE_ACTIVE_DIR = DEPLOY_ROOT / "licenses" / "active"
LICENSE_REQUESTS_DIR = DEPLOY_ROOT / "licenses" / "requests"
LOG_DIR = DEPLOY_ROOT / "logs"


def _display_path(path: str | Path) -> str:
    return str(path).replace("/", "\\")


LIGHT_QSS = """
QWidget {
    background: #f4f6fa;
    color: #172033;
    font-family: "Microsoft YaHei", "Segoe UI";
    font-size: 13px;
}
QFrame#Card {
    background: #ffffff;
    border: 1px solid #d9e2ef;
    border-radius: 10px;
}
QLabel#Title {
    font-size: 28px;
    font-weight: 700;
}
QLabel#Subtitle {
    color: #65748b;
    font-size: 14px;
}
QLabel#SectionTitle {
    background: #f0f3f8;
    font-size: 16px;
    font-weight: 700;
    padding: 6px 4px;
}
QLabel#StatusBadge {
    background: #e8eff9;
    border: 1px solid #cfd9e8;
    border-radius: 8px;
    padding: 7px 10px;
    font-weight: 600;
}
QListWidget {
    background: #f7f9fc;
    border: 1px solid #e2e8f2;
    border-radius: 8px;
    padding: 6px;
}
QListWidget::item {
    padding: 10px 8px;
    border-radius: 6px;
}
QListWidget::item:selected {
    background: #e6ebf3;
    border-left: 3px solid #1f7a8c;
}
QLineEdit, QComboBox, QTextEdit {
    background: #ffffff;
    border: 1px solid #cbd7e8;
    border-radius: 8px;
    padding: 7px 9px;
}
QPushButton {
    background: #e8eef7;
    border: 1px solid #c9d5e8;
    border-radius: 8px;
    padding: 9px 12px;
}
QPushButton:hover {
    background: #dde8f6;
}
QPushButton#Primary {
    background: #1f77b4;
    color: white;
    border: 1px solid #176092;
    font-weight: 700;
    padding: 12px 14px;
}
QPushButton#Primary:hover {
    background: #19689f;
}
"""

DARK_QSS = """
QWidget {
    background: #151b23;
    color: #e8eef7;
    font-family: "Microsoft YaHei", "Segoe UI";
    font-size: 13px;
}
QFrame#Card {
    background: #202a36;
    border: 1px solid #334357;
    border-radius: 10px;
}
QLabel#Title {
    font-size: 28px;
    font-weight: 700;
}
QLabel#Subtitle {
    color: #9ba9bf;
    font-size: 14px;
}
QLabel#SectionTitle {
    background: #263241;
    font-size: 16px;
    font-weight: 700;
    padding: 6px 4px;
}
QLabel#StatusBadge {
    background: #2d3d52;
    border: 1px solid #455f7d;
    border-radius: 8px;
    padding: 7px 10px;
    font-weight: 600;
}
QListWidget {
    background: #1c2531;
    border: 1px solid #334357;
    border-radius: 8px;
    padding: 6px;
}
QListWidget::item {
    padding: 10px 8px;
    border-radius: 6px;
}
QListWidget::item:selected {
    background: #2a3645;
    border-left: 3px solid #38a3a5;
}
QLineEdit, QComboBox, QTextEdit {
    background: #1c2531;
    color: #eef4fc;
    border: 1px solid #3a4d63;
    border-radius: 8px;
    padding: 7px 9px;
}
QPushButton {
    background: #2a3645;
    color: #e8eef7;
    border: 1px solid #3a4d63;
    border-radius: 8px;
    padding: 9px 12px;
}
QPushButton:hover {
    background: #334357;
}
QPushButton#Primary {
    background: #1d79c4;
    color: white;
    border: 1px solid #1765a4;
    font-weight: 700;
    padding: 12px 14px;
}
QPushButton#Primary:hover {
    background: #1a6db0;
}
"""


def ensure_license_dirs() -> None:
    LICENSE_ACTIVE_DIR.mkdir(parents=True, exist_ok=True)
    LICENSE_REQUESTS_DIR.mkdir(parents=True, exist_ok=True)
    LOG_DIR.mkdir(parents=True, exist_ok=True)


def inspect_license() -> tuple[LicenseInfo | None, str]:
    try:
        info = validate_license_with_fallback(
            primary_root=str(DEPLOY_ROOT),
            fallback_roots=[str(PROJECT_ROOT)],
        )
        return info, ""
    except LicenseValidationError as exc:
        return None, str(exc)
    except Exception:
        return None, traceback.format_exc()


def generate_machine_fingerprint_request() -> Path:
    ensure_license_dirs()
    payload = build_machine_identity_payload()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"machine_fingerprint_{sanitize_machine_label(payload['machine_label'])}_{timestamp}.json"
    output_path = LICENSE_REQUESTS_DIR / filename
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return output_path


class SchedulerWorker(QThread):
    finished_ok = Signal(str)
    failed = Signal(str)
    progress = Signal(str)

    def __init__(self, config: FortuneBjConfig) -> None:
        super().__init__()
        self.config = config

    def run(self) -> None:  # pragma: no cover - UI thread integration
        try:
            result = run_fortune_bj_schedule(self.config, progress=self.progress.emit)
            self.finished_ok.emit(str(result.report_path))
        except Exception:
            self.failed.emit(traceback.format_exc())


class CardFrame(QFrame):
    def __init__(self, title: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("Card")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)
        self.title = QLabel(title)
        self.title.setObjectName("SectionTitle")
        layout.addWidget(self.title)
        self.body = QWidget(self)
        self.body_layout = QVBoxLayout(self.body)
        self.body_layout.setContentsMargins(0, 0, 0, 0)
        self.body_layout.setSpacing(8)
        layout.addWidget(self.body)

    def set_title(self, title: str) -> None:
        self.title.setText(title)


class MainWindow(QMainWindow):
    def __init__(self, settings=None) -> None:
        super().__init__()
        ensure_license_dirs()
        self.settings = settings if settings is not None else QSettings(
            SETTINGS_ORGANIZATION,
            SETTINGS_APPLICATION,
        )
        self.setWindowTitle(APP_TITLE)
        self.resize(1380, 840)
        self.setMinimumSize(1120, 720)
        self.worker: SchedulerWorker | None = None
        self.last_report: Path | None = None
        self.current_run_log_path: Path | None = None
        self.optimization_start_month_manual = False
        self._build_ui()
        self._wire_events()
        self._restore_settings()
        self.optimization_start_month_manual = False
        self._set_optimization_start_period_to_current()
        self._apply_theme(self.theme_combo.currentText())
        self.refresh_license_status()
        self._log("启动器已就绪。")

    def _build_ui(self) -> None:
        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(14, 14, 14, 14)
        root_layout.setSpacing(10)
        root_layout.addWidget(self._build_header())

        content = QHBoxLayout()
        content.setSpacing(10)
        content.addWidget(self._build_sidebar(), 0)
        self.stack = QStackedWidget()
        self.stack.addWidget(self._scroll_page(self._build_home_page()))
        self.stack.addWidget(self._scroll_page(self._build_config_page()))
        self.stack.addWidget(self._scroll_page(self._build_license_page()))
        content.addWidget(self.stack, 1)
        root_layout.addLayout(content, 1)
        self.setCentralWidget(root)

    def _build_header(self) -> QWidget:
        card = CardFrame("页眉")
        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        left = QVBoxLayout()
        self.header_title = QLabel(APP_TITLE)
        self.header_title.setObjectName("Title")
        self.header_subtitle = QLabel("用于 Fortune BJ 工单产能分析、资源映射校验、OR-Tools 有限产能模拟和 Excel 分析报告。")
        self.header_subtitle.setObjectName("Subtitle")
        left.addWidget(self.header_title)
        left.addWidget(self.header_subtitle)
        right = QHBoxLayout()
        right.setAlignment(Qt.AlignRight | Qt.AlignTop)
        self.theme_combo = NoWheelComboBox()
        self.theme_combo.addItems(["跟随系统", "浅色", "深色"])
        self.language_combo = NoWheelComboBox()
        self.language_combo.addItems(["中文"])
        self.help_button = QPushButton("帮助")
        right.addWidget(QLabel("主题"))
        right.addWidget(self.theme_combo)
        right.addWidget(QLabel("语言"))
        right.addWidget(self.language_combo)
        right.addWidget(self.help_button)
        row.addLayout(left, 1)
        row.addLayout(right)
        card.body_layout.addLayout(row)
        return card

    def _build_sidebar(self) -> QWidget:
        card = CardFrame("导航与设置")
        card.setFixedWidth(370)
        self.nav = QListWidget()
        for label in ("首页", "配置", "许可与诊断"):
            QListWidgetItem(label, self.nav)
        self.nav.setCurrentRow(0)
        self.nav.setFixedHeight(150)
        card.body_layout.addWidget(self.nav)

        settings_scroll = QScrollArea()
        settings_scroll.setWidgetResizable(True)
        settings_scroll.setFrameShape(QFrame.NoFrame)
        holder = QWidget()
        layout = QVBoxLayout(holder)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        layout.addWidget(self._build_sidebar_option_card())
        layout.addWidget(self._build_sidebar_runtime_card())
        layout.addStretch(1)
        settings_scroll.setWidget(holder)
        card.body_layout.addWidget(settings_scroll, 1)

        self.sidebar_status = QLabel("授权状态：检查中")
        self.sidebar_status.setObjectName("StatusBadge")
        card.body_layout.addWidget(self.sidebar_status)
        return card

    def _build_sidebar_option_card(self) -> QWidget:
        card = CardFrame("产能分析设置")
        form = QFormLayout()
        self.schedule_mode = NoWheelComboBox()
        self.schedule_mode.addItems(["ModeA：无限产能分析", "ModeB：100%产能优化建议"])
        self.mode_b_optimization_granularity = NoWheelComboBox()
        self.mode_b_optimization_granularity.addItems(["周", "月"])
        self.operation_flow_mode = NoWheelComboBox()
        self.operation_flow_mode.addItems(["整批流转", "半批流转", "单件流转", "交期强制"])
        self.mode_b_optimization_start_month = NoWheelDateEdit()
        self.mode_b_optimization_start_month.setCalendarPopup(True)
        self.mode_b_optimization_start_month.setMinimumWidth(120)
        self._set_optimization_start_period_to_current()
        self.mode_b_max_window_tasks = self._number_spinbox(2000, minimum=1, maximum=200000, suffix=" 条")
        self.mode_b_solver_max_seconds = self._number_spinbox(60, minimum=1, maximum=3600, suffix=" 秒")
        self.hot_surface_mode = NoWheelComboBox()
        self.hot_surface_mode.addItems(["同机加逻辑", "热处/表处专用逻辑"])
        self.objective_profile = NoWheelComboBox()
        self.objective_profile.addItems(["默认：产能缺口最小", "紧急类型优先（可选）", "产能均衡（待开发）", "减少资源空闲/换线（待开发）"])
        self.enable_forecast = QCheckBox("需求预测导入")
        self.enable_urgent = QCheckBox("启用紧急类型优先级")
        self.enable_urgent.setChecked(True)
        form.addRow("计算模式", self.schedule_mode)
        form.addRow("优化粒度", self.mode_b_optimization_granularity)
        form.addRow("工序流转逻辑", self.operation_flow_mode)
        form.addRow("优化开始日期", self.mode_b_optimization_start_month)
        form.addRow("", self.enable_forecast)
        form.addRow("周期参考工序数", self.mode_b_max_window_tasks)
        form.addRow("OR-Tools求解上限", self.mode_b_solver_max_seconds)
        form.addRow("热处/表处", self.hot_surface_mode)
        form.addRow("优化目标", self.objective_profile)
        form.addRow("", self.enable_urgent)
        card.body_layout.addLayout(form)
        return card

    def _build_sidebar_runtime_card(self) -> QWidget:
        card = CardFrame("运行设置")
        form = QFormLayout()
        self.out_dir = self._path_edit(REPORT_DIR)
        form.addRow("报告输出", self._browse_row(self.out_dir, file_mode=False))
        card.body_layout.addLayout(form)
        return card

    def _build_home_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        actions = CardFrame("操作")
        grid = QGridLayout()
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(10)
        self.run_btn = QPushButton("运行产能分析")
        self.run_btn.setObjectName("Primary")
        self.open_report_btn = QPushButton("打开报告目录")
        self.open_log_btn = QPushButton("打开日志目录")
        self.open_license_btn = QPushButton("打开授权目录")
        self.open_requests_btn = QPushButton("打开机器指纹目录")
        grid.addWidget(self.run_btn, 0, 0, 1, 2)
        grid.addWidget(self.open_report_btn, 1, 0)
        grid.addWidget(self.open_log_btn, 1, 1)
        grid.addWidget(self.open_license_btn, 2, 0)
        grid.addWidget(self.open_requests_btn, 2, 1)
        actions.body_layout.addLayout(grid)
        layout.addWidget(actions)

        status = CardFrame("运行状态")
        self.progress_phase = QLabel("当前阶段：待运行")
        self.progress_phase.setObjectName("StatusBadge")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 1000)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_metrics = QLabel("进度：0/0 | 已用 0秒 | 预计剩余 计算中")
        self.progress_metrics.setObjectName("Subtitle")
        self.window_task_count = QLabel("优化周期实际工序数：待运行")
        self.window_task_count.setObjectName("Subtitle")
        status.body_layout.addWidget(self.progress_phase)
        status.body_layout.addWidget(self.progress_bar)
        status.body_layout.addWidget(self.progress_metrics)
        status.body_layout.addWidget(self.window_task_count)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(260)
        status.body_layout.addWidget(self.log)
        layout.addWidget(status, 1)
        return page

    def _build_config_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        files = CardFrame("数据文件")
        form = QFormLayout()
        self.ops_path = self._path_edit(DATA_DIR / OPS_TEMPLATE_NAME)
        self.demand_path = self._path_edit(DATA_DIR / DEMAND_TEMPLATE_NAME)
        self.wc_path = self._path_edit(DATA_DIR / WC_TEMPLATE_NAME)
        self.optional_ops_path = self._path_edit(DATA_DIR / OPTIONAL_OPS_TEMPLATE_NAME)
        self.calendar_path = self._path_edit(DATA_DIR / CALENDAR_TEMPLATE_NAME)
        self.forecast_path = self._path_edit(DATA_DIR / FORECAST_TEMPLATE_NAME)
        form.addRow("生产订单工序表", self._browse_row(self.ops_path, file_mode=True))
        form.addRow("订单交期数量表", self._browse_row(self.demand_path, file_mode=True))
        form.addRow("工作中心表", self._browse_row(self.wc_path, file_mode=True))
        form.addRow("可选工序表", self._browse_row(self.optional_ops_path, file_mode=True))
        form.addRow("工作日历表", self._browse_row(self.calendar_path, file_mode=True))
        form.addRow("物料需求预测表", self._browse_row(self.forecast_path, file_mode=True))
        files.body_layout.addLayout(form)
        layout.addWidget(files)

        paths = CardFrame("目录")
        path_form = QFormLayout()
        self.deploy_root_label = QLineEdit(_display_path(DEPLOY_ROOT))
        self.deploy_root_label.setReadOnly(True)
        self.data_dir_label = QLineEdit(_display_path(DATA_DIR))
        self.data_dir_label.setReadOnly(True)
        path_form.addRow("应用目录", self.deploy_root_label)
        path_form.addRow("数据导入目录", self.data_dir_label)
        paths.body_layout.addLayout(path_form)
        layout.addWidget(paths)
        layout.addStretch(1)
        return page

    def _build_license_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        license_card = CardFrame("许可状态")
        form = QFormLayout()
        self.license_status = QLineEdit()
        self.license_status.setReadOnly(True)
        self.license_id = QLineEdit()
        self.license_id.setReadOnly(True)
        self.license_customer = QLineEdit()
        self.license_customer.setReadOnly(True)
        self.license_expiry = QLineEdit()
        self.license_expiry.setReadOnly(True)
        self.license_path = QLineEdit()
        self.license_path.setReadOnly(True)
        form.addRow("状态", self.license_status)
        form.addRow("License ID", self.license_id)
        form.addRow("授权客户", self.license_customer)
        form.addRow("到期日", self.license_expiry)
        form.addRow("授权文件", self.license_path)
        license_card.body_layout.addLayout(form)
        row = QHBoxLayout()
        self.refresh_license_btn = QPushButton("刷新许可状态")
        self.generate_fingerprint_btn = QPushButton("生成机器指纹请求")
        row.addWidget(self.refresh_license_btn)
        row.addWidget(self.generate_fingerprint_btn)
        row.addStretch(1)
        license_card.body_layout.addLayout(row)
        layout.addWidget(license_card)

        diag = CardFrame("诊断信息")
        self.license_message = QTextEdit()
        self.license_message.setReadOnly(True)
        self.license_message.setMinimumHeight(220)
        diag.body_layout.addWidget(self.license_message)
        layout.addWidget(diag, 1)
        return page

    def _scroll_page(self, page: QWidget) -> QScrollArea:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setWidget(page)
        return scroll

    def _wire_events(self) -> None:
        self.nav.currentRowChanged.connect(self.stack.setCurrentIndex)
        self.theme_combo.currentTextChanged.connect(self._apply_theme)
        self.help_button.clicked.connect(self._show_help)
        self.run_btn.clicked.connect(self.run_schedule)
        self.open_report_btn.clicked.connect(lambda: self._open_folder(Path(self.out_dir.text())))
        self.open_log_btn.clicked.connect(lambda: self._open_folder(LOG_DIR))
        self.open_license_btn.clicked.connect(lambda: self._open_folder(LICENSE_ACTIVE_DIR))
        self.open_requests_btn.clicked.connect(lambda: self._open_folder(LICENSE_REQUESTS_DIR))
        self.refresh_license_btn.clicked.connect(self.refresh_license_status)
        self.generate_fingerprint_btn.clicked.connect(self.generate_fingerprint)
        self.mode_b_optimization_start_month.dateChanged.connect(self._mark_optimization_start_month_manual)
        self.mode_b_optimization_granularity.currentTextChanged.connect(self._handle_optimization_granularity_changed)

    @staticmethod
    def _restore_combo_text(combo: QComboBox, value: object) -> None:
        text = str(value or "")
        index = combo.findText(text)
        if index >= 0:
            combo.setCurrentIndex(index)

    @staticmethod
    def _setting_bool(value: object, default: bool = False) -> bool:
        if value is None:
            return default
        if isinstance(value, bool):
            return value
        return str(value).strip().lower() in {"1", "true", "yes", "on"}

    def _restore_settings(self) -> None:
        settings = self.settings
        geometry = settings.value("window/geometry")
        if geometry is not None:
            self.restoreGeometry(geometry)

        self._restore_combo_text(self.theme_combo, settings.value("ui/theme", "跟随系统"))
        self._restore_combo_text(self.schedule_mode, settings.value("analysis/schedule_mode"))
        self._restore_combo_text(
            self.mode_b_optimization_granularity,
            settings.value("analysis/optimization_granularity"),
        )
        self._restore_combo_text(self.operation_flow_mode, settings.value("analysis/operation_flow_mode"))
        self._restore_combo_text(self.hot_surface_mode, settings.value("analysis/hot_surface_mode"))
        self._restore_combo_text(self.objective_profile, settings.value("analysis/objective_profile"))

        self.enable_forecast.setChecked(
            self._setting_bool(settings.value("analysis/enable_forecast"), self.enable_forecast.isChecked())
        )
        self.enable_urgent.setChecked(
            self._setting_bool(settings.value("analysis/enable_urgent"), self.enable_urgent.isChecked())
        )
        self.mode_b_max_window_tasks.setValue(
            int(settings.value("analysis/max_window_tasks", self.mode_b_max_window_tasks.value()))
        )
        self.mode_b_solver_max_seconds.setValue(
            int(settings.value("analysis/solver_max_seconds", self.mode_b_solver_max_seconds.value()))
        )

        path_controls = {
            "paths/operations": self.ops_path,
            "paths/demand": self.demand_path,
            "paths/workcenter": self.wc_path,
            "paths/optional_operations": self.optional_ops_path,
            "paths/calendar": self.calendar_path,
            "paths/forecast": self.forecast_path,
            "paths/output": self.out_dir,
        }
        for key, edit in path_controls.items():
            saved = settings.value(key)
            if saved not in (None, ""):
                edit.setText(_display_path(str(saved)))

        saved_page = int(settings.value("ui/current_page", 0))
        self.nav.setCurrentRow(max(0, min(saved_page, self.nav.count() - 1)))

    def _save_settings(self) -> None:
        self._normalize_path_fields()
        settings = self.settings
        settings.setValue("window/geometry", self.saveGeometry())
        settings.setValue("ui/theme", self.theme_combo.currentText())
        settings.setValue("ui/current_page", self.nav.currentRow())
        settings.setValue("analysis/schedule_mode", self.schedule_mode.currentText())
        settings.setValue("analysis/optimization_granularity", self.mode_b_optimization_granularity.currentText())
        settings.setValue("analysis/operation_flow_mode", self.operation_flow_mode.currentText())
        settings.setValue("analysis/max_window_tasks", self.mode_b_max_window_tasks.value())
        settings.setValue("analysis/solver_max_seconds", self.mode_b_solver_max_seconds.value())
        settings.setValue("analysis/hot_surface_mode", self.hot_surface_mode.currentText())
        settings.setValue("analysis/objective_profile", self.objective_profile.currentText())
        settings.setValue("analysis/enable_forecast", self.enable_forecast.isChecked())
        settings.setValue("analysis/enable_urgent", self.enable_urgent.isChecked())
        settings.setValue("paths/operations", self.ops_path.text())
        settings.setValue("paths/demand", self.demand_path.text())
        settings.setValue("paths/workcenter", self.wc_path.text())
        settings.setValue("paths/optional_operations", self.optional_ops_path.text())
        settings.setValue("paths/calendar", self.calendar_path.text())
        settings.setValue("paths/forecast", self.forecast_path.text())
        settings.setValue("paths/output", self.out_dir.text())
        settings.sync()

    def closeEvent(self, event) -> None:  # pragma: no cover - UI lifecycle
        self._save_settings()
        super().closeEvent(event)

    def _path_edit(self, path: Path) -> QLineEdit:
        edit = QLineEdit(_display_path(path))
        edit.setMinimumWidth(620)
        return edit

    def _set_optimization_start_period_to_current(self) -> None:
        today = QDate.currentDate()
        start_date = QDate(today.year(), today.month(), today.day())
        display_format = "yyyy-MM-dd"
        if hasattr(self, "mode_b_optimization_start_month"):
            self.mode_b_optimization_start_month.blockSignals(True)
            self.mode_b_optimization_start_month.setDisplayFormat(display_format)
            self.mode_b_optimization_start_month.setDate(start_date)
            self.mode_b_optimization_start_month.blockSignals(False)

    def _mark_optimization_start_month_manual(self, *_args) -> None:
        self.optimization_start_month_manual = True

    def _handle_optimization_granularity_changed(self, *_args) -> None:
        if not self.optimization_start_month_manual:
            self._set_optimization_start_period_to_current()
            return
        self.mode_b_optimization_start_month.setDisplayFormat("yyyy-MM-dd")

    def _selected_optimization_start_month(self) -> datetime:
        if not self.optimization_start_month_manual:
            self._set_optimization_start_period_to_current()
        value = self.mode_b_optimization_start_month.date()
        return datetime(value.year(), value.month(), value.day())

    def _days_spinbox(self, value: int, *, minimum: int) -> QSpinBox:
        return self._number_spinbox(value, minimum=minimum, maximum=3650, suffix=" 天")

    def _number_spinbox(self, value: int, *, minimum: int, maximum: int, suffix: str) -> QSpinBox:
        spin = NoWheelSpinBox()
        spin.setRange(minimum, maximum)
        spin.setValue(value)
        spin.setSuffix(suffix)
        spin.setMinimumWidth(120)
        return spin

    def _browse_row(self, edit: QLineEdit, *, file_mode: bool) -> QWidget:
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(edit, 1)
        btn = QPushButton("浏览")
        btn.clicked.connect(lambda: self._browse(edit, file_mode=file_mode))
        layout.addWidget(btn)
        return row

    def _browse(self, edit: QLineEdit, *, file_mode: bool) -> None:
        if file_mode:
            selected, _ = QFileDialog.getOpenFileName(self, "选择数据文件", edit.text(), "数据文件 (*.csv *.xlsx *.xls)")
        else:
            selected = QFileDialog.getExistingDirectory(self, "选择文件夹", edit.text())
        if selected:
            edit.setText(_display_path(selected))

    def _normalize_path_fields(self) -> None:
        for edit in (
            getattr(self, "ops_path", None),
            getattr(self, "demand_path", None),
            getattr(self, "wc_path", None),
            getattr(self, "optional_ops_path", None),
            getattr(self, "calendar_path", None),
            getattr(self, "forecast_path", None),
            getattr(self, "out_dir", None),
        ):
            if edit is not None:
                edit.setText(_display_path(edit.text()))

    def refresh_license_status(self) -> None:
        info, error = inspect_license()
        if info is None:
            self.sidebar_status.setText("授权状态：未通过")
            self.license_status.setText("无效 / 未配置")
            self.license_id.setText("")
            self.license_customer.setText("")
            self.license_expiry.setText("")
            self.license_path.setText(_display_path(LICENSE_ACTIVE_DIR / "license.json"))
            self.license_message.setPlainText(error)
            return
        self.sidebar_status.setText(f"授权状态：有效，{info.expiry_date} 到期")
        self.license_status.setText(info.status)
        self.license_id.setText(info.license_id)
        self.license_customer.setText(info.customer_name)
        self.license_expiry.setText(info.expiry_date)
        self.license_path.setText(_display_path(info.license_path))
        self.license_message.setPlainText(
            "\n".join([
                f"License ID: {info.license_id}",
                f"Type: {info.license_type}",
                f"Customer: {info.customer_name}",
                f"Binding: {info.binding_mode}",
                f"Machine: {info.machine_label}",
                f"Features: {json.dumps(info.features, ensure_ascii=False)}",
                f"File: {_display_path(info.license_path)}",
            ])
        )

    def generate_fingerprint(self) -> None:
        try:
            path = generate_machine_fingerprint_request()
        except Exception as exc:
            self._show_error("生成机器指纹失败", str(exc))
            return
        self._log(f"机器指纹请求已生成：{_display_path(path)}")
        QMessageBox.information(self, "完成", f"机器指纹请求已生成：\n{_display_path(path)}")

    def run_schedule(self) -> None:
        self._begin_run_log("schedule")
        self._log("Run requested: schedule")
        self._normalize_path_fields()
        self.refresh_license_status()
        info, error = inspect_license()
        if info is None:
            self._show_error("许可校验失败", error)
            self.nav.setCurrentRow(2)
            self._finish_run_log("FAILED", error)
            return
        forecast_text = self.forecast_path.text().strip()
        if self.enable_forecast.isChecked() and not forecast_text:
            message = "已勾选需求预测导入，但物料需求预测表路径为空。请先选择需求预测_产能分析输入模板.xlsx。"
            self._show_error("需求预测输入缺失", message)
            self._finish_run_log("FAILED", message)
            return
        config = FortuneBjConfig(
            operations_path=Path(self.ops_path.text()),
            demand_path=Path(self.demand_path.text()),
            workcenter_path=Path(self.wc_path.text()),
            optional_operations_path=Path(self.optional_ops_path.text()),
            calendar_path=Path(self.calendar_path.text()),
            forecast_path=Path(forecast_text) if forecast_text else None,
            output_dir=Path(self.out_dir.text()),
            schedule_mode="ModeB" if self.schedule_mode.currentText().startswith("ModeB") else "ModeA",
            mode_b_optimization_granularity=self.mode_b_optimization_granularity.currentText(),
            mode_b_optimization_start_month=self._selected_optimization_start_month(),
            mode_b_max_window_tasks=self.mode_b_max_window_tasks.value(),
            mode_b_solver_max_seconds=float(self.mode_b_solver_max_seconds.value()),
            enable_urgent=self.enable_urgent.isChecked(),
            enable_forecast=self.enable_forecast.isChecked(),
            operation_flow_mode=self.operation_flow_mode.currentText(),
            hot_surface_mode=self.hot_surface_mode.currentText(),
            objective_profile=self.objective_profile.currentText(),
        )
        self.run_btn.setEnabled(False)
        self._reset_progress("准备运行")
        if config.schedule_mode == "ModeB":
            self.window_task_count.setText(
                f"优化周期实际工序数：等待周期开始 | 参考 {config.mode_b_max_window_tasks:,} 条"
            )
            self._log(
                "ModeB参数："
                f"优化粒度 {config.mode_b_optimization_granularity}，"
                f"工序流转 {config.operation_flow_mode}，"
                f"优化开始日期 {config.mode_b_optimization_start_month.strftime('%Y-%m-%d') if config.mode_b_optimization_start_month else '未指定'}，"
                f"需求预测 {'启用' if config.enable_forecast else '未启用'}，"
                f"按{config.mode_b_optimization_granularity}覆盖全部数据，"
                f"参考工序数 {config.mode_b_max_window_tasks:,} 条，"
                f"OR-Tools求解上限 {config.mode_b_solver_max_seconds:g} 秒；"
                "参考工序数只写入日志和报告，不作为硬性跳过阈值。"
            )
        else:
            self.window_task_count.setText(
                f"优化周期实际工序数：ModeA 按{config.mode_b_optimization_granularity}度倒排平移基线"
            )
            self._log(
                "ModeA参数："
                f"优化粒度 {config.mode_b_optimization_granularity}，"
                f"工序流转 {config.operation_flow_mode}，"
                f"优化开始日期 {config.mode_b_optimization_start_month.strftime('%Y-%m-%d') if config.mode_b_optimization_start_month else '未指定'}，"
                f"需求预测 {'启用' if config.enable_forecast else '未启用'}，"
                "按交期倒排；逾期订单或倒排后早于优化开始日期的订单整单平移。"
            )
        self._log("开始运行产能分析...")
        self.worker = SchedulerWorker(config)
        self.worker.progress.connect(self._handle_progress)
        self.worker.finished_ok.connect(self._run_finished)
        self.worker.failed.connect(self._run_failed)
        self.worker.start()

    def _run_finished(self, report_path: str) -> None:
        self.run_btn.setEnabled(True)
        self.last_report = Path(report_path)
        self.progress_phase.setText("当前阶段：完成")
        self.progress_bar.setValue(1000)
        self.progress_metrics.setText("进度：100% | 已完成")
        self._log(f"运行完成，报告已生成：{_display_path(report_path)}")
        self._finish_run_log("SUCCESS", f"Report: {_display_path(report_path)}")
        self._show_run_complete_dialog(Path(report_path))

    def _run_failed(self, error: str) -> None:
        self.run_btn.setEnabled(True)
        self.progress_phase.setText("当前阶段：失败")
        self._log(error)
        self._finish_run_log("FAILED", error)
        report_path = self._extract_report_path_from_error(error)
        if report_path is None:
            QMessageBox.critical(self, "运行失败", error[-3000:])
            return
        dialog = QMessageBox(self)
        dialog.setIcon(QMessageBox.Critical)
        dialog.setWindowTitle("分析前数据校验未通过")
        dialog.setText("正式产能分析已停止，请先修复数据源。")
        dialog.setInformativeText(f"{error[-2200:]}\n\n校验报告：{report_path}")
        open_button = dialog.addButton("打开校验报告", QMessageBox.AcceptRole)
        dialog.addButton("关闭", QMessageBox.RejectRole)
        dialog.setDefaultButton(open_button)
        dialog.exec()
        if dialog.clickedButton() == open_button:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(report_path)))

    def _extract_report_path_from_error(self, error: str) -> Path | None:
        match = re.search(r"校验报告[:：]\s*(?P<path>[^\r\n]+?\.xlsx)", error)
        if match is None:
            return None
        raw_path = match.group("path").strip().strip("'\"")
        path = Path(raw_path)
        candidates = [path]
        if not path.is_absolute():
            candidates.extend([Path.cwd() / path, DEPLOY_ROOT / path])
        for candidate in candidates:
            if candidate.exists():
                return candidate.resolve()
        return path

    def _show_run_complete_dialog(self, report_path: Path) -> None:
        dialog = QMessageBox(self)
        dialog.setIcon(QMessageBox.Information)
        dialog.setWindowTitle("运行完成")
        dialog.setText("报告已生成。")
        dialog.setInformativeText(_display_path(report_path))
        open_button = dialog.addButton("打开报告", QMessageBox.AcceptRole)
        dialog.addButton("关闭", QMessageBox.RejectRole)
        dialog.setDefaultButton(open_button)
        dialog.exec()
        if dialog.clickedButton() == open_button:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(report_path)))

    def _open_folder(self, path: Path) -> None:
        path.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))

    def _show_help(self) -> None:
        QMessageBox.information(
            self,
            "帮助",
            "1. 将 license.json 放入 licenses\\active。\n"
            "2. 如需机绑授权，先在“许可与诊断”生成机器指纹请求。\n"
            "3. 配置页确认输入文件，再回首页运行产能分析。",
        )

    def _show_error(self, title: str, message: str) -> None:
        self._log(f"{title}: {message}")
        QMessageBox.critical(self, title, message[-3000:])

    def _reset_progress(self, phase: str) -> None:
        self.progress_phase.setText(f"当前阶段：{phase}")
        self.progress_bar.setValue(0)
        self.progress_metrics.setText("进度：0/0 | 已用 0秒 | 预计剩余 计算中")
        self.window_task_count.setText("优化周期实际工序数：待运行")

    def _handle_progress(self, text: str) -> None:
        self._log(text)
        self._update_window_task_count(text)
        match = re.match(
            r"^(?P<phase>[^:：]+)[:：]\s+"
            r"(?P<current>[\d,]+)/(?P<total>[\d,]+)\s+"
            r"\((?P<percent>[\d.]+)%\)\s+\|\s+"
            r"已用\s+(?P<elapsed>[^|]+)\s+\|\s+"
            r"预计剩余\s+(?P<eta>[^|]+)"
            r"(?:\|\s+(?P<detail>.*))?$",
            text,
        )
        if not match:
            phase = text.split(":", 1)[0].split("：", 1)[0].strip()
            if phase:
                self.progress_phase.setText(f"当前阶段：{phase}")
            return
        current = int(match.group("current").replace(",", ""))
        total = int(match.group("total").replace(",", ""))
        percent = float(match.group("percent"))
        detail = (match.group("detail") or "").strip()
        self.progress_phase.setText(f"当前阶段：{match.group('phase')}")
        self.progress_bar.setValue(max(0, min(int(percent * 10), 1000)))
        metrics = (
            f"进度：{current:,}/{total:,} ({percent:.1f}%) | "
            f"已用 {match.group('elapsed').strip()} | "
            f"预计剩余 {match.group('eta').strip()}"
        )
        if detail:
            metrics += f" | {detail}"
        self.progress_metrics.setText(metrics)

    def _update_window_task_count(self, text: str) -> None:
        if "ModeB优化周期" not in text:
            return
        window_match = re.search(r"ModeB优化周期(?P<window>\d+)", text)
        window_text = f"周期 {window_match.group('window')}" if window_match else "当前周期"
        task_match = re.search(r"工序\s*(?P<count>[\d,]+)", text)
        if task_match is None:
            task_match = re.search(r"窗口工序数\s*(?P<count>[\d,]+)", text)
        if task_match is None and "本周期内无到期订单" in text:
            self.window_task_count.setText(f"优化周期实际工序数：{window_text} / 0 条")
            return
        if task_match is None:
            return
        count = int(task_match.group("count").replace(",", ""))
        status = ""
        if "成功" in text:
            status = " / 成功"
        elif "跳过" in text:
            status = " / 跳过"
        elif "失败" in text:
            status = " / 失败"
        limit = self.mode_b_max_window_tasks.value() if hasattr(self, "mode_b_max_window_tasks") else None
        limit_text = f" | 参考 {limit:,} 条" if limit else ""
        self.window_task_count.setText(f"优化周期实际工序数：{window_text} / {count:,} 条{status}{limit_text}")

    def _apply_theme(self, theme_name: str) -> None:
        if theme_name == "深色":
            self.setStyleSheet(DARK_QSS)
            return
        if theme_name == "浅色":
            self.setStyleSheet(LIGHT_QSS)
            return
        app = QApplication.instance()
        if app is not None and app.palette().color(QPalette.Window).lightness() < 128:
            self.setStyleSheet(DARK_QSS)
        else:
            self.setStyleSheet(LIGHT_QSS)

    def _begin_run_log(self, kind: str) -> None:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_kind = re.sub(r"[^A-Za-z0-9_-]+", "_", kind.strip() or "run")
        self.current_run_log_path = LOG_DIR / f"{safe_kind}_{timestamp}.log"
        header = [
            "Fortune BJ Optimizer run log",
            f"Started: {datetime.now():%Y-%m-%d %H:%M:%S}",
            f"Kind: {safe_kind}",
            f"Deploy root: {DEPLOY_ROOT}",
            "",
        ]
        self.current_run_log_path.write_text("\n".join(header), encoding="utf-8")

    def _finish_run_log(self, status: str, detail: str = "") -> None:
        path = self.current_run_log_path
        if path is None:
            return
        footer = [
            "",
            f"Finished: {datetime.now():%Y-%m-%d %H:%M:%S}",
            f"Status: {status}",
        ]
        if detail:
            footer.append(f"Detail: {detail}")
        try:
            with path.open("a", encoding="utf-8") as handle:
                handle.write("\n".join(footer))
                handle.write("\n")
        finally:
            self.current_run_log_path = None
        self._log(f"运行日志已保存：{path}")

    def _log(self, text: str) -> None:
        line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {text}"
        if hasattr(self, "log"):
            self.log.append(line)
        path = self.current_run_log_path
        if path is not None:
            try:
                with path.open("a", encoding="utf-8") as handle:
                    handle.write(line)
                    handle.write("\n")
            except OSError:
                pass


def main() -> int:
    if not PYSIDE6_AVAILABLE:
        raise RuntimeError("缺少 PySide6，请先安装 requirements.txt 后再启动 Fortune BJ 桌面应用。")
    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
