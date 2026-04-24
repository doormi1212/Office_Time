# -*- coding: utf-8 -*-
"""
应用配置模块：数据目录、Excel 文件名、上传限制等。

前后端分离后由 FastAPI（api.py）读取本配置；部署到服务器时可用环境变量覆盖路径。
"""
import os


class BaseConfig:
    """基础配置。"""

    # 上传体积限制（证明文件等），单位：字节（默认 16MB）
    MAX_CONTENT_LENGTH = int(os.environ.get("MAX_CONTENT_LENGTH", 16 * 1024 * 1024))

    # -------------------------------------------------------------------------
    # 路径配置（均相对于项目根目录，除非给出绝对路径）
    # -------------------------------------------------------------------------
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))

    # 学生志愿时长总表目录
    DATA_DIR = os.environ.get("DATA_DIR", os.path.join(BASE_DIR, "data"))
    
    @staticmethod
    def get_master_excel_path(grade: str) -> str:
        """根据年级获取对应的总表路径。"""
        # 兼容 2022, 2023, 2024, 2025
        filename = f"学生志愿时长总表_{grade}.xlsx"
        return os.path.join(BaseConfig.DATA_DIR, filename)

    # 默认总表路径（兼容旧逻辑或作为默认值）
    MASTER_EXCEL_FILENAME = "学生志愿时长总表.xlsx"
    MASTER_EXCEL_PATH = os.path.join(DATA_DIR, MASTER_EXCEL_FILENAME)

    # 学生反馈汇总表（追加写入）
    FEEDBACK_EXCEL_FILENAME = "学生反馈汇总.xlsx"
    FEEDBACK_EXCEL_PATH = os.path.join(DATA_DIR, FEEDBACK_EXCEL_FILENAME)

    # 学生时长补录汇总表（校内/校外追加写入）
    TOPUP_EXCEL_FILENAME = "学生时长补录汇总.xlsx"
    TOPUP_EXCEL_PATH = os.path.join(DATA_DIR, TOPUP_EXCEL_FILENAME)

    # 管理员账号配置文件（JSON）
    ADMIN_USERS_FILENAME = "admin_users.json"
    ADMIN_USERS_PATH = os.path.join(DATA_DIR, ADMIN_USERS_FILENAME)

    # 学生上传的证明文件目录
    PROOF_UPLOAD_DIR = os.environ.get(
        "PROOF_UPLOAD_DIR", os.path.join(BASE_DIR, "uploads", "proofs")
    )

    # 校内/校外证明文件子目录（便于后续人工核对与权限控制）
    TOPUP_PROOF_INTERNAL_DIR = os.path.join(PROOF_UPLOAD_DIR, "topup_internal")
    TOPUP_PROOF_EXTERNAL_DIR = os.path.join(PROOF_UPLOAD_DIR, "topup_external")

    # 允许上传的证明文件扩展名（可按业务扩展）
    ALLOWED_PROOF_EXTENSIONS = {"pdf", "png", "jpg", "jpeg", "gif", "webp"}

    # 允许管理员上传的总表扩展名（当前以 xlsx 为主，读表依赖 openpyxl）
    ALLOWED_MASTER_EXTENSIONS = {"xlsx"}


class DevelopmentConfig(BaseConfig):
    DEBUG = True


class ProductionConfig(BaseConfig):
    DEBUG = False


# 根据环境变量 FLASK_ENV 选择配置（默认可用于 create_app）
config_by_name = {
    "development": DevelopmentConfig,
    "production": ProductionConfig,
    "default": DevelopmentConfig,
}


def get_config(name=None):
    """供 create_app 使用：返回配置类。"""
    if name is None:
        name = os.environ.get("FLASK_ENV", "development")
    return config_by_name.get(name, DevelopmentConfig)
