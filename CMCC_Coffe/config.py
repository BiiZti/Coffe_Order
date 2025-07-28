#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CMCC Coffee 订单管理系统 - 配置文件
"""

import os

class Config:
    """基础配置类"""
    
    # Flask基础配置
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'cmcc-coffee-secret-key-2024'
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size
    SEND_FILE_MAX_AGE_DEFAULT = 0
    
    # 服务器配置
    HOST = '0.0.0.0'
    PORT = 5000
    DEBUG = False
    THREADED = True
    PROCESSES = 1
    
    # 连接池配置
    MAX_CONNECTIONS = 100
    CONNECTION_TIMEOUT = 30
    
    # Excel文件配置
    EXCEL_RETRY_COUNT = 3
    EXCEL_RETRY_DELAY = 2  # 秒
    EXCEL_LOCK_TIMEOUT = 30  # 秒
    
    # 请求频率限制
    REQUEST_RATE_LIMIT = 5  # 秒
    
    # 数据刷新间隔
    DATA_REFRESH_INTERVAL = 60  # 秒

class DevelopmentConfig(Config):
    """开发环境配置"""
    DEBUG = True
    THREADED = True

class ProductionConfig(Config):
    """生产环境配置"""
    DEBUG = False
    THREADED = True

# 配置映射
config = {
    'development': DevelopmentConfig,
    'production': ProductionConfig,
    'default': DevelopmentConfig
} 