@echo off
chcp 65001 >nul
title JO - 客戶滿意度問卷資料整合系統 (CSAT Consolidator)

if "%~1"=="" (
    python main.py
) else (
    python main.py "%~1"
)
