# 学生缺勤登记系统 启动脚本（PowerShell 5 语法）
Set-Location -Path $PSScriptRoot
conda run --no-capture-output -n absence_tracker python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
