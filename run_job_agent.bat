@echo off
cd /d "%~dp0"
".venv\Scripts\python.exe" java_contract_job_agent.py >> job_agent_log.txt 2>&1
