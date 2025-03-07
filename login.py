from openpyxl import load_workbook

workbook = load_workbook('L3巡检密码.xlsx')
sheet = workbook.active
# 网络审计
WLSJ_url = sheet['B2'].value
WLSJ_username = sheet['C2'].value
WLSJ_password = sheet['D2'].value
# 日志收集
RZSJ_url = sheet['B3'].value
RZSJ_username = sheet['C3'].value
RZSJ_password = sheet['D3'].value
# 漏洞扫描
LDSM_url = sheet['B4'].value
LDSM_username = sheet['C4'].value
LDSM_password = sheet['D4'].value
# IPS
IPS_url = sheet['B5'].value
IPS_username = sheet['C5'].value
IPS_password = sheet['D5'].value
# WAF
WAF_url = sheet['B6'].value
WAF_username = sheet['C6'].value
WAF_password = sheet['D6'].value
# 防病毒软件
FBD_url = sheet['B7'].value
FBD_username = sheet['C7'].value
FBD_password = sheet['D7'].value
# 隔离区防病毒软件
GLQFBD_url = sheet['B12'].value
GLQFBD_username = sheet['C12'].value
GLQFBD_password = sheet['D12'].value
# L3虚拟平台
L3YPT_url = sheet['B8'].value
L3YPT_username = sheet['C8'].value
L3YPT_password = sheet['D8'].value
