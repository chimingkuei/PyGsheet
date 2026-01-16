import gspread

share_link = "https://docs.google.com/spreadsheets/d/1k2yjKUA4m886_Vj24V_jnpt4ClTkQIjoabBBQKb4Jko/edit"
gc = gspread.service_account(filename='token.json')
worksheet = gc.open_by_url(share_link).sheet1
for v in worksheet.col_values(3)[4:]:
    print(v)

# import schedule
# import time

# def job():
#     print("hello")

# # schedule 預設最小單位是秒，所以可以用 every(3).seconds
# schedule.every(3).seconds.do(job)

# while True:
#     schedule.run_pending()
#     time.sleep(1)  # 避免 CPU 100%
