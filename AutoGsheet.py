import gspread

share_link = "https://docs.google.com/spreadsheets/d/1k2yjKUA4m886_Vj24V_jnpt4ClTkQIjoabBBQKb4Jko/edit"
gc = gspread.service_account(filename='token.json')
worksheet = gc.open_by_url(share_link).worksheet('01.06(二)')
for v in worksheet.col_values(3)[4:]:
    print(v)

