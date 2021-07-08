import pandas as pd

projectId = 202356
clientName = '自家品牌'
productCampaign = '減肥/夏季之旅'
period = '2021/06/01-2021/07/01'
adsDays = 30

# Create a Pandas dataframe from some data.
df = pd.DataFrame({
    '': ['Budget', 'Period', '', 'Impression', 'Click', 'CTR', '觀看數', 'PV', '總觸及人數']
    })

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('pandas_image.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='總表', startrow=10, index=False)

# 設定 xlsxwriter 的 sheets 內容
workbook  = writer.book
worksheet = writer.sheets['總表']


worksheet.set_row(1, cell_format=workbook.add_format({'bg_color': '#21AD25'})) #設定第一行的高度為40

worksheet.write(2, 0, '專案編號(NO.)')
worksheet.write(2, 1, projectId)
worksheet.write(3, 0, '客戶(Client) ')
worksheet.write(3, 1, clientName)
worksheet.write(4, 0, '產品(Product) / 活動(Campaign) ')
worksheet.write(4, 1, productCampaign)
worksheet.write(5, 0, '走期(Period) ')
worksheet.write(5, 1, period)
worksheet.write(6, 0, '廣告天數')
worksheet.write(6, 1, adsDays)

worksheet.set_row(7, cell_format=workbook.add_format({'bg_color': '#21AD25'})) #設定第一行的高度為40

# 設定 Logo 圖片
worksheet.set_column(0, 0, 30) #設定第一列寬為 30
worksheet.set_row(0, 80)       #設定第一行的高度為 80
imgOption = {
    'x_offset': 40, #水平偏移
    'y_offset': 14, #垂直偏移
}
worksheet.insert_image(0, 0, 'img/logo.png', imgOption)

# 關閉寫入, 並存檔
writer.save()