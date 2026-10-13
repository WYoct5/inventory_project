import csv
from openpyxl import Workbook


stock_data = {}

with open("data/stock.csv",encoding="utf-8") as f:
    reader = csv.DictReader(f)

    for row in reader:
        stock = int(row["stock"])
        stock_data[row["product"]] = stock

    print(stock_data)



shipment_totals = {}

with open("data/shipments.csv",encoding="utf-8") as f:
    reader = csv.DictReader(f)

    for row in reader:
        product = row["product"]
        quantity = int(row["quantity"])

        if product in shipment_totals:
            shipment_totals[product] += quantity
        else:
            shipment_totals[product] = quantity

    print(shipment_totals)

days = 6
result = []

for product,stock in stock_data.items():
    
    if product in shipment_totals:
        shipment = shipment_totals[product]
    else:
        shipment = 0

    average_daily_shipment = shipment / days

    if average_daily_shipment > 0:
        days_remaining = stock / average_daily_shipment

        if days_remaining <= 7:
            risk = "危険"
        elif days_remaining <= 14:
            risk = "注意"
        else:
            risk = "問題なし"

    else:
        days_remaining = None
        risk = "出荷なし"

    result.append({
        "product": product,
        "stock": stock,
        "shipment": shipment,
        "average_daily_shipment": average_daily_shipment,
        "days_remaining": days_remaining,
        "risk": risk
    })

for row in result:

    days_remaining = row["days_remaining"]

    if days_remaining is None:
        days_remaining = "-"
    else:
        days_remaining = round(days_remaining,1)


    print(
        row["product"],
        row["stock"],
        row["shipment"],
        round(row["average_daily_shipment"],1),
        days_remaining,
        row["risk"]
    )

    wb = Workbook()
    ws = wb.active

    ws.append([
        "商品コード",
        "現在庫",
        "出荷合計",
        "1日平均出荷数",
        "残り日数",
        "リスク",
    ])

    for row in result:
        ws.append([
            row["product"],
            row["stock"],
            row["shipment"],
            round(row["average_daily_shipment"],1),
            row["days_remaining"],
            row["risk"]
        ])

wb.save("inventory_result.xlsx")
print("Excelファイルを保存しました")