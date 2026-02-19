import random
from datetime import date, timedelta
import config

# Reuse your SKU and store mapping
SKUS1 = [
    "MK1059-03","MK1060-03","MK1061-03","MK1062-06","MK1063-04","MK1064-05","MK1065-05","MK1066-04",
    "MK1067-03","MK1068-03","MK1069-04","MK1070-07","MK1071-06","MK1072-05","MK1073-04","MK1074-03",
    "MK1075-06","MK1076-06","MK1077-06","MK1078-05","MK1079-04","MK1080-04","MK1081-05","MK1082-06",
    "MK1083-05","MK1084-04","MK1085-05","MK1086-03","MK1087-03","MK1088-03","MK1089-03","MK1090-04",
    "MK1091-06","MK1092-05","MK1093-04","MK1094-05","MK1095-03","MK1096-03","MK1097-03","MK1098-03",
    "MK1099-04","MK1100-06","MK1101-05","MK1102-04","MK1103-04","MK1104-03","MK1105-03","MK1106-06",
    "MK1107-04","MK1108-04","MK1109-06","MK1110-05","MK1111-04","MK1112-04","MK1113-03","MK1114-03",
    "MK1115-04","MK1116-02","MK1117-02","MK1118-02","MK1119-02","MK1120-02","MK1121-03","MK1122-03",
    "MK1123-02","MK1124-02","MK1125-03","MK1126-02","MK1127-04","MK1128-03","MK1129-03","MK1130-02",
    "MK1131-03","MK1132-03","MK1133-02","MK1134-02","MK1135-04","MK1136-03","MK1137-03","MK1138-03",
    "MK1139-03","MK1140-02","MK1141-02","MK1142-04","MK1143-03","MK1144-03","MK1145-03","MK1146-03",
    "MK1147-02","MK1148-02","MK1149-04","MK1150-04","MK1151-03","MK1152-03","MK1153-03","MK1154-03",
    "MK1155-02","MK1156-02","MK1157-04","MK1158-03","MK1159-03","MK1160-03","MK1161-03","MK1162-02",
    "MK1163-02","MK1164-02","MK1165-04","MK1166-03","MK1167-03","MK1168-03","MK1169-03","MK1170-02",
    "MK1171-02","MK1172-02","MK1173-02","MK1174-02","MK1175-06","MK1176-04","MK1177-05","MK1178-05",
    "MK1179-04","MK1180-03","MK1181-02","MK1182-03","MK1183-04","MK1184-03","MK1185-03","MK1186-03",
    "MK1187-03","MK1188-02","MK1189-02","MK1190-02","MK1191-02","MK1192-02","MK1193-04","MK1194-03",
    "MK1195-03","MK1196-03","MK1197-03","MK1198-02","MK1199-02","MK1200-02","MK1201-02","MK1202-02",
    "MK1203-04","MK1204-03","MK1205-03","MK1206-03","MK1207-03","MK1208-02","MK1209-02","MK1210-02",
    "MK1211-02","MK1212-04","MK1213-03","MK1214-03","MK1215-03","MK1216-02","MK1217-02","MK1218-02",
    "MK1226-04","MK1227-03","MK1228-03","MK1229-03","MK1230-03","MK1231-02","MK1232-02","MK1233-02",
    "MK1234-03","MK1235-03","MK1236-02","MK1237-03","MK1238-03","MK1239-02","MK1240-02","MK1241-02",
    "MK1242-04","MK1243-03","MK1244-03","MK1245-03","MK1246-03","MK1247-02","MK1248-02","MK1249-02",
    "MK1250-02","MK1251-02","MK1252-02","MK1253-04","MK1254-03","MK1255-03","MK1256-03","MK1257-03",
    "MK1258-02","MK1259-02","MK1260-02","MK1261-02","MK1262-02","MK1263-02","MK1264-04","MK1265-03",
    "MK1266-03","MK1267-03","MK1268-03","MK1269-02","MK1270-02","MK1271-02","MK1272-02","MK1273-02",
    "MK1274-02","MK1275-03","MK1276-02","MK1277-02","MK1278-02","MK1279-03","MK1280-02","MK1281-02",
    "MK1282-04","MS1516","MS1517","MS1518","MS1519","MS1520","MS1521","MS1522","MS1523","MS1524","MS1525",
    "MS1526","MS1527","MS1528","MS1529","MS1530","MS1531","MS1532","MS1533","MS1534","MS1535","MS1536",
    "MS1537","MS1538","MS1539"
]

SKUS2 = ["MS2920"]

STATES = ["CA", "NY", "TX", "FL", "IL", "PA", "OH", "GA", "NC", "MI"]
FIRST_NAMES = ["John", "Jane", "Alice", "Bob", "Charlie", "Diana", "Eve", "Frank"]
LAST_NAMES = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Miller", "Davis", "Garcia"]

def generate_test_orders(num_orders):
    """
    Returns a list of fake orders in the same structure as get_shipments() returns,
    so that extract_todays_shipments() can process them directly.
    """
    orders = []

    for i in range(1, num_orders + 1):
        store_id = random.choice(list(config.STORE_MAP.keys()))
        num_items = random.randint(1, 3)  # some orders have multiple SKUs
        items = []
        for _ in range(num_items):
            sku = random.choice(SKUS2)
            #qty = random.randint(1, 2)
            qty = 1
            items.append({"sku": sku, "quantity": qty})

        # Fake customer info
        first = random.choice(FIRST_NAMES)
        last = random.choice(LAST_NAMES)

        # shipByDate = today + 0-2 days
        ship_by_date = (date.today() + timedelta(days=random.randint(0, 2))).isoformat()

        order = {
            "orderNumber": f"ORD{i:04d}",
            "orderDate": date.today().isoformat(),
            "shipByDate": ship_by_date,
            "shipTo": {
                "name": f"{first} {last}",
                "state": random.choice(STATES)
            },
            "advancedOptions": {
                "storeId": store_id
            },
            "items": items
        }
        orders.append(order)

    return orders
