import os
import requests
from pathlib import Path
from dotenv import load_dotenv

current_dir = Path(__file__).resolve()
project_root = next(p for p in current_dir.parents if (p / '.env').exists())
env_path = project_root / '.env'

load_dotenv(dotenv_path=env_path)

BASE_URL = "https://ssapi.shipstation.com"

API_KEY = os.getenv("SHIPSTATION_API_KEY")
API_SECRET = os.getenv("SHIPSTATION_API_SECRET")

def get_shipments():

    orders = []
    page = 1

    while True:
        r = requests.get(
            f"{BASE_URL}/orders",
            auth=(API_KEY, API_SECRET),
            params={
                "orderStatus": "awaiting_shipment",
                "pageSize": 100,
                "page": page
            }
        )

        r.raise_for_status
        data = r.json()

        orders.extend(data.get("orders", []))

        if page >= data.get("pages",1):
            break

        page += 1

    return orders