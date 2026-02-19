import os
import requests
from pathlib import Path
from dotenv import load_dotenv

# 1. Get the path of the script (check_codes.py)
script_path = Path(__file__).resolve()

# 2. Go up two levels to reach 'daily_tool' folder
project_root = script_path.parent.parent.parent

# 3. Join with '.env'
env_path = project_root / '.env'

# 4. Load it explicitly
load_dotenv(dotenv_path=env_path)

API_KEY = os.getenv("SHIPSTATION_API_KEY")
API_SECRET = os.getenv("SHIPSTATION_API_SECRET")
URL_BASE = 'https://ssapi.shipstation.com'

def list_codes():
    print("Connecting to ShipStation...")
    try:
        res = requests.get(f"{URL_BASE}/carriers", auth=(API_KEY, API_SECRET))
        
        # If the status is not 200, print the error text and stop
        if res.status_code != 200:
            print(f"FAILED! Status Code: {res.status_code}")
            print(f"Response from Server: {res.text}")
            return

        carriers = res.json()
        for carrier in carriers:
            c_code = carrier['code']
            print(f"\n[CARRIER: {c_code}]")

            # Get Services
            s_res = requests.get(f"{URL_BASE}/carriers/listservices?carrierCode={c_code}", auth=(API_KEY, API_SECRET))
            if s_res.status_code == 200:
                for s in s_res.json():
                    print(f"  - Service: {s['code']}")
            
            # Get Packages
            p_res = requests.get(f"{URL_BASE}/carriers/listpackages?carrierCode={c_code}", auth=(API_KEY, API_SECRET))
            if p_res.status_code == 200:
                for p in p_res.json():
                    print(f"  - Package: {p['code']}")

    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    list_codes()