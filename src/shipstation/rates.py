import os
import requests
from datetime import date, timedelta, datetime
from pathlib import Path
from dotenv import load_dotenv
import json

SHIPMENT_URL="https://api.shipstation.com/v2/shipments"
RATES_URL = "https://api.shipstation.com/v2/rates"

URL = "https://api.shipstation.com/v2/rates/estimate"

CARRIER_MAP = {
    "usps": "se-167930",
    "ups": "se-196204"
}

# 1. Get the path of the script (check_codes.py)
script_path = Path(__file__).resolve()
# 2. Go up two levels to reach 'daily_tool' folder
project_root = script_path.parent.parent.parent
# 3. Join with '.env'
env_path = project_root / '.env'
# 4. Load it explicitly
load_dotenv(dotenv_path=env_path)

V2_API_KEY = os.getenv("SHIPSTATION_V2_PRODUCTION_KEY")
V1_SHIPSTATION_API_KEY=os.getenv("SHIPSTATION_API_KEY")
V1_SHIPSTATION_API_SECRET=os.getenv("SHIPSTATION_API_SECRET")

def get_live_rates(order_no,carrier, service, pkg, weight, dims=None, to_state="CA", to_zip="90058",is_residential=False):

    """
        Fetches real-time shipping rates from the ShipStation V2 API.
        Process follows 'Create -> Rate -> Cancel' Shipment Workflow:

        1. Create a temporary shipment to generate a shipment_id
        2. Request rates for that specific shipment_id
        3. Cancel the temporary shipment so it doesnt appear in Shipstation
        4. Filter and standardize the results for the main application

        Returns: tuple: (list of processed_rate_dicts, boolean is_residential)
    """

    print(f"DEBUG get_live_rates|Order:{order_no} entered")

    order_info = get_order_address(order_no)
    if not order_info:
        return [], is_residential
    
    addr = order_info['ship_to']
    carrier_id = CARRIER_MAP.get(carrier.lower())
    if not carrier_id:
        print(f"Unknown carrier for V2: {carrier}")
        return [], is_residential
    
    flat_rate_codes = ['flat_rate_padded_envelope', 'flat_rate_envelope', 'medium_flat_rate_box', 'large_flat_rate_box']

    if pkg in flat_rate_codes:
        print(f"DEBUG get_rate_estimate|Order:{order_no} entering get_rate_estimate")
        est_result = get_rate_estimate(carrier_id, service, pkg, weight, dims, to_state, to_zip, addr)
        
        # 2. If it's a tuple, just take the first part (the list)
        if isinstance(est_result, tuple):
            est_result = est_result[0]
            
        # 3. Now return it safely
        return est_result, is_residential
        
    payload = {
        "shipments": [{
            "validate_address": "no_validation",
            "ship_to": {
                "name": "Test Customer",
                "address_line1": addr.get("street1"),
                "address_line2": addr.get("street2"),
                "city_locality": addr.get("city"),
                "state_province": to_state,
                "postal_code": str(to_zip)[:5],
                "country_code": "US"
            },
            "ship_from": {
                "name": "3317 E 50th St",
                "phone": "323-510-3700",
                "address_line1": "3317 E 50th St",
                "city_locality": "Vernon",
                "state_province": "CA",
                "postal_code": "90058",
                "country_code": "US"
            },
            "packages": [{
                "weight": {"value": float(weight), "unit": "pound"},
                "dimensions": {
                    "unit": "inch",
                    "length": dims[0] if dims else 1,
                    "width": dims[1] if dims else 1,
                    "height": dims[2] if dims else 1
                }
            }]
        }]
    }

    headers = {
        "api-key": V2_API_KEY,
        "Content-Type": "application/json"
    }

    try:
        ship_response = requests.post(SHIPMENT_URL,json=payload, headers=headers)
        
        if ship_response.status_code != 200:
            print(f"Shipment Creation Failed: {ship_response.text}")
            return [], is_residential
        
        shipment_data = ship_response.json()

        shipments_list = shipment_data.get("shipments",[])

        if not shipments_list:
            print(f"Error: No shipments returned in the response for {order_no}")
            return [], is_residential
        
        residential_indicator = shipments_list[0].get("ship_to", {}).get("address_residential_indicator", "unknown")

        if residential_indicator in ["no","unknown"]:
            is_residential = False
        else:
            is_residential = True
        
        shipment_id = shipments_list[0].get("shipment_id")

        if not shipment_id:
            print(f"Error: shipment_id missing from the first shipment for {order_no}")
            return [], is_residential

        rate_payload = {
            "shipment_id": shipment_id,
            "rate_options": {
                "carrier_ids": [carrier_id]
            }
        }

        rate_response = requests.post(RATES_URL,json=rate_payload, headers=headers)
        cancel_url = f"https://api.shipstation.com/v2/shipments/{shipment_id}/cancel"
        requests.put(cancel_url, headers=headers)

        if rate_response.status_code == 200:
            rate_data = rate_response.json()

            rates = rate_data.get("rate_response", {}).get("rates",[])

            processed_rates = []
            for r in rates:
                # Calculate cost once here so main.py doesn't have to
                base = r.get("shipping_amount", {}).get("amount", 0.0)
                other = r.get("other_amount", {}).get("amount", 0.0)
                
                # Build a standardized dictionary for main.py
                processed_rates.append({
                    "shipmentCost": round(base + other, 2),
                    "serviceCode": r.get("service_code"),
                    "serviceName": r.get("service_type"),
                    "carrierCode": r.get("carrier_code"),
                    "packageType": r.get("package_type"),
                    "estimated_delivery_date": r.get("estimated_delivery_date"),
                    "comparison_log": f"{r.get('service_name')} (Direct)",
                    "realName": r.get("service_name") or r.get("service_type")
                })

           # ONLY RETURN THE SERVICES WE ACTUALLY CARE ABOUT
            target_services = ['usps_ground_advantage', 'usps_priority_mail', 'ups_ground', 'ups_ground_saver','usps_first_class_mail']
            
            is_standard_search = any(x in pkg.upper() for x in ["OZ", "PACKAGE"])

            if not service: # If shopping/optimizing
                return [
                    r for r in processed_rates 
                    if r['serviceCode'] in target_services 
                    and (
                        # Match A: Strict match for Flat Rates (Fixes your $10.30 problem)
                        r['packageType'] == pkg.lower() 
                        or 
                        # Match B: Fuzzy match for OZ/Standard (Fixes your None problem)
                        (is_standard_search and r['packageType'] in ["package", "parcel", None])
                    )
                ], is_residential
            
            filtered_results = []
            for r in processed_rates:
                service_match = r['serviceCode'] == service.lower()

                if is_standard_search:
                    # If SKU says "5-8 OZ", look for "package" in API results
                    package_match = r['packageType'] in ["package", "parcel", None]
                else:
                    # If SKU says "flat_rate_envelope", look for exact match
                    package_match = r['packageType'] == pkg.lower()

                if service_match and package_match:
                    filtered_results.append(r)
            
            return filtered_results, is_residential
        else:
            print(f"V2 Error {rate_response.status_code}: {rate_response.text}")
            return [], is_residential
    except Exception as e:
        print(f"V2 Connection Error: {e}")
        return [], is_residential
    
def get_order_address(order_no):
    url = f"https://ssapi.shipstation.com/orders?orderNumber={order_no}"

    response = requests.get(url, auth=(V1_SHIPSTATION_API_KEY, V1_SHIPSTATION_API_SECRET))

    if response.status_code == 200:
        data = response.json()
        if data.get('orders'):
            # V1 returns a list; we grab the first match
            order = data['orders'][0]
            
            return {
                "order_id": order.get("orderId"),
                "ship_to": order.get("shipTo")
            }
    return None

def get_rate_estimate(carrier_id, service, pkg, weight, dims, to_state, to_zip, addr):
    headers = {
        "api-key": V2_API_KEY,
        "Content-Type": "application/json"
    }

    ship_date_str = datetime.now().strftime("%Y-%m-%dT15:00:00.000Z")

    payload = {
        "carrier_id": carrier_id,
        "from_country_code": "US",
        "from_postal_code": "90058",
        "from_city_locality": "VERNON",
        "from_state_province": "CA",
        "to_city_locality": addr.get("city"),
        "to_address_line1": addr.get("street1"),
        "to_state": to_state,
        "to_postal_code": str(to_zip)[:5],
        "to_country_code": "US",
        "weight": {"value": float(weight), "unit": "pound"},
        "confirmation": "none",
        "address_residential_indicator": "unknown",
        "ship_date": ship_date_str
    }

    if pkg in ['flat_rate_padded_envelope', 'flat_rate_envelope', 'medium_flat_rate_box', 'large_flat_rate_box']:
        payload["package_code"] = pkg
    else:
        payload["dimensions"] = {
            "unit": "inch",
            "length": dims[0] if dims else 1,
            "width": dims[1] if dims else 1,
            "height": dims[2] if dims else 1
        }

    try:
        response = requests.post(URL, json=payload, headers=headers)
        if response.status_code == 200:
            rates = response.json()
            
            processed_rates = []
            for r in rates:
                api_service = r.get("service_code")
                api_pkg = r.get("package_type")

                # 1. MUST match the service (e.g., usps_priority_mail)
                if service and api_service != service.lower():
                    continue

                # IMPROVED BOUNCER:
                # 1. If it's a Flat Rate, it MUST be a strict match (Fixes the $10.30 vs $11.10)
                # 2. If it's standard Ground, allow 'package' OR 'parcel'
                if pkg:
                    target_pkg = pkg.lower()
                    current_api_pkg = api_pkg.lower() if api_pkg else ""
                    
                    if "flat_rate" in target_pkg:
                        if current_api_pkg != target_pkg:
                            continue
                    else:
                        # For Ground/UPS, 'package' and 'parcel' are the same thing
                        if target_pkg == "package" and current_api_pkg not in ["package", "parcel"]:
                            continue
                        elif target_pkg != "package" and current_api_pkg != target_pkg:
                            continue

                base = r.get("shipping_amount", {}).get("amount", 0.0)
                other = r.get("other_amount", {}).get("amount", 0.0)
                total_cost = round(float(base) + float(other), 2)

                processed_rates.append({
                    "shipmentCost": total_cost,
                    "serviceCode": api_service,
                    "serviceName": r.get("service_type"),
                    "carrierCode": r.get("carrier_code"),
                    "packageType": pkg,
                    "estimated_delivery_date": r.get("estimated_delivery_date"),
                    "comparison_log": f"{r.get('service_type')} ({api_pkg}) (only)"
                })

            return processed_rates, False
        else:
            print(f"Estimate API Error: {response.text}")
            return [], False
    except Exception as e:
        print(f"Estimate Connection Error: {e}")
        return [], False