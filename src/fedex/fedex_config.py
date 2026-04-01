from datetime import date, datetime
import time
import requests
import os
from dotenv import load_dotenv, find_dotenv
from concurrent.futures import ThreadPoolExecutor

from src.shipstation.rates import get_live_rates, get_order_address

load_dotenv(find_dotenv())

SERVICE_MAP = {
    "fedex_home_delivery": {
        "ss": "fedex_home_delivery",
        "direct": "GROUND_HOME_DELIVERY"
    },
    "fedex_ground": {
        "ss": "fedex_ground",
        "direct": "FEDEX_GROUND"
    },
    "fedex_ground_economy": {
        "ss": "fedex_ground_economy_parcel_select",
        "direct": "SMART_POST"
    },
    "fedex_2day_onerate": {
        "ss": "fedex_2day_one_rate"
    },
    "fedex_express_saver_onerate": {
        "ss": "fedex_economy_one_rate"
    }
}

pkg_map = {
        "F": ("FEDEX_ENVELOPE", [0, 0, 0]),
        "P": ("FEDEX_PAK", [0, 0, 0]),
        "Q1R": ("FEDEX_SMALL_BOX", [12, 9, 3]),
        "Q1F": ("FEDEX_SMALL_BOX", [12, 9, 3]),
        "Q1C": ("FEDEX_MEDIUM_BOX", [12, 9, 5]),
        "Q2R": ("FEDEX_MEDIUM_BOX", [12, 9, 5]),
        "Q2C": ("FEDEX_MEDIUM_BOX", [12, 9, 5]),
        "10x10x10": ("FEDEX_EXTRA_LARGE_BOX", [12, 11, 12]),
        "Q3C": ("FEDEX_EXTRA_LARGE_BOX", [12, 11, 12]),
        "Q5C": ("FEDEX_EXTRA_LARGE_BOX", [16, 15, 6]),
        "L": ("FEDEX_EXTRA_LARGE_BOX", [16, 15, 6]),
        "Q5F": ("FEDEX_EXTRA_LARGE_BOX", [16, 15, 6]),
        "Q4R": ("FEDEX_EXTRA_LARGE_BOX", [16, 15, 6]),
        "Q4F": ("FEDEX_EXTRA_LARGE_BOX", [16, 15, 6]),
        "Q3F": ("FEDEX_EXTRA_LARGE_BOX", [16, 15, 6]),
        "Q3R": ("FEDEX_EXTRA_LARGE_BOX", [16, 15, 6]),
        "Q2F": ("FEDEX_EXTRA_LARGE_BOX", [16, 15, 6]),
    }

session = requests.Session()
_FEDEX_TOKEN_CACHE = {"token": None, "expires": 0}
def get_fedex_token():
    global _FEDEX_TOKEN_CACHE

    now = time.time()
    if _FEDEX_TOKEN_CACHE["token"] and now < _FEDEX_TOKEN_CACHE["expires"]:
        return _FEDEX_TOKEN_CACHE["token"]
    
    # Use "https://apis-sandbox.fedex.com" for testing
    url = "https://apis.fedex.com/oauth/token"
    payload = {
        'grant_type': 'client_credentials',
        'client_id': os.getenv('FEDEX_KEY'),
        'client_secret': os.getenv('FEDEX_SECRET')
    }
    headers = {'Content-Type': "application/x-www-form-urlencoded"}
    response = session.post(url, data=payload, headers=headers)
    token = response.json().get('access_token')
    _FEDEX_TOKEN_CACHE = {"token": token, "expires": now + 3000} # Cache for 50 mins
    return token

def get_fedex_rate(payload):
    token = get_fedex_token()
    if not token:
        return {"error": "Could not retrieve access token"}

    # Use "https://apis-sandbox.fedex.com/rate/v1/rates/quotes" for sandbox
    url = "https://apis.fedex.com/rate/v1/rates/quotes"
    
    headers = {
        'Content-Type': "application/json",
        'Authorization': f"Bearer {token}"
    }

    response = session.post(url, json=payload, headers=headers)
    
    if response.status_code == 200:
        return response.json()
    else:
        return {
            "error": f"API Error {response.status_code}",
            "details": response.text
        }
    
def get_days_from_today(arrival_date_str):
    if not arrival_date_str:
        return None
    try:
        # APIs often return ISO format: 2026-02-26T00:00:00Z
        arrival_date = datetime.fromisoformat(arrival_date_str.replace('Z', '+00:00')).date()
        today = date.today()
        delta = (arrival_date - today).days
        return delta if delta >= 0 else 0
    except Exception:
        return None
    
def get_fedex_comparison_logic(order_no, weight, dims, winning_pkg_str):
    addr_data = get_order_address(order_no)
    if not addr_data:
        return 0.0, 0.0, 0.0, "", "Address Fetch Error"
    
    ship_to = addr_data['ship_to']
    is_res = ship_to.get("residential", False)

    active_services = SERVICE_MAP.copy()
    if is_res:
        active_services.pop("fedex_ground", None)
    else:
        active_services.pop("fedex_home_delivery", None)

    all_quotes = []
    base_days = 0 

    # --- 1. PARALLELIZE SHIPSTATION CALLS ---
    def fetch_ss(service_key, ss_service):
        try:
            ss_res, _ = get_live_rates(
                order_no, addr_data, "fedex", ss_service, "package", weight, dims,
                ship_to.get("state"), ship_to.get("postalCode"), is_residential=is_res
            )
            if ss_res:
                rate = ss_res[0]
                return {
                    "source": "SS",
                    "service": service_key,
                    "price": float(rate["shipmentCost"]),
                    "days": get_days_from_today(rate.get("estimated_delivery_date"))
                }
        except Exception as e:
            print(f"SS Thread Error: {e}")
        return None

    ss_tasks = []
    for service_key, codes in active_services.items():
        if service_key == "fedex_ground" or service_key == "fedex_home_delivery":
            ss_service = "fedex_home_delivery" if is_res else "fedex_ground"
        else:
            ss_service = codes["ss"]
        ss_tasks.append((service_key, ss_service))

    with ThreadPoolExecutor(max_workers=len(ss_tasks)) as executor:
        results = executor.map(lambda p: fetch_ss(*p), ss_tasks)
        for res in results:
            if res:
                all_quotes.append(res)
                # Capture base_days for Economy logic from Ground/Home Delivery
                if res['service'] in ["fedex_ground", "fedex_home_delivery"] and res['days']:
                    base_days = res['days']

    # Define Packaging Mapping
    # Logic: (Fedex Packaging Constant, Manual Dimensions Override)
    selected_pkg, override_dims = pkg_map.get(winning_pkg_str, ("YOUR_PACKAGING", dims))
    final_dims = override_dims if override_dims else dims

    # --- 2. SINGLE CALL FOR ALL FEDEX DIRECT RATES ---
    # We omit serviceType to get everything back at once
    fedex_payload = {
        "accountNumber": {"value": os.getenv("FEDEX_ACCOUNT")},
        "rateRequestControlParameters": {"returnTransitTimes": True},
        "requestedShipment": {
            "preferredCurrency": "USD",
            "rateRequestType": ["ACCOUNT"],
            "shipper": {
                "address": {
                    "streetLines": ["3317 E 50th St"],
                    "city": "Vernon",
                    "stateOrProvinceCode": "CA",
                    "postalCode": "90058",
                    "countryCode": "US"
                }
            },
            "recipient": {
                "address": {
                    "streetLines": [ship_to.get("street1"), ship_to.get("street2")],
                    "city": ship_to.get("city"),
                    "stateOrProvinceCode": ship_to.get("state"),
                    "postalCode": ship_to.get("postalCode"),
                    "countryCode": "US",
                    "residential": is_res
                }
            },
            "pickupType": "DROPOFF_AT_FEDEX_LOCATION",
            "packagingType": selected_pkg,
            "requestedPackageLineItems": [
                {
                    "weight": {"units": "LB", "value": weight},
                    "dimensions": {
                        "length": final_dims[0] if final_dims else 1,
                        "width": final_dims[1] if final_dims else 1,
                        "height": final_dims[2] if dims else 1,
                        "units": "IN"
                    }
                }
            ]
        }
    }

    if "fedex_ground_economy" in active_services:
        fedex_payload["requestedShipment"]["smartPostDetail"] = {
            "indicia": "PARCEL_SELECT",
            "hubId": "5929"
        }

    # If you have specific SmartPost needs, you can still add them here if applicable, 
    # but the general Quote API usually returns standard services.
    f_response = get_fedex_rate(fedex_payload)
    if 'output' in f_response:
        rate_details = f_response['output'].get('rateReplyDetails', [])
        for detail in rate_details:
            f_code = detail.get('serviceType')
            
            # Find which service_key matches this FedEx code
            matched_key = None
            for skey, scodes in active_services.items():
                # Direct match from your map or logic-based match for Ground/Home
                if scodes.get("direct") == f_code:
                    matched_key = skey
                    break
                elif f_code == "GROUND_HOME_DELIVERY" and skey == "fedex_home_delivery":
                    matched_key = skey
                    break
                elif f_code == "FEDEX_GROUND" and skey == "fedex_ground":
                    matched_key = skey
                    break

            if matched_key:
                f_date_str = detail.get('operationalDetail', {}).get('deliveryDate')
                f_price = detail['ratedShipmentDetails'][0]['totalNetCharge']
                f_days = get_days_from_today(f_date_str)
                
                all_quotes.append({
                    "source": "Direct",
                    "service": matched_key,
                    "price": float(f_price),
                    "days": f_days
                })
                if matched_key in ["fedex_ground", "fedex_home_delivery"] and f_days:
                    base_days = f_days

    # --- 3. FINAL LOGIC & WINNER CALCULATION ---
    if not all_quotes:
        return 0.0, 0.0, 0.0, "", "No rates found for any service"
    
    quote_strings = []
    for q in all_quotes:
        # Your specific Economy logic
        if q['service'] == "fedex_ground_economy":
            if q['days'] is None or q['days'] == 0:
                q['days'] = base_days + 1

        d_label = f"{q['days']}d" if q['days'] is not None else "?d"
        quote_strings.append(f"{q['source']} {q['service']} ${q['price']:.2f} ({d_label})")

    # REMOVE RATES IF DAYS > 7 (OR IF DAYS ARE UNKNOWN)
    all_quotes = [q for q in all_quotes if q['days'] is not None and q['days'] <= 7]

    full_comparison_log = " | ".join(quote_strings)
    winner = min(all_quotes, key=lambda x: x["price"])

    return winner['price'], winner['service'], winner['days'], winner['source'], full_comparison_log
    
if __name__ == "__main__":
    test_order_no = "16-14269-67937"
    test_weight = 4
    test_dims = [10,8,6]
    winning_pkg_str = "Q3R"

    print(f"--- Starting FedEx Comparison Test for Order: {test_order_no} ---")

    try:
        # Note: We updated the return signature to (price, service_name, comparison_text)
        best_price, best_service, days, source, comparison_text = get_fedex_comparison_logic(
            test_order_no, 
            test_weight, 
            test_dims,
            winning_pkg_str
        )

        print("-" * 50)
        print(f"WINNER: {best_service}")
        print(f"PRICE:  ${best_price}")
        print(f"LOG:    {comparison_text}")
        print("-" * 50)

        print(f"Result: {source} {best_service} is the best value at ${best_price:.2f} arriving in {days} days.")

    except Exception as e:
        print(f"\nTest Failed: {str(e)}")
        import traceback
        traceback.print_exc()
    

    # shipment_details = {
    #     "accountNumber": {"value": os.getenv("FEDEX_ACCOUNT")},
    #     "rateRequestControlParameters": {"returnTransitTimes": True},
    #     "requestedShipment": {
    #         "preferredCurrency": "USD",
    #         "rateRequestType": ["ACCOUNT"],
    #         "shipper": {
    #             "address": {
    #                 "streetLines": ["3317 E 50th St"],
    #                 "city": "Vernon",
    #                 "stateOrProvinceCode": "CA",
    #                 "postalCode": "90058",
    #                 "countryCode": "US"
    #             }
    #         },
    #         "recipient": {
    #             "address": {
    #                 "streetLines": ["110 INTERNATIONALE BLVD", "HKLX48X"],
    #                 "city": "GLENDALE HEIGHTS",
    #                 "stateOrProvinceCode": "IL",
    #                 "postalCode": "60139",
    #                 "countryCode": "US",
    #                 "residential": False
    #             } 
    #         }, 
    #         "pickupType": "DROPOFF_AT_FEDEX_LOCATION",
    #         "serviceType":"FEDEX_GROUND",
    #         "packagingType": "YOUR_PACKAGING",
    #         "requestedPackageLineItems": [
    #             {
    #                 "weight": {"units": "LB", "value": 36},
    #                 "dimensions": {"length": 10, "width": 10, "height": 10, "units": "IN"}
    #             }
    #         ]
    #     }
    # }
    
    # print("Testing Rate API...")
    # result = get_fedex_rate(shipment_details)
    
    # if 'output' in result:
    #     details = result['output']['rateReplyDetails'][0]

    #     service_type = details['serviceType']
    #     delivery_date = details['commit']['dateDetail']['dayFormat']
        
    #     # Pull the charge from the first rated shipment detail
    #     net_charge = details['ratedShipmentDetails'][0]['totalNetCharge']
    #     currency = details['ratedShipmentDetails'][0]['currency']

    #     print("-" * 30)
    #     print(f"SERVICE: {service_type}")
    #     print(f"DELIVERY: {delivery_date}")
    #     print(f"RATE: {net_charge} {currency}")
    #     print("-" * 30)
    # else:
    #     print("Error in response:", result)