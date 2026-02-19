from src.shipstation.rates import get_live_rates
from datetime import datetime, timedelta,date
import pandas as pd
from src.shipping.engine import parse_dims
import config

def shop_and_optimize(order_no,weight, dims, to_state, to_zip, sku_info, store_id=None, is_residential=False):

    """
        Optimization Engine: Compares multiple carriers and packaging options

        This method performs a competitive 'rate shop' by:
        1. Calculating the 'Must Arrive By' date based on store's handling rules
        2. Testing different dimension sets (Primary, Alt, UPS Dims) against multiple carriers
        3. Filtering for delivery speed (via process_and_validate)
        4. Selecting the lowest-cost winner or falling back to Priority Mail if Ground is too slow.

        Returns:
            dict: The 'Winner' rate obj containing cost, service, and comparison logs
            None: If no valid rates are found or data is missing
    """

    # 254467 is 7001 account set it for 5 days because of top rated 1 day handling
    #days_offset = 5 if str(store_id) == "254467" else 7
    days_offset = 7
    max_delivery_date = date.today() + timedelta(days=days_offset)

    # Validate Weight Logic for non flat rates without SKU Weight
    try:
        if weight is None or pd.isna(weight) or float(weight) <= 0:
            is_weight_invalid = True
        else:
            is_weight_invalid = False
    except (TypeError, ValueError):
        is_weight_invalid = True

    if is_weight_invalid:
        current_pkg = str(sku_info.get("Package","")).upper()
        weight_independent_boxes = ["L", "M", "F", "P"]

        if current_pkg not in weight_independent_boxes:
            print(f" [!] Order {order_no} BLOCK: Package '{current_pkg}' requires weight, but weight is {weight}. Skipping optimizer.")
            return None
    
    # 1. Collect all dimension sets to test
    dim_sets = [("Primary", dims, sku_info.get("Package"))]
    
    # Add ALT PACKAGE if it exists
    alt_pkg = sku_info.get("ALT PACKAGE")
    if alt_pkg and not pd.isna(alt_pkg):
        alt_dims = parse_dims(alt_pkg)
        if alt_dims and alt_dims != dims: # Only add if different from primary
            dim_sets.append(("ALT", alt_dims, alt_pkg))
            
    # Add UPS DIMENSION if it exists
    ups_dim_pkg = sku_info.get("UPS DIMENSION")
    if ups_dim_pkg and not pd.isna(ups_dim_pkg):
        ups_dims = parse_dims(ups_dim_pkg)
        # Only add if different from primary and alt
        if ups_dims and ups_dims != dims and (len(dim_sets) < 2 or ups_dims != dim_sets[1][1]):
            dim_sets.append(("UPSD", ups_dims, ups_dim_pkg))

    all_raw_rates = []

    checked_priority_codes = set()

    # 2. Fetch rates for all dimension sets
    for label, d, pkg_str in dim_sets:
        print(f"--- Fetching rates for {label}: {d} ---")
        usps, verified_res = get_live_rates(order_no, "usps", "usps_ground_advantage", "package", weight, d, to_state, to_zip, is_residential)
        
        # After running get_live_rates on usps, it'll update the is_residential so that we can use it for ups
        is_residential = verified_res
        ups = []
        if is_residential and not (datetime.now().weekday() == 5 or (datetime.now().weekday() == 4 and datetime.now().hour >= 12)):
            #ups = get_live_rates(order_no, "ups", "ups_ground_saver", "package", weight, d, to_state, to_zip, is_residential)
            ups, _ = get_live_rates(order_no, "ups", None, "package", weight, d, to_state, to_zip, is_residential) # set it as none to get both ups_ground and ups_ground_saver

        print(f"{order_no} | [SHOP_AND_OPTIMIZE] DEBUG: UPS call returned {len(ups)} rates")

        priority_std_raw, _ = get_live_rates(order_no, "usps", "usps_priority_mail", "package", weight, d, to_state, to_zip, is_residential)

        priority_std = [
            r for r in priority_std_raw 
            if (r.get("packageType") or r.get("package_type")) in ["package", "parcel", None]
            and (r.get("serviceCode") or r.get("service_code")) == "usps_priority_mail"
        ]

        if priority_std:
            priority_std = [min(priority_std, key=lambda x: x.get("shipmentCost", 999))]

        # Check Priority Mail
        priority_check = []
        primary_pkg = [str(sku_info.get("Package","")).strip(), str(sku_info.get("ALT PACKAGE","")).strip()]
        
        for p_code in set(primary_pkg): # set() avoids duplicates
            if p_code in config.pkg_map and p_code not in checked_priority_codes:
                ss_code = config.pkg_map[p_code]
                # Pass None for dims when using specific Flat Rate package codes
                res, _ = get_live_rates(order_no, "usps", "usps_priority_mail", ss_code, weight, None, to_state, to_zip, is_residential)
                filtered_res = [r for r in res if (r.get("packageType") or r.get("package_type")) == ss_code]
                for fr_rate in filtered_res:
                    fr_rate["dim_source"] = f"FLAT_{p_code}"
                    fr_rate["winning_pkg_str"] = p_code 
                    priority_check.append(fr_rate)
                
                checked_priority_codes.add(p_code)

        # Add the label to the rate object so we know where it came from
        for r in (usps + ups + priority_std):
            r["dim_source"] = label

            serviceCode = r.get("serviceCode")

            if "ground" in serviceCode.lower() or "saver" in serviceCode.lower():
                r["winning_pkg_str"] = config.BOX_MAP.get(pkg_str.upper(),pkg_str)
            else:
                r["winning_pkg_str"] = pkg_str

            all_raw_rates.append(r)

        all_raw_rates.extend(priority_check)

    # Filter out the "0.0" and return only valid prices
    valid_raw = [r for r in all_raw_rates if r.get("shipmentCost",0) > 0]

    if not valid_raw:
        return None

    # 3. Process and validate all rates together
    valid_rates, comp_log = process_and_validate(valid_raw, max_delivery_date)

    # 4. FINAL DECISION
    if valid_rates:
        winner = min(valid_rates, key=lambda x: x["shipmentCost"])

        if " vs " not in comp_log:
            winner["comparison_log"] = f"{comp_log} ONLY"
        else:
            winner["comparison_log"] = comp_log
        print(f"  >>> WINNER: {winner['serviceName']} ({winner['winning_pkg_str']}) at ${winner['shipmentCost']}")
        return winner
    else:
        # FINAL FALLBACK: PRIORITY MAIL
        print("  [!] No Ground options met date. Falling back to Priority Mail...")
        fallback_pkg = str(sku_info.get("Package","")).strip()
        priority_raw, _ = get_live_rates(order_no, "usps", "usps_priority_mail", "package", weight, dims, to_state, to_zip, is_residential)
        priority = [
            r for r in priority_raw 
            if (r.get("packageType") or r.get("package_type")) in ["package", "parcel", None]
        ]

        if priority:
            winner = priority[0]

            base = winner.get("shipping_amount", {}).get("amount", 0.0)
            other = winner.get("other_amount", {}).get("amount", 0.0)
            winner["shipmentCost"] = round(base + other, 2)
            
            winner["carrierCode"] = "usps"
            winner["serviceCode"] = winner.get("serviceCode") or winner.get("service_code")
            winner["serviceName"] = winner.get("serviceName") or winner.get("service_type")
            winner["packageType"] = winner.get("packageType") or winner.get("package_type")
            winner["comparison_log"] = "ALL GROUND LATE [FINAL FALLBACK]"
            winner["is_priority_fallback"] = True
            winner["winning_pkg_str"] = fallback_pkg
            return winner
        return None

def process_and_validate(rates_list, max_date):
    """Helper to calculate costs and return valid rates + the comparison string."""
    valid = []
    comp_parts = []
    
    for r in rates_list:
        delivery_str = r.get("estimated_delivery_date")
        if not delivery_str:
            continue
            
        est_arrival = date.fromisoformat(delivery_str[:10])
        arrival_mm_dd = est_arrival.strftime("%m-%d")
        
        # 1. Use existing keys from rates.py if they exist
        cost = r.get("shipmentCost", 0.0)
        if cost is None: # Calculate if rates.py didn't do it
            base = r.get("shipping_amount", {}).get("amount", 0.0)
            other = r.get("other_amount", {}).get("amount", 0.0)
            cost = round(base + other, 2)

        s_name = (r.get("serviceName") or 
                  r.get("service_type") or 
                  r.get("service_code") or 
                  "UNKNOWN")
        s_code = r.get("serviceCode") or r.get("service_code")
        c_code = r.get("carrierCode") or r.get("carrier_code")
        source = r.get("dim_source", "")
        winner_pkg = r.get("winning_pkg_str")
        
        # 2. Re-inject to ensure all dictionaries have standard keys for main.py
        r["shipmentCost"] = cost
        r["serviceName"] = s_name
        r["serviceCode"] = s_code
        r["carrierCode"] = c_code
        
        comp_parts.append(f"{cost} [{s_name}-{winner_pkg};{arrival_mm_dd}]")
        
        if est_arrival <= max_date:
            valid.append(r)
            
    return valid, " vs ".join(comp_parts)