import pandas as pd
import re
import config

SKU_DB = pd.read_excel("data/DailyOutTools.xlsx", sheet_name=None)

def get_sku_info_from_dailyouttools(sku):
    """Lookup SKU specs from DailyOutTools.xlsx(Sheet: DB or Nonmount)"""
    for sheet in ["DB","Nonmounts"]:
        df = SKU_DB.get(sheet)
        if df is not None:
            match = df[df['SKU'] == sku]
            if not match.empty:
                return match.iloc[0].to_dict()
    return None

def parse_dims(dim_string):
    """Rule 5: Parse LxWxH format."""
    if not dim_string or pd.isna(dim_string):
        return None
    parts= re.split('[xX]',str(dim_string))
    if len(parts) == 3:
        try:
            return float(parts[0]), float(parts[1]), float(parts[2])
        except ValueError:
            return None
    return None

def get_weight_from_pkg_string(pkg_val):
    """
    Rule 3: Calculate weight based on the string found in the Package column.
    """
    pkg_val = str(pkg_val).upper().strip()
    
    if "1-4 OZ" in pkg_val:   return 0.25
    if "5-8 OZ" in pkg_val:   return 0.5
    if "9-12 OZ" in pkg_val:  return 0.75
    if "13-16 OZ" in pkg_val: return 0.99 # Rule 3 target
    return None

def get_carrier_service(row_data):
    """
    Core logic to determine Carrier/Service/Package based rules.

    This method applies business rules (Weight, Flat Rate, Q-codes) to decide
    if an order can be shipped via a fixed service or if it needs to be sent to the 
    shop_and_optimize method from optimizer.py for price optimization

    Returns: tuple: (Carrier, Service, PackageCode, Weight, Dimensions) OR ("ERROR", Message, None, 0)
    """
    current_sku = row_data.get('SKU','UNKNOWN')
    sku_info = get_sku_info_from_dailyouttools(row_data['SKU'])
    if not sku_info:
        return "ERROR", f"SKU {current_sku} NOT FOUND IN DB/NONMOUNT", None, 0

    try:
        # Get weight, default to 16 if missing or NaN
        raw_w = sku_info.get("Weight")
        db_weight = float(raw_w) if pd.notna(raw_w) else 16.0
    except (ValueError, TypeError):
        db_weight = 16.0
    
    # Fallback logic: Pkg > Alt Pkg > UPS Dimension
    pkg = sku_info.get("Package") or sku_info.get("ALT PACKAGE") or sku_info.get("UPS DIMENSION")
    if not pkg:
        return "ERROR",f"SKU {current_sku}: MISSING PKG/DIMENSIONS", None, 0
    
    pkg = str(pkg).strip()

    # Rule 2: USPS Priority Flat Rates (F, P, M, L)
    alt_pkg = sku_info.get("ALT PACKAGE")
    ups_dim = sku_info.get("UPS DIMENSION")

    # if alt or upsdim exists, we SHOP_RATES regardless of the primary pkg
    if(alt_pkg and pd.notna(alt_pkg)) or (ups_dim and pd.notna(ups_dim)):
        if pkg in config.DIM_MAP:
            primary_dims = config.DIM_MAP[pkg]
        else:
            # Fallback to parsing "10x10x10" string or default to (0,0,0)
            primary_dims = parse_dims(pkg) or (0, 0, 0)

        # 2. Return for SHOP_RATES
        return "SHOP_RATES", "OPTIMIZE", pkg, db_weight, primary_dims

    flat_rate_map = {
        'F': ('usps', 'usps_priority_mail', 'flat_rate_envelope'),
        'P': ('usps', 'usps_priority_mail', 'flat_rate_padded_envelope')
    }

    if pkg == 'M':
        return "SHOP_RATES","OPTIMIZE","M", db_weight, (12,9,6)
    
    if pkg == 'L':
        return "SHOP_RATES","OPTIMIZE","L", db_weight, (12,12,6)

    if pkg in flat_rate_map:
        return flat_rate_map[pkg] + (db_weight, None) # default weight for flat rate
    
    # Rule 3 : USPS First Class / Weight (oz)
    weight = get_weight_from_pkg_string(pkg)
    if weight is not None:
        return "usps","usps_first_class_mail",pkg, weight, None
    
    raw_weight = sku_info.get("Weight")
    if not raw_weight or pd.isna(raw_weight):
        return "ERROR", f"SKU {current_sku}: WEIGHT MISSING IN DB", None, 0
    
    try:
        db_weight = float(raw_weight)
    except ValueError:
        return "ERROR", f"SKU {current_sku}: INVALID WEIGHT '{raw_weight}'", None, 0
    
    # Rule 4 : Q-codes (standard boxes)
    if pkg in config.DIM_MAP:
        l, w, h = config.DIM_MAP[pkg]
        return "SHOP_RATES","OPTIMIZE",pkg,db_weight, (l,w,h)
    
    dims = parse_dims(pkg)
    if dims:
        return "SHOP_RATES", "OPTIMIZE", pkg, db_weight,dims