import pandas as pd
import re
from rdflib import Graph, Namespace, URIRef, Literal
from rdflib.namespace import RDF, RDFS, XSD, OWL
from datetime import datetime
import requests 

# Configuration files
mapping_path = "mapping1.xlsx"
instances_path = "instances.xlsx"
output_path = "output2.ttl"
BASE_NS = "http://example.org/"

GEONAMES_USERNAME = "th_iheb" 

# Dependency check
try:
    import geonamescache
    gc = geonamescache.GeonamesCache()
    if not gc.cities:
        print("--- WARNING: geonamescache loaded, but its city data is empty. ---")
except ImportError:
    print("--- WARNING: geonamescache not installed. ---")
    gc = None

# Utility functions

def make_safe_uri_label(value):
    if not isinstance(value, str):
        value = str(value)
    clean = value.strip()
    clean = re.sub(r"[^\w\s-]", "", clean)
    clean = re.sub(r"\s+", "_", clean)
    return clean

def parse_normalized_dates(date_str):
    if not isinstance(date_str, str): return None, None
    date_str = date_str.strip()
    if re.match(r"^\d{8}-\d{8}$", date_str):
        start, end = date_str.split("-", 1)
        return start.strip(), end.strip()
    elif re.match(r"^\d{4}-\d{4}$", date_str):
        start, end = date_str.split("-", 1)
        return start.strip(), end.strip()
    elif re.match(r"^\d{8}$", date_str): return date_str.strip(), None
    elif re.match(r"^\d{4}$", date_str): return date_str.strip(), None
    return None, None

def format_date_for_xsd(date_str):
    if not date_str: return None, None
    if len(date_str) == 8:
        try:
            dt_obj = datetime.strptime(date_str, "%Y%m%d")
            return dt_obj.strftime("%Y-%m-%d"), XSD.date
        except ValueError: return None, None
    elif len(date_str) == 4:
        try:
            int(date_str) 
            return date_str, XSD.gYear
        except ValueError: return None, None
    return None, None

def detect_object_term(obj_val_str, prefixes):
    if obj_val_str is None: return Literal("", datatype=XSD.string)
    s = str(obj_val_str).strip()
    if re.search(r"\bviaf\.org\/viaf\/\d+\b", s, re.IGNORECASE):
        if not s.lower().startswith(("http://", "https://")): s = "https://" + s.lstrip("/")
        return URIRef(s)
    if s.lower().startswith("http://") or s.lower().startswith("https://") or s.lower().startswith("www."):
        if s.lower().startswith("www."): s = "https://" + s
        return URIRef(s)
    if ":" in s and not s.lower().startswith("http"):
        prefix, local = s.split(":", 1)
        ns = prefixes.get(prefix)
        if ns is not None: return ns[local]
    return Literal(s, datatype=XSD.string)

def find_geonames_id_by_label(label):
    label_id = make_safe_uri_label(label).replace('_', '').lower()
    if gc and gc.cities:
        cities_data = gc.get_cities_by_name(label) 
        if isinstance(cities_data, dict): data_iterator = cities_data.values()
        elif isinstance(cities_data, list): data_iterator = [cities_data]
        else: data_iterator = []
        for city_list in data_iterator: 
            for city_data_wrapper in city_list: 
                if city_data_wrapper and isinstance(city_data_wrapper, dict):
                    city_details = next(iter(city_data_wrapper.values()))
                    if city_details['name'].lower() == label.lower():
                        return city_details['geonameid']
        for city_list in data_iterator: 
            for city_data_wrapper in city_list: 
                if city_data_wrapper and isinstance(city_data_wrapper, dict):
                    city_details = next(iter(city_data_wrapper.values()))
                    return city_details['geonameid']
    
    if GEONAMES_USERNAME == "your_username_here": return None

    try:
        url = "http://api.geonames.org/searchJSON"
        params = {'q': label, 'maxRows': 1, 'username': GEONAMES_USERNAME}
        response = requests.get(url, params=params)
        response.raise_for_status()
        data = response.json()
        if data and data.get('geonames'):
            return int(data['geonames'][0]['geonameId'])
    except Exception as e:
        print(f"  -> ERROR calling GeoNames API for '{label}': {e}")
    return None

def fetch_and_add_geonames_features(g, place_uri, geonames_id, place_label):
    if not geonames_id: return
    geonames_uri = URIRef(f"http://sws.geonames.org/{geonames_id}/")
    g.add((place_uri, OWL.sameAs, geonames_uri))
    lat, lon, fclass, fcode = None, None, None, None
    
    if gc and gc.cities:
        city_details = gc.cities.get(geonames_id)
        if city_details:
            lat = city_details.get('latitude')
            lon = city_details.get('longitude')
            fclass = city_details.get('feature_class')
            fcode = city_details.get('feature_code')

    if (not fclass or not lat) and GEONAMES_USERNAME != "your_username_here":
        try:
            url = "http://api.geonames.org/getJSON"
            params = {'geonameId': geonames_id, 'username': GEONAMES_USERNAME}
            response = requests.get(url, params=params)
            response.raise_for_status()
            data = response.json()
            if data:
                lat = data.get('lat')
                lon = data.get('lng')
                fclass = data.get('fcl')
                fcode = data.get('fcode')
        except Exception: pass

    if lat and lon:
        g.add((place_uri, WGS84["lat"], Literal(lat, datatype=XSD.decimal)))
        g.add((place_uri, WGS84["long"], Literal(lon, datatype=XSD.decimal)))
    if fclass: g.add((place_uri, NS_GN["featureClass"], Literal(fclass, datatype=XSD.string)))
    if fcode: g.add((place_uri, NS_GN["featureCode"], Literal(fcode, datatype=XSD.string)))

# Main loop

g = Graph()
prefixes = {"rdf": RDF, "rdfs": RDFS, "xsd": XSD, "owl": OWL}
for pfx, ns in prefixes.items(): g.bind(pfx, ns)
# Namespace for rico
prefixes["rico"] = Namespace(f"{BASE_NS}rico#")
g.bind("rico", prefixes["rico"])
ns_rico = prefixes["rico"]
# Namespace for place
prefixes["place"] = Namespace(f"{BASE_NS}place/")
g.bind("place", prefixes["place"]) 
# Namespace for date
prefixes["date"] = Namespace(f"{BASE_NS}date/")
g.bind("date", prefixes["date"]) 
# Namespace for Identifier
prefixes["identifier"] = Namespace(f"{BASE_NS}identifier/")
g.bind("identifier", prefixes["identifier"]) 
ns_ident = prefixes["identifier"]
# Namespace for Titles
prefixes["title"] = Namespace(f"{BASE_NS}title/")
g.bind("title", prefixes["title"]) 
ns_title = prefixes["title"]

# Namespace for agents
prefixes["person"] = Namespace(f"{BASE_NS}person/")
g.bind("person", prefixes["person"]) 
prefixes["agent"] = Namespace(f"{BASE_NS}agent/")
g.bind("agent", prefixes["agent"]) 
# Namespace for Geonames
NS_GN = Namespace("http://www.geonames.org/ontology#")
g.bind("gn", NS_GN)
WGS84 = Namespace("http://www.w3.org/2003/01/geo/wgs84_pos#")
g.bind("wgs84", WGS84)

def get_namespace(term):
    if ":" in term and not term.startswith("http"):
        prefix, _ = term.split(":", 1)
        if prefix not in prefixes:
            prefixes[prefix] = Namespace(f"{BASE_NS}{prefix}#")
            g.bind(prefix, prefixes[prefix])
        return prefixes[prefix]
    return None

# Load Excel
mapping_excel = pd.ExcelFile(mapping_path)
instances_excel = pd.ExcelFile(instances_path)

instances_dfs = {}
for name, df in instances_excel.parse(sheet_name=None).items():
    name_lower = name.strip().lower() 
    df.columns = df.columns.astype(str).str.strip().str.lower()
    instances_dfs[name_lower] = df 

# Constants
RICO_LOCATION_URI = ns_rico["isAssociatedWithPlace"]
RICO_ASSOC_PLACE_URI = ns_rico["isAssociatedWithPlace"]
RICO_HAS_TITLE_URI = ns_rico["hasOrHadTitle"]
RICO_TITLE_CLASS = ns_rico["Title"]
RICO_HAS_IDENTIFIER_URI = ns_rico["hasOrHadIdentifier"]
RICO_IDENTIFIER_CLASS = ns_rico["Identifier"]

RICO_DATE_PREDICATE = ns_rico["dateOrDateRange"]
RICO_DATE_DATATYPE = ns_rico["Date"]
RICO_HAS_BEGIN_DATE = ns_rico["hasBeginningDate"]
RICO_HAS_END_DATE = ns_rico["hasEndDate"]
RICO_HAS_CREATION_DATE = ns_rico["hasCreationDate"]
RICO_EXPRESSED_DATE = ns_rico["expressedDate"]
RICO_NORMALIZED_DATE = ns_rico["normalizedDateValue"]

# Predicates that typically link to a Person/Agent
# These trigger the logic for the "Fascicolo" sheet
PERSON_PREDICATES = {
    ns_rico["hasSender"],
    ns_rico["hasRecipient"],
    ns_rico["isAssociatedWith"], 
    ns_rico["personIsTargetOf"],
    ns_rico["hasAgent"],
    ns_rico["hasCreator"],
}

LITERAL_TO_ENTITY_PREDICATES = {
    ns_rico["isAssociatedWithPlace"],
    RICO_ASSOC_PLACE_URI, 
    RICO_HAS_BEGIN_DATE,
    RICO_HAS_END_DATE,
    RICO_HAS_CREATION_DATE,
}

STRUCTURAL_URI_PREDICATES = {
    ns_rico["isDirectlyIncludedIn"],
    ns_rico["isIncludedIn"]
}

# Processing loop

for mapping_sheet in mapping_excel.sheet_names:
    print(f"\n🔹 Processing mapping sheet: {mapping_sheet}")

    mapping_df = mapping_excel.parse(mapping_sheet)
    mapping_df.columns = mapping_df.columns.astype(str).str.strip()

    mapping_sheet_lower = mapping_sheet.strip().lower()
    instance_df = instances_dfs.get(mapping_sheet_lower)
    
    if instance_df is None:
        print(f" No matching instance sheet for '{mapping_sheet}', skipping.")
        continue

    required_cols = {"Subject", "Predicate", "Object", "Column Subject", "Column Object"}
    if not required_cols.issubset(mapping_df.columns):
        print(f" Missing required columns in sheet '{mapping_sheet}', skipping.")
        continue

    for col_name in ["Predicate", "Object"]:
        for val in mapping_df[col_name].dropna():
            val_str = str(val).strip()
            if ":" in val_str and not val_str.startswith("http"):
                get_namespace(val_str)

    for _, row in mapping_df.iterrows():
        subj_col = str(row["Column Subject"]).strip() if pd.notna(row["Column Subject"]) else None
        obj_col = str(row["Column Object"]).strip() if pd.notna(row["Column Object"]) else None
        predicate_str = str(row["Predicate"]).strip() if pd.notna(row["Predicate"]) else None
        base_object = str(row["Object"]).strip() if pd.notna(row["Object"]) else None

        if not subj_col or not predicate_str:
            continue

        ns_pred = get_namespace(predicate_str)
        pred = ns_pred[predicate_str.split(":", 1)[1]] if ns_pred else URIRef(predicate_str) 
        
        # SAFEGUARD: Skip explicit rdf:type mappings for Agents in Excel
        # This ensures the custom logic controls the entity creation
        if pred == RDF.type and (ns_rico["Person"] in str(base_object) or ns_rico["Agent"] in str(base_object)):
             continue

        for _, inst_row in instance_df.iterrows():
            subj_val = inst_row.get(subj_col.lower()) if subj_col else None
            obj_val = inst_row.get(obj_col.lower()) if obj_col else base_object
            
            if pd.isna(subj_val) or pd.isna(obj_val):
                continue

            subj_uri = URIRef(f"{BASE_NS}{make_safe_uri_label(subj_val)}")
            obj_val_str = str(obj_val).strip()
            
            final_obj_term = None 
            final_pred = pred 
            handled_custom = False

            
            # 1. LOGIC: SHEET "DOCUMENTO" (mittenti extra + viaf extra)
            
            if mapping_sheet_lower == "documento" and obj_col and "mittenti extra" in obj_col.lower():
                
                name_val = obj_val_str
                
                # Look for 'viaf extra' column
                viaf_code = None
                for c in inst_row.index:
                    if "viaf" in c and "extra" in c:
                        viaf_code = inst_row[c]
                        break
                
                external_link = None
                
                #  Check if both Name and VIAF exist
                is_person = False
                if pd.notna(viaf_code):
                    viaf_val_str = str(viaf_code).strip()
                    # Must not be empty or 'nan'
                    if viaf_val_str and viaf_val_str.lower() != "nan":
                        viaf_match = re.search(r"(\d+)$", viaf_val_str)
                        if viaf_match:
                             viaf_id = viaf_match.group(1)
                             external_link = URIRef(f"http://viaf.org/viaf/{viaf_id}/")
                             is_person = True

                if is_person:
                    entity_type = ns_rico["Person"]
                    entity_ns = prefixes["person"]
                    # print(f"   -> [Documento] Identified Person: {name_val}")
                else:
                    entity_type = ns_rico["Agent"]
                    entity_ns = prefixes["agent"]
                    # print(f"   -> [Documento] Identified Agent: {name_val}")

                safe_label = make_safe_uri_label(name_val)
                agent_uri = entity_ns[safe_label]
                
                if (agent_uri, RDF.type, entity_type) not in g:
                    g.add((agent_uri, RDF.type, entity_type))
                    g.add((agent_uri, RDFS.label, Literal(name_val)))
                    g.add((agent_uri, ns_rico["hasOrHadName"], Literal(name_val)))
                    if external_link:
                        g.add((agent_uri, OWL.sameAs, external_link))
                        
                final_obj_term = agent_uri
                handled_custom = True

            
            # 2. LOGIC: SHEET "FASCICOLO" (Predicates + Column "VIAF")
            
            elif mapping_sheet_lower == "fascicolo" and final_pred in PERSON_PREDICATES:
                name_val = obj_val_str
                
                # Look for column named specifically "VIAF" (or "link viaf")
                viaf_code = None
                for c in inst_row.index:
                    # Checks strict 'viaf' or 'link viaf' 
                    if c == "viaf" or c == "link viaf":
                        viaf_code = inst_row[c]
                        break
                
                external_link = None
                is_person = False

                # LOGIC: If VIAF column has data -> Person, else Agent
                if pd.notna(viaf_code):
                     viaf_val_str = str(viaf_code).strip()
                     if viaf_val_str and viaf_val_str.lower() != "nan":
                        viaf_match = re.search(r"(\d+)$", viaf_val_str)
                        if viaf_match:
                             viaf_id = viaf_match.group(1)
                             external_link = URIRef(f"http://viaf.org/viaf/{viaf_id}/")
                             is_person = True

                if is_person:
                    entity_type = ns_rico["Person"]
                    entity_ns = prefixes["person"]
                    # print(f"   -> [Fascicolo] Identified Person: {name_val}")
                else:
                    entity_type = ns_rico["Agent"]
                    entity_ns = prefixes["agent"]
                    # print(f"   -> [Fascicolo] Identified Agent: {name_val}")

                safe_label = make_safe_uri_label(name_val)
                agent_uri = entity_ns[safe_label]
                
                if (agent_uri, RDF.type, entity_type) not in g:
                    g.add((agent_uri, RDF.type, entity_type))
                    g.add((agent_uri, RDFS.label, Literal(name_val)))
                    g.add((agent_uri, ns_rico["hasOrHadName"], Literal(name_val)))
                    if external_link:
                        g.add((agent_uri, OWL.sameAs, external_link))
                
                final_obj_term = agent_uri
                handled_custom = True


            
            # 3. STANDARD LOGIC (Dates/Places)
            
            elif not handled_custom and final_pred in LITERAL_TO_ENTITY_PREDICATES:
                final_safe_label = make_safe_uri_label(obj_val_str) 
                entity_label = obj_val_str 
                
                if final_pred in {RICO_LOCATION_URI, RICO_ASSOC_PLACE_URI}:
                    entity_type_uri = ns_rico["Place"]
                    uri_path = "place"
                elif final_pred in {RICO_HAS_BEGIN_DATE, RICO_HAS_END_DATE, RICO_HAS_CREATION_DATE}:
                    entity_type_uri = ns_rico["Date"]
                    uri_path = "date"
                    start_str, end_str = parse_normalized_dates(obj_val_str)
                    date_uri_part = None
                    if final_pred == RICO_HAS_BEGIN_DATE: date_uri_part = start_str if end_str else None
                    elif final_pred == RICO_HAS_END_DATE: date_uri_part = end_str if end_str else None
                    elif final_pred == RICO_HAS_CREATION_DATE: date_uri_part = start_str if not end_str else None
                    if date_uri_part:
                        unique_date_string = date_uri_part
                        final_safe_label = make_safe_uri_label(unique_date_string)
                        entity_label = date_uri_part 
                    else:
                        if final_pred in {RICO_HAS_BEGIN_DATE, RICO_HAS_END_DATE, RICO_HAS_CREATION_DATE}:
                            continue 
                else: 
                    entity_type_uri = ns_rico["Agent"] 
                    uri_path = "agent" 
                
                entity_uri = URIRef(f"{BASE_NS}{uri_path}/{final_safe_label}") 
                if (entity_uri, RDF.type, entity_type_uri) not in g:
                    g.add((entity_uri, RDF.type, entity_type_uri)) 
                    g.add((entity_uri, RDFS.label, Literal(entity_label))) 
                    if entity_type_uri == ns_rico["Date"]:
                        value, datatype = format_date_for_xsd(entity_label)
                        if value:
                            g.add((entity_uri, RICO_NORMALIZED_DATE, Literal(value, datatype=datatype)))
                            g.add((subj_uri, RICO_NORMALIZED_DATE, Literal(value, datatype=datatype)))
                        if mapping_sheet_lower == "documento":
                            expressed_date_val = inst_row.get("data") 
                            if pd.notna(expressed_date_val):
                                g.add((entity_uri, RICO_EXPRESSED_DATE, Literal(str(expressed_date_val).strip(), datatype=XSD.string)))
                                g.add((subj_uri, RICO_EXPRESSED_DATE, Literal(str(expressed_date_val).strip(), datatype=XSD.string)))
                    elif entity_type_uri == ns_rico["Place"]:
                        geonames_id = find_geonames_id_by_label(str(entity_label))
                        fetch_and_add_geonames_features(g, entity_uri, geonames_id, str(entity_label))
                        if g.value(entity_uri, NS_GN.featureClass):
                             g.add((subj_uri, NS_GN.featureClass, g.value(entity_uri, NS_GN.featureClass)))
                        if g.value(entity_uri, NS_GN.featureCode):
                             g.add((subj_uri, NS_GN.featureCode, g.value(entity_uri, NS_GN.featureCode)))
                
                final_obj_term = entity_uri
                handled_custom = True

            
            # 4. RANGE LOGIC (Buste 22-29)
            
            elif not handled_custom and final_pred in {ns_rico["directlyIncludes"], ns_rico["includes"]}:
                
                # Split by semicolon to handle distinct groups (e.g. "Buste 1-5; Busta 10...")
                groups = obj_val_str.split(';')
                
                for group in groups:
                    group = group.strip()
                    if not group: continue

                    # Variable to remember the last Busta number seen in THIS group
                    # This allows "Busta 13, fascc. 1" to know it belongs to Busta 13
                    current_busta_num = None

                    # Regex: Finds "bust..." or "fasc..." followed by numbers
                    type_pattern = r"(?P<type>(?:bust|fasc)[a-z\.]*)\s*(?P<nums>[\d\s,\-\+]+)"
                    
                    matches = re.finditer(type_pattern, group, re.IGNORECASE)
                    
                    for match in matches:
                        type_str = match.group("type").lower()
                        nums_str = match.group("nums").strip()
                        
                        # Normalize separators (+ becomes ,)
                        nums_str = nums_str.replace('+', ',')
                        nums_str = nums_str.strip(', ')
                        
                        # Expand the numbers string into a list of integers
                        expanded_nums = []
                        parts = nums_str.split(',')
                        for part in parts:
                            part = part.strip()
                            if '-' in part:
                                try:
                                    s, e = part.split('-')
                                    expanded_nums.extend(range(int(s), int(e) + 1))
                                except ValueError: pass
                            elif part.isdigit():
                                expanded_nums.append(int(part))

                        # Process the expanded numbers based on type
                        if "bust" in type_str:
                            for num in expanded_nums:
                                # 1. Create the Busta URI
                                # ID Structure: Subject + _B + Number (e.g., S1_SS1_B13)
                                child_suffix = f"_B{num}"
                                child_uri_str = f"{subj_val}{child_suffix}"
                                
                                child_uri = URIRef(f"{BASE_NS}{make_safe_uri_label(child_uri_str)}")
                                g.add((subj_uri, final_pred, child_uri))
                                
                                # 2. Update context: Any subsequent fascicoli belong to this Busta
                                current_busta_num = num

                        elif "fasc" in type_str:
                            for num in expanded_nums:
                                # 1. Create the Fascicolo URI
                                if current_busta_num is not None:
                                    # ID Structure: Subject + _B{Busta} + _{FascPad}
                                    # Example: S1_SS1_B13_001
                                    child_suffix = f"_B{current_busta_num}_{num:03d}"
                                else:
                                    # Fallback if no Busta defined before (e.g., "Fascicoli 1-5" appearing alone)
                                    child_suffix = f"_F{num:03d}"
                                
                                child_uri_str = f"{subj_val}{child_suffix}"
                                
                                child_uri = URIRef(f"{BASE_NS}{make_safe_uri_label(child_uri_str)}")
                                g.add((subj_uri, final_pred, child_uri))

                handled_custom = True

            
            # 5. STRUCTURAL / TITLES / IDENTIFIERS / DEFAULT
            
            if not handled_custom:
                if final_pred in STRUCTURAL_URI_PREDICATES:
                    final_obj_term = URIRef(f"{BASE_NS}{make_safe_uri_label(obj_val_str)}")
                
                elif final_pred == RICO_HAS_IDENTIFIER_URI:
                    safe_id_label = make_safe_uri_label(obj_val_str)
                    id_uri = ns_ident[safe_id_label] 
                    if (id_uri, RDF.type, RICO_IDENTIFIER_CLASS) not in g:
                        g.add((id_uri, RDF.type, RICO_IDENTIFIER_CLASS))
                        g.add((id_uri, RDFS.label, Literal(obj_val_str))) 
                    final_obj_term = id_uri

                elif final_pred == RICO_HAS_TITLE_URI:
                    safe_title_label = make_safe_uri_label(obj_val_str)
                    title_uri = ns_title[safe_title_label] 
                    if (title_uri, RDF.type, RICO_TITLE_CLASS) not in g:
                        g.add((title_uri, RDF.type, RICO_TITLE_CLASS))
                        g.add((title_uri, RDFS.label, Literal(obj_val_str)))
                    final_obj_term = title_uri
                
                elif final_pred in {RICO_DATE_PREDICATE, RICO_EXPRESSED_DATE, RICO_NORMALIZED_DATE}:
                    continue
                
                else:
                    final_obj_term = detect_object_term(obj_val_str, prefixes)

            if final_obj_term:
                g.add((subj_uri, final_pred, final_obj_term))

print(f"\n✅ RDF graph built successfully.")
print(f"Total triples: {len(g)}")
g.serialize(destination=output_path, format="turtle")
print(f" Saved RDF graph to {output_path}")