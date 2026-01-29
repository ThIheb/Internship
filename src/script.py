import pandas as pd
import re
import os
import requests
from rdflib import Graph, Namespace, URIRef, Literal
from rdflib.namespace import RDF, RDFS, XSD, OWL
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(BASE_DIR)

mapping_path = os.path.join(PROJECT_ROOT, "data", "input", "mapping1.xlsx")
instances_path = os.path.join(PROJECT_ROOT, "data", "input", "instances.xlsx")
output_path = os.path.join(PROJECT_ROOT, "data", "output", "output2.ttl")
BASE_NS = "http://www.w3.org/ns/DigitalAlbini#"

GEONAMES_USERNAME = "th_iheb" 

try:
    import geonamescache
    gc = geonamescache.GeonamesCache()
    if not gc.cities:
        print("--- WARNING: geonamescache loaded, but its city data is empty. ---")
except ImportError:
    print("--- WARNING: geonamescache not installed. Falling back to API only. ---")
    gc = None

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
    if not label: return None
    clean_label = label.strip()
    
    if gc and gc.cities:
        cities_data = gc.get_cities_by_name(clean_label) 
        if cities_data:
            if isinstance(cities_data, dict): data_iterator = cities_data.values()
            elif isinstance(cities_data, list): data_iterator = [cities_data]
            else: data_iterator = []
            for city_list in data_iterator: 
                for city_data_wrapper in city_list: 
                    if city_data_wrapper and isinstance(city_data_wrapper, dict):
                        city_details = next(iter(city_data_wrapper.values()))
                        if city_details.get('name', '').lower() == clean_label.lower():
                            return city_details['geonameid']
    
    try:
        url = "http://api.geonames.org/searchJSON"
        params = {'q': clean_label, 'maxRows': 1, 'username': GEONAMES_USERNAME, 'style': 'FULL'}
        response = requests.get(url, params=params, timeout=5)
        response.raise_for_status()
        data = response.json()
        if data and data.get('geonames'):
            return int(data['geonames'][0]['geonameId'])
    except Exception as e:
        pass
    return None

def fetch_and_add_geonames_features(g, place_uri, geonames_id, place_label):
    if not geonames_id: return
    geonames_uri = URIRef(f"http://sws.geonames.org/{geonames_id}/")
    g.add((place_uri, OWL.sameAs, geonames_uri))
    
    lat, lon, fclass, fcode = None, None, None, None
    if gc and gc.cities:
        city_details = gc.cities.get(str(geonames_id))
        if city_details:
            lat = city_details.get('latitude')
            lon = city_details.get('longitude')
            fclass = city_details.get('feature_class')
            fcode = city_details.get('feature_code')

    if (not lat or not lon) and GEONAMES_USERNAME:
        try:
            url = "http://api.geonames.org/getJSON"
            params = {'geonameId': geonames_id, 'username': GEONAMES_USERNAME}
            response = requests.get(url, params=params, timeout=5)
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


g = Graph()
prefixes = {"rdf": RDF, "rdfs": RDFS, "xsd": XSD, "owl": OWL}
for pfx, ns in prefixes.items(): g.bind(pfx, ns)

prefixes["rico"] = Namespace("https://www.ica.org/standards/RiC/ontology#")
g.bind("rico", prefixes["rico"])
ns_rico = prefixes["rico"]

prefixes["temp"] = Namespace(f"{BASE_NS}temp/")
g.bind("temp", prefixes["temp"])

prefixes["storageid"] = Namespace(f"{BASE_NS}storageid/")
g.bind("storageid", prefixes["storageid"])
ns_storageid = prefixes["storageid"]

prefixes["type"] = Namespace(f"{BASE_NS}type/")
g.bind("type", prefixes["type"])
ns_type = prefixes["type"]

prefixes["corporateBody"] = Namespace(f"{BASE_NS}corporateBody/")
g.bind("corporateBody", prefixes["corporateBody"])

prefixes["record"] = Namespace(f"{BASE_NS}Record/")
g.bind("record", prefixes["record"])

prefixes["recordset"] = Namespace(f"{BASE_NS}RecordSet/")
g.bind("recordset", prefixes["recordset"])

prefixes["place"] = Namespace(f"{BASE_NS}place/")
g.bind("place", prefixes["place"]) 

prefixes["physloc"] = Namespace(f"{BASE_NS}physloc/")
g.bind("physloc", prefixes["physloc"])

prefixes["date"] = Namespace(f"{BASE_NS}date/")
g.bind("date", prefixes["date"])

prefixes["internalIdentifier"] = Namespace(f"{BASE_NS}internalIdentifier/")
g.bind("internalIdentifier", prefixes["internalIdentifier"]) 
ns_ident = prefixes["internalIdentifier"]
prefixes["identifier"] = prefixes["internalIdentifier"] 

prefixes["title"] = Namespace(f"{BASE_NS}title/")
g.bind("title", prefixes["title"]) 
ns_title = prefixes["title"]

prefixes["inst"] = Namespace(f"{BASE_NS}inst/")
g.bind("inst", prefixes["inst"]) 
ns_inst = prefixes["inst"]

prefixes["person"] = Namespace(f"{BASE_NS}person/")
g.bind("person", prefixes["person"]) 
prefixes["agent"] = Namespace(f"{BASE_NS}agent/")
g.bind("agent", prefixes["agent"]) 

NS_GN = Namespace("http://www.geonames.org/ontology#")
g.bind("gn", NS_GN)
WGS84 = Namespace("http://www.w3.org/2003/01/geo/wgs84_pos#")
g.bind("wgs84", WGS84)

def get_namespace(term):
    if ":" in term and not term.startswith("http"):
        prefix, _ = term.split(":", 1)
        if prefix not in prefixes:
            prefixes[prefix] = Namespace(f"{BASE_NS}{prefix}/")
            g.bind(prefix, prefixes[prefix])
        return prefixes[prefix]
    return None

mapping_excel = pd.ExcelFile(mapping_path)
instances_excel = pd.ExcelFile(instances_path)

instances_dfs = {}
for name, df in instances_excel.parse(sheet_name=None).items():
    name_lower = name.strip().lower() 
    df.columns = df.columns.astype(str).str.strip().str.lower()
    instances_dfs[name_lower] = df 

RICO_LOCATION_URI = ns_rico["isAssociatedWithPlace"]
RICO_ASSOC_PLACE_URI = ns_rico["isAssociatedWithPlace"]
RICO_HAS_TITLE_URI = ns_rico["hasOrHadTitle"]
RICO_TITLE_CLASS = ns_rico["Title"]
RICO_HAS_IDENTIFIER_URI = ns_rico["hasOrHadIdentifier"]
RICO_IDENTIFIER_CLASS = ns_rico["Identifier"]
RICO_HAS_INSTANTIATION = ns_rico["hasOrHadInstantiation"]
RICO_INSTANTIATION_CLASS = ns_rico["Instantiation"]

RICO_DATE_PREDICATE = prefixes["temp"]["dateProcessing"]
TEMP_BOX_ID_PREDICATE = prefixes["temp"]["boxIdentifier"]
TEMP_PROPAGATE_SENDER = prefixes["temp"]["propagateSender"]
TEMP_INTERMEDIATE_SENDER = prefixes["temp"]["intermediateSender"]

RICO_DATE_DATATYPE = ns_rico["Date"]
RICO_HAS_BEGIN_DATE = ns_rico["hasBeginningDate"]
RICO_HAS_END_DATE = ns_rico["hasEndDate"]
RICO_HAS_CREATION_DATE = ns_rico["hasCreationDate"]
RICO_EXPRESSED_DATE = ns_rico["expressedDate"]
RICO_NORMALIZED_DATE = ns_rico["normalizedDateValue"]

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
    RICO_DATE_PREDICATE 
}

STRUCTURAL_URI_PREDICATES = {
    ns_rico["isDirectlyIncludedIn"],
    ns_rico["isIncludedIn"],
    ns_rico["directlyIncludes"],
    ns_rico["includes"]
}


for mapping_sheet in mapping_excel.sheet_names:
    print(f"\n🔹 Processing mapping sheet: {mapping_sheet}")
    
    if "immagin" in mapping_sheet.strip().lower() or "image" in mapping_sheet.strip().lower():
        print(f"   ⚠️ Skipping '{mapping_sheet}' (Matched ignored list).")
        continue

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
        
        if pred == RDF.type:
             continue

        for _, inst_row in instance_df.iterrows():
            subj_val = inst_row.get(subj_col.lower()) if subj_col else None
            obj_val = inst_row.get(obj_col.lower()) if obj_col else base_object
            
            if pd.isna(subj_val) or pd.isna(obj_val):
                continue
            
            id_parts = len(str(subj_val).split('_'))

            if mapping_sheet_lower in ["serie", "sottoserie", "fascicolo", "fascicoli"]:
                current_prefix = prefixes["recordset"]
                rdf_class = ns_rico["RecordSet"]
            
            elif mapping_sheet_lower in ["documento", "documenti"]:
                if id_parts <= 4:
                    current_prefix = prefixes["recordset"]
                    rdf_class = ns_rico["RecordSet"]
                else:
                    current_prefix = prefixes["record"]
                    rdf_class = ns_rico["Record"]
            else:
                current_prefix = Namespace(BASE_NS)
                rdf_class = None

            subj_uri = current_prefix[make_safe_uri_label(subj_val)]
            if rdf_class and (subj_uri, RDF.type, rdf_class) not in g:
                g.add((subj_uri, RDF.type, rdf_class))

            
            if rdf_class == ns_rico["Record"]:
                safe_id_label = make_safe_uri_label(subj_val)
                id_uri = ns_ident[safe_id_label]
                g.add((subj_uri, ns_rico["hasOrHadIdentifier"], id_uri))
                g.add((id_uri, ns_rico["isOrWasIdentifierOf"], subj_uri))

                if (id_uri, RDF.type, RICO_IDENTIFIER_CLASS) not in g:
                    g.add((id_uri, RDF.type, RICO_IDENTIFIER_CLASS))
                    g.add((id_uri, RDFS.label, Literal(subj_val, datatype=XSD.string)))
                    g.add((id_uri, ns_rico["hasIdentifierType"], ns_type["internalIdentifier"]))
                    g.add((ns_type["internalIdentifier"], RDFS.label, Literal("InternalIdentifier", datatype=XSD.string)))
            

            obj_val_str = str(obj_val).strip()
            final_obj_term = None 
            final_pred = pred 
            handled_custom = False

            if final_pred == TEMP_PROPAGATE_SENDER:
                name_val = obj_val_str
                
                entity_type = ns_rico["Person"]
                entity_ns = prefixes["person"]
                is_institution = False

                busta_check = re.search(r"_B(\d+)(?:_|$)", str(subj_val), re.IGNORECASE)
                if busta_check:
                    b_num = int(busta_check.group(1))
                    if b_num >= 11:
                        entity_type = ns_rico["CorporateBody"]
                        entity_ns = prefixes["corporateBody"]
                        is_institution = True

                viaf_code = None
                for c in inst_row.index:
                    if c == "viaf" or c == "link viaf":
                        viaf_code = inst_row[c]
                        break
                
                external_link = None
                if pd.notna(viaf_code):
                      viaf_val_str = str(viaf_code).strip()
                      if viaf_val_str and viaf_val_str.lower() != "nan":
                        viaf_match = re.search(r"(\d+)$", viaf_val_str)
                        if viaf_match:
                             viaf_id = viaf_match.group(1)
                             external_link = URIRef(f"http://viaf.org/viaf/{viaf_id}/")

                safe_label = make_safe_uri_label(name_val)
                agent_uri = entity_ns[safe_label]
                
                if (agent_uri, RDF.type, entity_type) not in g:
                    g.add((agent_uri, RDF.type, entity_type))
                    g.add((agent_uri, ns_rico["hasOrHadName"], Literal(name_val)))
                    if external_link:
                        g.add((agent_uri, OWL.sameAs, external_link))
                
                g.add((subj_uri, TEMP_INTERMEDIATE_SENDER, agent_uri))
                handled_custom = True
            
            elif final_pred == TEMP_BOX_ID_PREDICATE:
                busta_match = re.search(r"_B(\d+)", obj_val_str, re.IGNORECASE)
                if not busta_match: busta_match = re.search(r"(?:box|busta|bust)\s*(\d+)", obj_val_str, re.IGNORECASE)
                if not busta_match: busta_match = re.search(r"(\d+)$", obj_val_str)

                if busta_match:
                    busta_num = busta_match.group(1)
                    box_label = f"Box {busta_num}"
                    safe_box_id = f"box{busta_num}"
                    
                    record_id_clean = make_safe_uri_label(subj_val)
                    inst_uri = ns_inst[record_id_clean]
                    
                    if (inst_uri, RDF.type, RICO_INSTANTIATION_CLASS) not in g:
                         g.add((inst_uri, RDF.type, RICO_INSTANTIATION_CLASS))
                         g.add((inst_uri, ns_rico["isOrWasInstantiationOf"], subj_uri))
                    
                    box_inst_uri = ns_inst[safe_box_id]
                    if (box_inst_uri, RDF.type, RICO_INSTANTIATION_CLASS) not in g:
                        g.add((box_inst_uri, RDF.type, RICO_INSTANTIATION_CLASS))
                        g.add((box_inst_uri, RDFS.label, Literal(f"Instantiation of {box_label}", datatype=XSD.string)))
                         
                    g.add((inst_uri, ns_rico["isOrWasPartOf"], box_inst_uri))
                    g.add((box_inst_uri, ns_rico["hasOrHadPart"], inst_uri))
                    
                    box_id_uri = ns_storageid[safe_box_id]
                    storage_type_uri = ns_type["StorageIdentifier"]
                    
                    g.add((storage_type_uri, RDF.type, ns_rico["IdentifierType"]))
                    g.add((storage_type_uri, RDFS.label, Literal("Storage Identifier", datatype=XSD.string)))
                    
                    g.add((box_inst_uri, ns_rico["hasOrHadIdentifier"], box_id_uri))
                    g.add((box_id_uri, ns_rico["isOrWasIdentifierOf"], box_inst_uri))
                    g.add((box_id_uri, ns_rico["hasIdentifierType"], storage_type_uri))

                    if (box_id_uri, RDF.type, RICO_IDENTIFIER_CLASS) not in g:
                        g.add((box_id_uri, RDF.type, RICO_IDENTIFIER_CLASS))
                        g.add((box_id_uri, RDFS.label, Literal(box_label, datatype=XSD.string)))

                else:
                    print(f"  [DEBUG] WARN: 'temp:boxIdentifier' regex failed on value: '{obj_val_str}'")
                handled_custom = True

            elif not handled_custom and final_pred in LITERAL_TO_ENTITY_PREDICATES:
                final_safe_label = make_safe_uri_label(obj_val_str) 
                entity_label = obj_val_str 
                
                if final_pred in {RICO_LOCATION_URI, RICO_ASSOC_PLACE_URI}:
                    place_uri = URIRef(f"{BASE_NS}place/{final_safe_label}")
                    physloc_uri = URIRef(f"{BASE_NS}physloc/{final_safe_label}")
                    
                    if (place_uri, RDF.type, ns_rico["Place"]) not in g:
                        g.add((place_uri, RDF.type, ns_rico["Place"]))
                        g.add((place_uri, RDFS.label, Literal(entity_label, datatype=XSD.string)))
                    
                    if (physloc_uri, RDF.type, ns_rico["PhysicalLocation"]) not in g:
                        g.add((physloc_uri, RDF.type, ns_rico["PhysicalLocation"]))
                        g.add((physloc_uri, RDFS.label, Literal(entity_label, datatype=XSD.string)))

                    g.add((place_uri, ns_rico["hasOrHadPhysicalLocation"], physloc_uri))
                    g.add((physloc_uri, ns_rico["isOrWasPhysicalLocationOf"], place_uri))

                    geonames_id = find_geonames_id_by_label(str(entity_label))
                    if geonames_id:
                        fetch_and_add_geonames_features(g, physloc_uri, geonames_id, str(entity_label))
                        
                        if g.value(physloc_uri, NS_GN.featureClass):
                             g.add((subj_uri, NS_GN.featureClass, g.value(physloc_uri, NS_GN.featureClass)))
                        if g.value(physloc_uri, NS_GN.featureCode):
                             g.add((subj_uri, NS_GN.featureCode, g.value(physloc_uri, NS_GN.featureCode)))

                    final_obj_term = place_uri
                    handled_custom = True

                elif final_pred in {RICO_HAS_BEGIN_DATE, RICO_HAS_END_DATE, RICO_HAS_CREATION_DATE, RICO_DATE_PREDICATE}:
                    entity_type_uri = ns_rico["Date"]
                    uri_path = "date"
                    start_str, end_str = parse_normalized_dates(obj_val_str)
                    
                    date_uri_part = None
                    if final_pred == RICO_HAS_BEGIN_DATE: date_uri_part = start_str if end_str else None
                    elif final_pred == RICO_HAS_END_DATE: date_uri_part = end_str if end_str else None
                    elif final_pred == RICO_HAS_CREATION_DATE: date_uri_part = start_str if not end_str else None
                    elif final_pred == RICO_DATE_PREDICATE: date_uri_part = start_str 

                    if date_uri_part:
                        final_safe_label = make_safe_uri_label(date_uri_part)
                        entity_label = date_uri_part 
                    else:
                        if final_pred in {RICO_HAS_BEGIN_DATE, RICO_HAS_END_DATE, RICO_HAS_CREATION_DATE}:
                            continue 
                        final_safe_label = make_safe_uri_label(obj_val_str)

                    entity_uri = URIRef(f"{BASE_NS}{uri_path}/{final_safe_label}") 
                    
                    if (entity_uri, RDF.type, entity_type_uri) not in g:
                        g.add((entity_uri, RDF.type, entity_type_uri)) 
                        if entity_type_uri == ns_rico["Date"]:
                            value, datatype = format_date_for_xsd(entity_label)
                            if value:
                                g.add((entity_uri, RICO_NORMALIZED_DATE, Literal(value, datatype=datatype)))
                                g.add((subj_uri, RICO_NORMALIZED_DATE, Literal(value, datatype=datatype)))
                            if mapping_sheet_lower == "documento":
                                expressed_date_val = inst_row.get("data") 
                                if pd.notna(expressed_date_val):
                                    val_str = str(expressed_date_val).strip()
                                    g.add((entity_uri, RICO_EXPRESSED_DATE, Literal(val_str, datatype=XSD.string)))
                                    g.add((subj_uri, RICO_EXPRESSED_DATE, Literal(val_str, datatype=XSD.string)))

                    if entity_type_uri == ns_rico["Date"]:
                        if final_pred == RICO_HAS_BEGIN_DATE:
                            g.add((entity_uri, ns_rico["isBeginningDateOf"], subj_uri))
                        elif final_pred == RICO_HAS_END_DATE:
                            g.add((entity_uri, ns_rico["isEndDateOf"], subj_uri))
                        elif final_pred in {RICO_HAS_CREATION_DATE, RICO_DATE_PREDICATE}:
                            g.add((entity_uri, ns_rico["isCreationDateOf"], subj_uri))

                    final_obj_term = entity_uri
                    handled_custom = True
                
                else: 
                    entity_type_uri = ns_rico["Agent"] 
                    uri_path = "agent"
                    entity_uri = URIRef(f"{BASE_NS}{uri_path}/{final_safe_label}")
                    
                    if (entity_uri, RDF.type, entity_type_uri) not in g:
                        g.add((entity_uri, RDF.type, entity_type_uri))

                    final_obj_term = entity_uri
                    handled_custom = True


            if not handled_custom:
                if final_pred in STRUCTURAL_URI_PREDICATES:
                    safe_obj_label = make_safe_uri_label(obj_val_str)
                    target_ns = Namespace(BASE_NS)
                    if mapping_sheet_lower == "documento": target_ns = prefixes["recordset"] 
                    elif mapping_sheet_lower == "fascicolo":
                        if final_pred in {ns_rico["includes"], ns_rico["directlyIncludes"]}: target_ns = prefixes["record"]
                        else: target_ns = prefixes["recordset"]
                    elif mapping_sheet_lower in ["serie", "sottoserie"]: target_ns = prefixes["recordset"]
                    final_obj_term = target_ns[safe_obj_label]
                
                elif final_pred == RICO_HAS_INSTANTIATION:
                    record_id_clean = make_safe_uri_label(subj_val)
                    inst_uri = ns_inst[record_id_clean]
                    if (inst_uri, RDF.type, RICO_INSTANTIATION_CLASS) not in g:
                        g.add((inst_uri, RDF.type, RICO_INSTANTIATION_CLASS))
                    g.add((inst_uri, ns_rico["isOrWasInstantiationOf"], subj_uri))
                    final_obj_term = inst_uri

                elif final_pred == RICO_HAS_IDENTIFIER_URI:
                    safe_id_label = make_safe_uri_label(obj_val_str)
                    id_uri = ns_ident[safe_id_label] 
                    if (id_uri, RDF.type, RICO_IDENTIFIER_CLASS) not in g:
                        g.add((id_uri, RDF.type, RICO_IDENTIFIER_CLASS))
                        internal_type_uri = ns_type["InternalIdentifier"]
                        g.add((internal_type_uri, RDF.type, ns_rico["IdentifierType"]))
                        g.add((internal_type_uri, RDFS.label, Literal("Internal Identifier", datatype=XSD.string)))
                        g.add((id_uri, ns_rico["hasIdentifierType"], internal_type_uri))
                    
                    g.add((id_uri, ns_rico["isOrWasIdentifierOf"], subj_uri))    
                    final_obj_term = id_uri

                elif final_pred == RICO_HAS_TITLE_URI:
                    safe_title_label = make_safe_uri_label(obj_val_str)
                    title_uri = ns_title[safe_title_label] 
                    if (title_uri, RDF.type, RICO_TITLE_CLASS) not in g:
                        g.add((title_uri, RDF.type, RICO_TITLE_CLASS))
                        g.add((title_uri, RDFS.label, Literal(obj_val_str, datatype=XSD.string)))

                    g.add((title_uri, ns_rico["isOrWasTitleOf"], subj_uri))
                    final_obj_term = title_uri
                
                elif final_pred in {RICO_DATE_PREDICATE, TEMP_BOX_ID_PREDICATE, TEMP_PROPAGATE_SENDER, RICO_EXPRESSED_DATE, RICO_NORMALIZED_DATE}:
                    continue
                
                else:
                    final_obj_term = detect_object_term(obj_val_str, prefixes)

            if final_obj_term:
                if final_pred == RDFS.label and not isinstance(final_obj_term, Literal): continue 
                if final_pred == RDF.type and isinstance(final_obj_term, Literal): continue 
                if final_pred in {RICO_DATE_PREDICATE, TEMP_BOX_ID_PREDICATE, TEMP_PROPAGATE_SENDER}: continue
                
                if str(final_pred).startswith(str(prefixes["temp"])):
                    continue
                
                val_check = str(final_obj_term).strip()
                if not val_check or val_check.lower() == "nan" or val_check == "Persona URI":
                    continue
                
                g.add((subj_uri, final_pred, final_obj_term))


print("\n🔹 Running Post-Processing: Propagating 'hasSender' from Fascicolo to Documents...")

triples_to_remove = []

for parent, _, sender in g.triples((None, TEMP_INTERMEDIATE_SENDER, None)):
    for _, _, child in g.triples((parent, ns_rico["includes"], None)): g.add((child, ns_rico["hasSender"], sender))
    for _, _, child in g.triples((parent, ns_rico["directlyIncludes"], None)): g.add((child, ns_rico["hasSender"], sender))
    for child, _, _ in g.triples((None, ns_rico["isIncludedIn"], parent)): g.add((child, ns_rico["hasSender"], sender))
    for child, _, _ in g.triples((None, ns_rico["isDirectlyIncludedIn"], parent)): g.add((child, ns_rico["hasSender"], sender))
    triples_to_remove.append((parent, TEMP_INTERMEDIATE_SENDER, sender))

for t in triples_to_remove: g.remove(t)

print(f"   -> Propagated senders to children and cleaned up temp links.")

print(f"\n✅ RDF graph built successfully.")
print(f"Total triples: {len(g)}")
g.serialize(destination=output_path, format="turtle")
print(f" Saved RDF graph to {output_path}")