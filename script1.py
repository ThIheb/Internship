import pandas as pd
import re
from rdflib import Graph, Namespace, URIRef, Literal
from rdflib.namespace import RDF, RDFS, XSD, OWL
from datetime import datetime
import requests 
try:
    
    import geonamescache
    gc = geonamescache.GeonamesCache()
    if not gc.cities:
        print("--- WARNING: geonamescache loaded, but its city data is empty. ---")
except ImportError:
    
    print("--- WARNING: geonamescache not installed. ---")
    gc = None


GEONAMES_USERNAME = "th_iheb" 

# Utility Functions

def make_safe_uri_label(value):
    if not isinstance(value, str):
        value = str(value)
    clean = value.strip()
    clean = re.sub(r"[^\w\s-]", "", clean)
    clean = re.sub(r"\s+", "_", clean)
    return clean

def parse_normalized_dates(date_str):
    """Parses various date/range formats."""
    if not isinstance(date_str, str):
        return None, None
    date_str = date_str.strip()

    # YYYYMMDD-YYYYMMDD
    if re.match(r"^\d{8}-\d{8}$", date_str):
        start, end = date_str.split("-", 1)
        return start.strip(), end.strip()

    # YYYY-YYYY
    elif re.match(r"^\d{4}-\d{4}$", date_str):
        start, end = date_str.split("-", 1)
        return start.strip(), end.strip()

    # YYYYMMDD
    elif re.match(r"^\d{8}$", date_str):
        return date_str.strip(), None
    
    # YYYY
    elif re.match(r"^\d{4}$", date_str):
        return date_str.strip(), None

    return None, None

def format_date_for_xsd(date_str):
    """Formats YYYYMMDD or YYYY into a value and datatype tuple."""
    if not date_str: return None, None
    
    if len(date_str) == 8: # YYYYMMDD
        try:
            dt_obj = datetime.strptime(date_str, "%Y%m%d")
            return dt_obj.strftime("%Y-%m-%d"), XSD.date
        except ValueError:
            return None, None # Invalid YYYYMMDD
    
    elif len(date_str) == 4: # YYYY
        try:
            int(date_str) 
            return date_str, XSD.gYear # Return the year and gYear datatype
        except ValueError:
            return None, None
    
    return None, None


def detect_object_term(obj_val_str, prefixes):
    """Detects if a string is a URI (VIAF/URL), CURIE, or default literal."""
    if obj_val_str is None:
        return Literal("", datatype=XSD.string)
    s = str(obj_val_str).strip()

    # Detect VIAF or full URL
    if re.search(r"\bviaf\.org\/viaf\/\d+\b", s, re.IGNORECASE):
        if not s.lower().startswith(("http://", "https://")):
            s = "https://" + s.lstrip("/")
        return URIRef(s)

    if s.lower().startswith("http://") or s.lower().startswith("https://") or s.lower().startswith("www."):
        if s.lower().startswith("www."):
            s = "https://" + s
        return URIRef(s)

    
    if ":" in s and not s.lower().startswith("http"):
        prefix, local = s.split(":", 1)
        ns = prefixes.get(prefix)
        if ns is not None:
            return ns[local]

    return Literal(s, datatype=XSD.string)


# input and output files

mapping_path = "mapping1.xlsx"
instances_path = "instances.xlsx"
output_path = "output2.ttl"
BASE_NS = "http://example.org/"


# Initialize graph & namespaces


g = Graph()
prefixes = {"rdf": RDF, "rdfs": RDFS, "xsd": XSD, "owl": OWL}
for pfx, ns in prefixes.items():
    g.bind(pfx, ns)

prefixes["rico"] = Namespace(f"{BASE_NS}rico#")
g.bind("rico", prefixes["rico"])
ns_rico = prefixes["rico"]

# Namespace for Place instances
prefixes["place"] = Namespace(f"{BASE_NS}place/")
g.bind("place", prefixes["place"]) 

prefixes["date"] = Namespace(f"{BASE_NS}date/")
g.bind("date", prefixes["date"]) 

NS_GN = Namespace("http://www.geonames.org/ontology#")
g.bind("gn", NS_GN)

WGS84 = Namespace("http://www.w3.org/2003/01/geo/wgs84_pos#")
g.bind("wgs84", WGS84)

RICO_PLACE_URI = ns_rico["Place"]
RICO_HAS_IDENTIFIER_URI = ns_rico["hasOrHadIdentifier"]
RICO_IDENTIFIER_CLASS = ns_rico["Identifier"]


# Namespace for Identifier instances
prefixes["identifier"] = Namespace(f"{BASE_NS}identifier/")
g.bind("identifier", prefixes["identifier"]) 
ns_ident = prefixes["identifier"]


def get_namespace(term):
    if ":" in term and not term.startswith("http"):
        prefix, _ = term.split(":", 1)
        if prefix not in prefixes:
            prefixes[prefix] = Namespace(f"{BASE_NS}{prefix}#")
            g.bind(prefix, prefixes[prefix])
        return prefixes[prefix]
    return None



def find_geonames_id_by_label(label):
    label_id = make_safe_uri_label(label).replace('_', '').lower()
    
    # Try to find with geonamescache first
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
        
        # If no exact match, return the first one
        for city_list in data_iterator: 
            for city_data_wrapper in city_list: 
                if city_data_wrapper and isinstance(city_data_wrapper, dict):
                    city_details = next(iter(city_data_wrapper.values()))
                    return city_details['geonameid']

    # Fallback to live API if cache fails or is empty
    if GEONAMES_USERNAME == "your_username_here":
        print("--- WARNING: GeoNames username not set. Cannot use live API. ---")
        return None

    try:
        url = "http://api.geonames.org/searchJSON"
        params = {'q': label, 'maxRows': 1, 'username': GEONAMES_USERNAME}
        response = requests.get(url, params=params)
        response.raise_for_status() # Raise an error for bad responses
        data = response.json()
        if data and data.get('geonames'):
            geoname_id = data['geonames'][0]['geonameId']
            print(f" -> Found live API match for '{label}': {geoname_id}")
            return int(geoname_id)
    except Exception as e:
        print(f"  -> ERROR calling GeoNames API for '{label}': {e}")

    return None


def fetch_and_add_geonames_features(g, place_uri, geonames_id, place_label):
    if not geonames_id:
        return
        
    # Add owl:sameAs link TO THE PLACE
    geonames_uri = URIRef(f"http://sws.geonames.org/{geonames_id}/")
    g.add((place_uri, OWL.sameAs, geonames_uri))
    
    lat, lon, fclass, fcode = None, None, None, None
    
    # Try to get details from geonamescache
    if gc and gc.cities:
        city_details = gc.cities.get(geonames_id)
        if city_details:
            lat = city_details.get('latitude')
            lon = city_details.get('longitude')
            fclass = city_details.get('feature_class')
            fcode = city_details.get('feature_code')

    # If cache failed, use live API
    if not fclass or not lat:
        if GEONAMES_USERNAME == "your_username_here":
            print(f"  -> WARNING: No details for ID {geonames_id} and username not set. Skipping enrichment.")
            return
        try:
            url = "http://api.geonames.org/getJSON"
            params = {'geonameId': geonames_id, 'username': GEONAMES_USERNAME}
            response = requests.get(url, params=params)
            response.raise_for_status()
            data = response.json()
            if data:
                lat = data.get('lat')
                lon = data.get('lng')
                fclass = data.get('fcl') # Feature Class
                fcode = data.get('fcode') # Feature Code
                print(f" -> Fetched live details for '{place_label}'")
        except Exception as e:
            print(f"  -> ERROR calling GeoNames API for ID {geonames_id}: {e}")
            return

    # Add all found triples TO THE PLACE
    if lat and lon:
        lat_lit = Literal(lat, datatype=XSD.decimal)
        lon_lit = Literal(lon, datatype=XSD.decimal)
        g.add((place_uri, WGS84["lat"], lat_lit))
        g.add((place_uri, WGS84["long"], lon_lit))
    
    if fclass:
        fclass_lit = Literal(fclass, datatype=XSD.string)
        g.add((place_uri, NS_GN["featureClass"], fclass_lit))
    if fcode:
        fcode_lit = Literal(fcode, datatype=XSD.string)
        g.add((place_uri, NS_GN["featureCode"], fcode_lit))

    print(f" -> Materialized features (incl. lat/lon) for {place_label} (ID: {geonames_id})")


# excel inputs


mapping_excel = pd.ExcelFile(mapping_path)
instances_excel = pd.ExcelFile(instances_path)

# Normalize all instance column and sheet names to lowercase
instances_dfs = {}
for name, df in instances_excel.parse(sheet_name=None).items():
    name_lower = name.strip().lower() # Lowercase sheet name
    df.columns = df.columns.astype(str).str.strip().str.lower() # Lowercase column names
    instances_dfs[name_lower] = df 


# Process each mapping sheet separately


for mapping_sheet in mapping_excel.sheet_names:
    print(f"\n🔹 Processing mapping sheet: {mapping_sheet}")

    mapping_df = mapping_excel.parse(mapping_sheet)
    mapping_df.columns = mapping_df.columns.astype(str).str.strip()

    # Use lowercase for sheet matching
    mapping_sheet_lower = mapping_sheet.strip().lower()
    instance_df = instances_dfs.get(mapping_sheet_lower)
    
    if instance_df is None:
        print(f" No matching instance sheet for '{mapping_sheet}' (tried '{mapping_sheet_lower}'), skipping.")
        continue

    # Column names are now already normalized in instance_df

    required_cols = {"Subject", "Predicate", "Object", "Column Subject", "Column Object"}
    if not required_cols.issubset(mapping_df.columns):
        print(f" Missing required columns in sheet '{mapping_sheet}', skipping.")
        continue

    for col_name in ["Predicate", "Object"]:
        for val in mapping_df[col_name].dropna():
            val_str = str(val).strip()
            if ":" in val_str and not val_str.startswith("http"):
                get_namespace(val_str)

    # Process each mapping row
    for _, row in mapping_df.iterrows():
        subj_col = str(row["Column Subject"]).strip() if pd.notna(row["Column Subject"]) else None
        obj_col = str(row["Column Object"]).strip() if pd.notna(row["Column Object"]) else None
        predicate_str = str(row["Predicate"]).strip() if pd.notna(row["Predicate"]) else None
        base_object = str(row["Object"]).strip() if pd.notna(row["Object"]) else None

        if not subj_col or not predicate_str:
            continue


        ns_pred = get_namespace(predicate_str)
        pred = ns_pred[predicate_str.split(":", 1)[1]] if ns_pred else URIRef(predicate_str) 
        
        
        RICO_LOCATION_URI = ns_rico["isAssociatedWithPlace"]
        RICO_ASSOC_PLACE_URI = ns_rico["isAssociatedWithPlace"]
        RICO_HAS_TITLE_URI = ns_rico["hasOrHadTitle"]
        RICO_TITLE_DATATYPE = ns_rico["Title"]
        
        RICO_DATE_PREDICATE = ns_rico["dateOrDateRange"]
        RICO_DATE_DATATYPE = ns_rico["Date"]
        RICO_HAS_BEGIN_DATE = ns_rico["hasBeginningDate"]
        RICO_HAS_END_DATE = ns_rico["hasEndDate"]
        RICO_HAS_CREATION_DATE = ns_rico["hasCreationDate"]
        RICO_EXPRESSED_DATE = ns_rico["expressedDate"]
        RICO_NORMALIZED_DATE = ns_rico["normalizedDateValue"]
        
        
        LITERAL_TO_ENTITY_PREDICATES = {
            ns_rico["isAssociatedWithPlace"],
            RICO_ASSOC_PLACE_URI, 
            ns_rico["hasSender"], 
            ns_rico["isAssociatedWith"], 
            RICO_HAS_BEGIN_DATE,
            RICO_HAS_END_DATE,
            RICO_HAS_CREATION_DATE,
        }

        # Predicates that must convert a literal object to an existing structural URI
        STRUCTURAL_URI_PREDICATES = {
            ns_rico["isDirectlyIncludedIn"],
            ns_rico["isIncludedIn"]
        }


        for _, inst_row in instance_df.iterrows():
            # Use lowercase for column lookup
            subj_val = inst_row.get(subj_col.lower()) if subj_col else None
            obj_val = inst_row.get(obj_col.lower()) if obj_col else base_object
            
            if pd.isna(subj_val) or pd.isna(obj_val):
                continue

            subj_uri = URIRef(f"{BASE_NS}{make_safe_uri_label(subj_val)}")
            obj_val_str = str(obj_val).strip()
            
            final_obj_term = None 
            final_pred = pred 

            
            if final_pred in LITERAL_TO_ENTITY_PREDICATES:
                
                final_safe_label = make_safe_uri_label(obj_val_str) # Default label
                entity_label = obj_val_str # Default label
                
                # Determine URI path and Entity Type based on predicate
                if final_pred in {RICO_LOCATION_URI, RICO_ASSOC_PLACE_URI}:
                    entity_type_uri = ns_rico["Place"]
                    uri_path = "place"
                    # Create SHARED Place URI
                    final_safe_label = make_safe_uri_label(obj_val_str)
                    entity_label = obj_val_str 
                
                # Check if it's one of the date predicates
                elif final_pred in {RICO_HAS_BEGIN_DATE, RICO_HAS_END_DATE, RICO_HAS_CREATION_DATE}:
                    entity_type_uri = ns_rico["Date"]
                    uri_path = "date"
                    
                    start_str, end_str = parse_normalized_dates(obj_val_str)
                    
                    date_uri_part = None
                    
                    if final_pred == RICO_HAS_BEGIN_DATE:
                        if not end_str: # It's a single date
                            continue # SKIP
                        date_uri_part = start_str
                    
                    elif final_pred == RICO_HAS_END_DATE:
                        if not end_str: # It's a single date
                            continue # SKIP
                        date_uri_part = end_str
                    
                    elif final_pred == RICO_HAS_CREATION_DATE:
                        if end_str: # It's a range
                            continue # SKIP
                        date_uri_part = start_str
                    
                    if date_uri_part:
                        # CREATE A UNIQUE URI using Subject + Date Part
                        unique_date_string = f"{subj_val}_{date_uri_part}"
                        final_safe_label = make_safe_uri_label(unique_date_string)
                        entity_label = date_uri_part 
                    else:
                        if final_pred in {RICO_HAS_BEGIN_DATE, RICO_HAS_END_DATE, RICO_HAS_CREATION_DATE}:
                            continue 
                        print(f"  ⚠️ WARNING: Could not parse date '{obj_val_str}' for predicate {final_pred}. Using full string for URI.")
                        
                else: 
                    # Default for agents (sender/associated)
                    entity_type_uri = ns_rico["Agent"] 
                    uri_path = "entity"
                
                # Create the Entity URI from the final safe label
                entity_uri = URIRef(f"{BASE_NS}{uri_path}/{final_safe_label}") 
                
                # Define the Entity instance, but only if it's new
                if (entity_uri, RDF.type, entity_type_uri) not in g:
                    g.add((entity_uri, RDF.type, entity_type_uri)) 
                    g.add((entity_uri, RDFS.label, Literal(entity_label))) 
                    g.add((entity_uri, ns_rico["name"], Literal(entity_label, datatype=XSD.string)))
                    
                    if entity_type_uri == ns_rico["Date"]:
                        # Add normalizedDateValue TO THE DATE ENTITY
                        value, datatype = format_date_for_xsd(entity_label)
                        if value and datatype:
                            norm_literal = Literal(value, datatype=datatype)
                            g.add((entity_uri, RICO_NORMALIZED_DATE, norm_literal))
                            # Add shortcut TO THE RECORD
                            g.add((subj_uri, RICO_NORMALIZED_DATE, norm_literal))
                        else:
                            print(f"  ⚠️ WARNING: Could not normalize date label '{entity_label}' for entity {entity_uri}.")
                        
                        # Add expressed date TO THE DATE ENTITY
                        if mapping_sheet_lower == "documento":
                            expressed_date_val = inst_row.get("data") 
                            if pd.notna(expressed_date_val):
                                exp_literal = Literal(str(expressed_date_val).strip(), datatype=XSD.string)
                                g.add((entity_uri, RICO_EXPRESSED_DATE, exp_literal))
                                # Add shortcut TO THE RECORD
                                g.add((subj_uri, RICO_EXPRESSED_DATE, exp_literal))
                    
                    elif entity_type_uri == ns_rico["Place"]:
                        # Enrich the Place entity ONCE
                        place_label = str(entity_label)
                        geonames_id = find_geonames_id_by_label(place_label)
                        # This function only adds triples to entity_uri
                        fetch_and_add_geonames_features(g, entity_uri, geonames_id, place_label)

                
                # ADD SHORTCUTS TO RECORD EVERY TIME
                if entity_type_uri == ns_rico["Place"]:
                    # Query the place entity for its enrichment data
                    fclass_lit = g.value(entity_uri, NS_GN.featureClass)
                    fcode_lit = g.value(entity_uri, NS_GN.featureCode)
                    
                    # Add them as shortcuts to the record
                    if fclass_lit: g.add((subj_uri, NS_GN.featureClass, fclass_lit))
                    if fcode_lit: g.add((subj_uri, NS_GN.featureCode, fcode_lit))
                
                # Set the final object term to the newly created URI
                final_obj_term = entity_uri
            
            # Custom Logic: Structural URIs
            elif final_pred in STRUCTURAL_URI_PREDICATES:
                final_obj_term = URIRef(f"{BASE_NS}{make_safe_uri_label(obj_val_str)}")
                
                # Custom Logic: rico:hasOrHadIdentifier
            elif final_pred == RICO_HAS_IDENTIFIER_URI:
            
            # shared URI for this identifier
                safe_id_label = make_safe_uri_label(obj_val_str)
                id_uri = ns_ident[safe_id_label] 
            
            # Define the new Identifier entity
                if (id_uri, RDF.type, RICO_IDENTIFIER_CLASS) not in g:
                    g.add((id_uri, RDF.type, RICO_IDENTIFIER_CLASS))
                    # Add the literal value 
                    
                    g.add((id_uri, RDFS.label, Literal(f"Identifier: {obj_val_str}")))
            
            # Set the final object to be the URI of this new entity
                final_obj_term = id_uri
        
                
            # Custom Logic: rico:hasOrHadTitle
            elif final_pred == RICO_HAS_TITLE_URI:
                final_obj_term = Literal(obj_val_str, datatype=RICO_TITLE_DATATYPE)

            # Custom Logic: (SKIPS)
            elif final_pred == RICO_DATE_PREDICATE or \
                 final_pred == RICO_EXPRESSED_DATE or \
                 final_pred == RICO_NORMALIZED_DATE:
                continue 
            
            # Default/Fallback
            else:
                final_obj_term = detect_object_term(obj_val_str, prefixes)

            #Add the Primary Triple
            g.add((subj_uri, final_pred, final_obj_term))

            
            
            # Special logic for 'documento' sheet
            if mapping_sheet_lower == "documento" and isinstance(final_obj_term, URIRef):
                if "viaf.org/viaf/" in str(final_obj_term).lower():
                    has_sender_pred = ns_rico["hasSender"] 
                    g.add((subj_uri, has_sender_pred, final_obj_term))
                    
                    associated_with_pred = ns_rico["isAssociatedWith"]
                    g.add((subj_uri, associated_with_pred, final_obj_term))
                    


print("\n✅ Place enrichment is now handled during the main processing loop.")
# Serialize


print(f"\n✅ RDF graph built successfully.")
print(f"Total triples: {len(g)}")
g.serialize(destination=output_path, format="turtle")
print(f" Saved RDF graph to {output_path}")