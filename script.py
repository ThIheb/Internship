import pandas as pd
import re
from rdflib import Graph, Namespace, URIRef, Literal
from rdflib.namespace import RDF, RDFS, XSD, OWL
from datetime import datetime
import requests 

# Attempt to import geonamescache
try:
    import geonamescache
    gc = geonamescache.GeonamesCache()
except ImportError:
    print("Warning: geonamescache not installed. Using placeholder ID lookup.")
    gc = None

# --- Configuration ---
INPUT_FILE = "mapping1.xlsx"
OUTPUT_FILE = "output.ttl"
BASE_NS_STR = "http://example.org/"
BASE_NS = Namespace(BASE_NS_STR)

# --- Utility Functions (from script 2) ---

def make_safe_uri_label(value):
    """Cleans a string to be used as part of a URI."""
    if not isinstance(value, str):
        value = str(value)
    clean = value.strip()
    clean = re.sub(r"[^\w\s-]", "", clean) # Remove non-alphanumeric (keep spaces, _, -)
    clean = re.sub(r"\s+", "_", clean)     # Replace spaces with underscore
    return clean

def parse_normalized_dates(date_str):
    """Parses 'YYYYMMDD-YYYYMMDD' or 'YYYYMMDD'."""
    if not isinstance(date_str, str):
        return None, None
    date_str = date_str.strip()
    # Range: YYYYMMDD-YYYYMMDD
    if re.match(r"^\d{8}-\d{8}$", date_str):
        start, end = date_str.split("-", 1)
        return start.strip(), end.strip()
    # Single: YYYYMMDD
    elif re.match(r"^\d{8}$", date_str):
        return date_str.strip(), None
    return None, None

def normalize_to_xsd(date_str):
    """Converts 'YYYYMMDD' to 'YYYY-MM-DD'."""
    if not isinstance(date_str, str):
        date_str = str(date_str)
    if re.match(r"^\d{8}$", date_str):
        year, month, day = int(date_str[:4]), int(date_str[4:6]), int(date_str[6:8])
        try:
            # Validate date components
            datetime(year, month, day)
            return f"{year:04d}-{month:02d}-{day:02d}"
        except ValueError:
            return None # Invalid date (e.g., 20230230)
    return None

def detect_object_term(obj_val_str, prefixes_map):
    """Detects if a string is a date, URI (VIAF/URL), CURIE, or default literal."""
    if obj_val_str is None:
        return Literal("", datatype=XSD.string)
    s = str(obj_val_str).strip()

    # Skip date-like strings (handled by parse_normalized_dates)
    if re.match(r"^\d{8}(-\d{8})?$", s) or re.match(r"^\d{4}-\d{2}-\d{2}$", s):
        return Literal(s, datatype=XSD.string)

    # Detect VIAF or full URL
    if re.search(r"\bviaf\.org\/viaf\/\d+\b", s, re.IGNORECASE):
        if not s.lower().startswith(("http://", "https://")):
            s = "https://" + s.lstrip("/")
        return URIRef(s)

    if s.lower().startswith("http://") or s.lower().startswith("https://") or s.lower().startswith("www."):
        if s.lower().startswith("www."):
            s = "https://" + s
        return URIRef(s)

    # Detect CURIE (prefixed name)
    if ":" in s and not s.lower().startswith("http"):
        prefix, local = s.split(":", 1)
        ns = prefixes_map.get(prefix)
        if ns is not None:
            return ns[local] # Return a URIRef

    # Default: return as a Literal string
    return Literal(s, datatype=XSD.string)

def find_geonames_id_by_label(label):
    """Finds a Geonames ID for a given place label."""
    label_id = make_safe_uri_label(label).replace('_', '').lower()
    
    if gc:
        cities_data = gc.get_cities_by_name(label) 
        if isinstance(cities_data, dict):
            data_iterator = cities_data.values()
        elif isinstance(cities_data, list):
            data_iterator = [cities_data]
        else:
            data_iterator = []
        for city_list in data_iterator: 
            for city_data_wrapper in city_list:  
                if city_data_wrapper and isinstance(city_data_wrapper, dict):
                    city_details = next(iter(city_data_wrapper.values()))
                    return city_details['geonameid']
    
    # Fallback placeholders
    if 'bologna' in label_id: return 3176192 
    if 'roma' in label_id: return 3169070
    if 'parigi' in label_id or 'paris' in label_id: return 2988507
    if 'newyork' in label_id: return 5128581
    
    return None

def fetch_and_add_geonames_features(g, place_uri, geonames_id, place_label):
    """Adds owl:sameAs and other triples from Geonames."""
    if not geonames_id:
        return
        
    # 1. Add the owl:sameAs link
    geonames_uri = URIRef(f"http://sws.geonames.org/{geonames_id}/")
    g.add((place_uri, OWL.sameAs, geonames_uri))
    
    # 2. Hardcoded feature map (as in script 2)
    feature_map = {
        3176192: ('P', 'Populated Place', 388129), # Bologna, Italy
        3169070: ('P', 'City', 2872800),  # Rome
        2988507: ('P', 'Capital', 2140526),  # Paris
        5128581: ('P', 'Populated Place', 8804190),  # New York
    }
    
    feature_code, feature_label, population_val = feature_map.get(
        geonames_id, ('P', 'Populated Place', 100000) # Default
    )
    
    # 3. Add the GeoNames feature class/code triples
    g.add((place_uri, NS_GN["featureClass"], Literal(feature_code, datatype=XSD.string)))
    g.add((place_uri, NS_GN["featureCode"], Literal(feature_label, datatype=XSD.string)))
    
    print(f" -> Materialized features for {place_label} (ID: {geonames_id})")


# --- RDF Graph Initialization ---
g = Graph()

# Define standard namespaces
prefix_map = {
    "rdf": RDF, 
    "rdfs": RDFS, 
    "xsd": XSD, 
    "owl": OWL,
    "rico": Namespace("http://www.ica.org/standards/RiC/ontology#"),
    "ex": BASE_NS, # Base namespace for instances
    "place": Namespace(f"{BASE_NS_STR}place/"), # Namespace for Place instances
    "gn": Namespace("http://www.geonames.org/ontology#")
}

# Bind all namespaces to the graph
for pfx, ns in prefix_map.items():
    g.bind(pfx, ns)

# --- Prefix Handling (from script 1) ---

def extract_prefixes(df):
    """Finds all CURIE prefixes in mapping columns."""
    prefixes = set()
    # Check columns that define the triples
    for col in ["Subject", "Predicate", "Object"]:
        if col not in df.columns:
            continue
        for val in df[col].dropna():
            val = str(val).strip()
            if ":" in val and not val.startswith("http"):
                prefixes.add(val.split(":", 1)[0])
    return prefixes

# Load mapping file
try:
    xls = pd.ExcelFile(INPUT_FILE)
    sheet_names = xls.sheet_names
    print(f"Found sheets: {sheet_names}")
except FileNotFoundError:
    print(f"Error: {INPUT_FILE} not found.")
    exit()

all_prefixes = set()
for name in sheet_names:
    df = pd.read_excel(INPUT_FILE, sheet_name=name)
    all_prefixes |= extract_prefixes(df)

# Bind unknown prefixes dynamically
KNOWN_PREFIXES = set(prefix_map.keys())
for p in all_prefixes:
    if p not in KNOWN_PREFIXES:
        ns_uri = f"{BASE_NS_STR}{p}#"
        prefix_map[p] = Namespace(ns_uri)
        g.bind(p, prefix_map[p])

print(f"Auto-detected prefixes: {list(prefix_map.keys())}")

# --- Parsing Function (from script 1, adapted) ---

def parse_prefixed(value):
    """Parses a CURIE or creates a new URI in the base namespace."""
    value = str(value).strip()
    if ":" in value:
        prefix, local = value.split(":", 1)
        if prefix in prefix_map:
            return prefix_map[prefix][local]
    # Fallback: create URI in the base namespace
    return BASE_NS[value.replace(" ", "_")]

# --- Main Processing Logic (Merged) ---

# Define constants for special logic
ns_rico = prefix_map["rico"]
RICO_LOCATION_URI = ns_rico["hasOrHadLocation"]
RICO_PLACE_URI = ns_rico["Place"]
NS_GN = prefix_map["gn"]

LITERAL_TO_ENTITY_PREDICATES = {
    RICO_LOCATION_URI,
    ns_rico["hasSender"], 
    ns_rico["isAssociatedWith"], 
}
STRUCTURAL_URI_PREDICATES = {
    ns_rico["isDirectlyIncludedIn"],
    ns_rico["isIncludedIn"]
}

# Date-related predicates
has_begin = ns_rico["hasBeginningDate"]
has_end = ns_rico["hasEndDate"]
has_date = ns_rico["hasDate"]

# Parse each sheet
for sheet in sheet_names:
    print(f"\n🔹 Processing sheet: {sheet}")
    df = pd.read_excel(INPUT_FILE, sheet_name=sheet)

    # Process each row as a direct triple
    for _, row in df.iterrows():
        subj_val = row.get("Subject")
        pred_val = row.get("Predicate")
        obj_val = row.get("Object")

        if pd.isna(subj_val) or pd.isna(pred_val) or pd.isna(obj_val):
            continue

        # 1. Parse Subject and Predicate (standard way)
        subj_uri = parse_prefixed(subj_val)
        pred_uri = parse_prefixed(pred_val)
        
        # 2. Parse Object (with special logic)
        obj_val_str = str(obj_val).strip()
        final_obj_term = None

        # --- Special Handling (from script 2) ---

        # A. Date handling: Check if object *value* is a date range
        start_date, end_date = parse_normalized_dates(obj_val_str)
        if start_date:
            # Add *additional* triples for start/end
            if end_date:
                g.add((subj_uri, has_begin, Literal(normalize_to_xsd(start_date), datatype=XSD.date)))
                g.add((subj_uri, has_end, Literal(normalize_to_xsd(end_date), datatype=XSD.date)))
            else:
                g.add((subj_uri, has_date, Literal(normalize_to_xsd(start_date), datatype=XSD.date)))
            # The main triple will be added below as a string literal

        # B. Literal-to-URI: Check if *predicate* is a special type
        if pred_uri in LITERAL_TO_ENTITY_PREDICATES:
            safe_label = make_safe_uri_label(obj_val_str)
            
            if pred_uri == RICO_LOCATION_URI:
                entity_type_uri = ns_rico["Place"]
                entity_uri = prefix_map["place"][safe_label] # Use place: namespace
            else: 
                entity_type_uri = ns_rico["Agent"] 
                entity_uri = BASE_NS[f"entity/{safe_label}"] # Use ex:entity/
            
            # Define the new entity
            g.add((entity_uri, RDF.type, entity_type_uri)) 
            g.add((entity_uri, RDFS.label, Literal(obj_val_str))) 
            
            final_obj_term = entity_uri # The object is the new URI

        # C. Structural URI: Check if *predicate* is structural
        elif pred_uri in STRUCTURAL_URI_PREDICATES:
            # Force the object to be a URI in the base namespace
            final_obj_term = BASE_NS[make_safe_uri_label(obj_val_str)]
        
        # D. Default: Detect term type (URL, CURIE, or Literal)
        else:
            final_obj_term = detect_object_term(obj_val_str, prefix_map)

        # Add the Primary Triple
        g.add((subj_uri, pred_uri, final_obj_term))

# --- Post-processing: Geonames Materialization (from script 2) ---

print("\n--- Debugging Place Creation ---")
all_places = list(g.subjects(RDF.type, RICO_PLACE_URI))
print(f"  Total RICO:Place entities found in graph: {len(all_places)}")
if len(all_places) > 0:
    print(f"  Example Place URI: {all_places[0]}")
print("---------------------------------")

print("\nGeonames Feature Materialization...")
places_processed = 0

for place_uri in all_places: 
    place_label_triple = g.value(subject=place_uri, predicate=RDFS.label)
    if place_label_triple is not None:
        place_label = str(place_label_triple)
        geonames_id = find_geonames_id_by_label(place_label)
        if geonames_id:
            fetch_and_add_geonames_features(g, place_uri, geonames_id, place_label)
            places_processed += 1

print(f"✅ Materialization complete. Processed {places_processed} Place entities.")

# --- Serialize to Turtle File (from script 1) ---
print(f"\n✅ RDF graph built successfully.")
g.serialize(destination=OUTPUT_FILE, format="turtle")
print(f"  Total unique triples: {len(g)}")
print(f"  RDF graph saved to {OUTPUT_FILE}")