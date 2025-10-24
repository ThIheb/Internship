import pandas as pd
import re
from rdflib import Graph, Namespace, URIRef, Literal
from rdflib.namespace import RDF, RDFS, XSD, OWL
from datetime import datetime


# Utility Functions


def make_safe_uri_label(value):
    if not isinstance(value, str):
        value = str(value)
    clean = value.strip()
    clean = re.sub(r"[^\w\s-]", "", clean)
    clean = re.sub(r"\s+", "_", clean)
    return clean

def parse_normalized_dates(date_str):
    if not isinstance(date_str, str):
        return None, None
    date_str = date_str.strip()
    if re.match(r"^\d{8}-\d{8}$", date_str):
        start, end = date_str.split("-", 1)
        return start.strip(), end.strip()
    elif re.match(r"^\d{8}$", date_str):
        return date_str.strip(), None
    return None, None

def normalize_to_xsd(date_str):
    if not isinstance(date_str, str):
        date_str = str(date_str)
    if re.match(r"^\d{8}$", date_str):
        year, month, day = int(date_str[:4]), int(date_str[4:6]), int(date_str[6:8])
        try:
            datetime(year, month, day)
            return f"{year:04d}-{month:02d}-{day:02d}"
        except ValueError:
            return None
    return None

def detect_object_term(obj_val_str, prefixes):
    """Detects if a string is a date, URI (VIAF/URL), CURIE, or default literal."""
    if obj_val_str is None:
        return Literal("", datatype=XSD.string)
    s = str(obj_val_str).strip()

    # Skip date-like strings
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

def get_namespace(term):
    if ":" in term and not term.startswith("http"):
        prefix, _ = term.split(":", 1)
        if prefix not in prefixes:
            prefixes[prefix] = Namespace(f"{BASE_NS}{prefix}#")
            g.bind(prefix, prefixes[prefix])
        return prefixes[prefix]
    return None


# excel inputs


mapping_excel = pd.ExcelFile(mapping_path)
instances_excel = pd.ExcelFile(instances_path)

# Normalize column names
instances_dfs = {}
for name, df in instances_excel.parse(sheet_name=None).items():
    df.columns = df.columns.astype(str).str.strip() 
    instances_dfs[name] = df 


# Process each mapping sheet separately


for mapping_sheet in mapping_excel.sheet_names:
    print(f"\n🔹 Processing mapping sheet: {mapping_sheet}")

    mapping_df = mapping_excel.parse(mapping_sheet)
    mapping_df.columns = mapping_df.columns.astype(str).str.strip()

    instance_df = instances_dfs.get(mapping_sheet)
    if instance_df is None:
        print(f" No matching instance sheet for '{mapping_sheet}', skipping.")
        continue

    instance_df.columns = instance_df.columns.astype(str).str.strip()

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

        if subj_col.lower() == "estremi cronologici" or (obj_col and obj_col.lower() == "estremi cronologici"):
            continue

        ns_pred = get_namespace(predicate_str)
        pred = ns_pred[predicate_str.split(":", 1)[1]] if ns_pred else URIRef(predicate_str) 
        
        
        RICO_LOCATION_URI = ns_rico["hasOrHadLocation"]
        
        
        LITERAL_TO_ENTITY_PREDICATES = {
            ns_rico["hasOrHadLocation"],
            ns_rico["hasSender"], 
            ns_rico["isAssociatedWith"], 
            # any other predicates should be added here
        }

        # Predicates that must convert a literal object to an existing structural URI
        STRUCTURAL_URI_PREDICATES = {
            ns_rico["isDirectlyIncludedIn"],
            ns_rico["isIncludedIn"]
        }


        for _, inst_row in instance_df.iterrows():
            subj_val = inst_row.get(subj_col)
            obj_val = inst_row.get(obj_col) if obj_col else base_object
            if pd.isna(subj_val) or pd.isna(obj_val):
                continue

            subj_uri = URIRef(f"{BASE_NS}{make_safe_uri_label(subj_val)}")
            obj_val_str = str(obj_val).strip()
            
            final_obj_term = None 
            final_pred = pred 

            # Date handling 
            start_date, end_date = parse_normalized_dates(obj_val_str)
            has_begin = ns_rico["hasBeginningDate"]
            has_end = ns_rico["hasEndDate"]
            has_date = ns_rico["hasDate"]
            normalized_pred = ns_rico["normalizedDateValue"]

            if start_date or end_date:
                # Omitted for brevity
                if end_date:
                    g.add((subj_uri, has_begin, Literal(normalize_to_xsd(start_date), datatype=XSD.date)))
                    g.add((subj_uri, has_end, Literal(normalize_to_xsd(end_date), datatype=XSD.date)))
                else:
                    g.add((subj_uri, has_date, Literal(normalize_to_xsd(start_date), datatype=XSD.date)))
                continue

            if predicate_str == "rico:normalizedDateValue":
                g.add((subj_uri, normalized_pred, Literal(obj_val_str)))
                continue

            # Custom Logic: Literal-to-URI 
            if final_pred in LITERAL_TO_ENTITY_PREDICATES:
                
                safe_label = make_safe_uri_label(obj_val_str)
                
                # Determine URI path and Entity Type based on predicate
                if final_pred == RICO_LOCATION_URI:
                    entity_type_uri = ns_rico["Place"]
                    uri_path = "place"
                else: 
                    # Default for agents (sender/associated)
                    entity_type_uri = ns_rico["Agent"] 
                    uri_path = "entity"
                
                # Create the Entity URI from the literal value
                entity_uri = URIRef(f"{BASE_NS}{uri_path}/{safe_label}") 
                
                # Define the Entity instance (Subject, Type, Label)
                g.add((entity_uri, RDF.type, entity_type_uri)) 
                g.add((entity_uri, RDFS.label, Literal(obj_val_str))) 
                
                # Set the final object term to the newly created URI
                final_obj_term = entity_uri
            
            
            elif final_pred in STRUCTURAL_URI_PREDICATES:
                # Force the object to a URIRef based on BASE_NS
                final_obj_term = URIRef(f"{BASE_NS}{make_safe_uri_label(obj_val_str)}")
                
            
            else:
                final_obj_term = detect_object_term(obj_val_str, prefixes)

            #Add the Primary Triple
            g.add((subj_uri, final_pred, final_obj_term))

            
            
            
            if mapping_sheet == "documento" and isinstance(final_obj_term, URIRef):
                if "viaf.org/viaf/" in str(final_obj_term).lower():
                    has_sender_pred = ns_rico["hasSender"] 
                    g.add((subj_uri, has_sender_pred, final_obj_term))
                    
                    associated_with_pred = ns_rico["isAssociatedWith"]
                    g.add((subj_uri, associated_with_pred, final_obj_term))
            


# Serialize


print(f"\n✅ RDF graph built successfully.")
print(f"Total triples: {len(g)}")
g.serialize(destination=output_path, format="turtle")
print(f" Saved RDF graph to {output_path}")