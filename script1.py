import pandas as pd
import re
from rdflib import Graph, Namespace, URIRef, Literal
from rdflib.namespace import RDF, RDFS, XSD, OWL



def make_safe_uri_label(value):
    if not isinstance(value, str):
        value = str(value)
    clean = value.strip()
    clean = re.sub(r"[^\w\s-]", "", clean)
    clean = re.sub(r"\s+", "_", clean)
    return clean

def detect_object_term(obj_val_str, prefixes):
    """Detects if a string is a URI, CURIE, or default literal."""
    if obj_val_str is None:
        return Literal("", datatype=XSD.string)
    s = str(obj_val_str).strip()

    if s.lower().startswith("http://") or s.lower().startswith("https://"):
        return URIRef(s)
    
    if ":" in s and not s.lower().startswith("http"):
        prefix, local = s.split(":", 1)
        ns = prefixes.get(prefix)
        if ns is not None:
            return ns[local]

    return Literal(s, datatype=XSD.string)



mapping_path = "mapping1.xlsx"
output_path = "output.ttl"
BASE_NS = "http://example.org/" 



g = Graph()
prefixes = {"rdf": RDF, "rdfs": RDFS, "xsd": XSD, "owl": OWL}
for pfx, ns in prefixes.items():
    g.bind(pfx, ns)

# Namespaces definition
prefixes["rico"] = Namespace(f"{BASE_NS}rico#")
g.bind("rico", prefixes["rico"])
ns_rico = prefixes["rico"]

prefixes["place"] = Namespace(f"{BASE_NS}place/")
g.bind("place", prefixes["place"]) 
ns_place = prefixes["place"]

prefixes["date"] = Namespace(f"{BASE_NS}date/")
g.bind("date", prefixes["date"]) 
ns_date = prefixes["date"]

prefixes["identifier"] = Namespace(f"{BASE_NS}identifier/")
g.bind("identifier", prefixes["identifier"]) 
ns_ident = prefixes["identifier"]

prefixes["title"] = Namespace(f"{BASE_NS}title/")
g.bind("title", prefixes["title"]) 
ns_title = prefixes["title"]

prefixes["appellation"] = Namespace(f"{BASE_NS}appellation/")
g.bind("appellation", prefixes["appellation"]) 
ns_appell = prefixes["appellation"]

def get_namespace(term):
    """Dynamically adds new prefixes if found."""
    if ":" in term and not term.startswith("http"):
        prefix, _ = term.split(":", 1)
        if prefix not in prefixes:
            print(f" -> Auto-defining new prefix '{prefix}'")
            prefixes[prefix] = Namespace(f"{BASE_NS}{prefix}#")
            g.bind(prefix, prefixes[prefix])
        return prefixes[prefix]
    return None



print(f"Loading mapping file from: {mapping_path}")
mapping_excel = pd.ExcelFile(mapping_path)

for mapping_sheet in mapping_excel.sheet_names:
    print(f"🔹 Processing structure for sheet: {mapping_sheet}")
    
    mapping_df = mapping_excel.parse(mapping_sheet)
    mapping_df.columns = mapping_df.columns.astype(str).str.strip()

    # Pre-scan for any new namespaces
    for col_name in ["Predicate", "Object"]:
        for val in mapping_df[col_name].dropna():
            get_namespace(str(val).strip())
            
    #RICO Variables
    RICO_LOCATION_URI = ns_rico["isAssociatedWithPlace"]
    RICO_ASSOC_PLACE_URI = ns_rico["isAssociatedWithPlace"]
    
    RICO_HAS_TITLE_URI = ns_rico["hasOrHadTitle"]
    RICO_HAS_APPELLATION_URI = ns_rico["hasOrHadAppellation"]
    RICO_HAS_IDENTIFIER_URI = ns_rico["hasOrHadIdentifier"]
    
    RICO_HAS_BEGIN_DATE = ns_rico["hasBeginningDate"]
    RICO_HAS_END_DATE = ns_rico["hasEndDate"]
    RICO_HAS_CREATION_DATE = ns_rico["hasCreationDate"]
    
    RICO_DATE_PREDICATE = ns_rico["dateOrDateRange"]
    RICO_EXPRESSED_DATE = ns_rico["expressedDate"]
    RICO_NORMALIZED_DATE = ns_rico["normalizedDateValue"]
    
    STRUCTURAL_URI_PREDICATES = {
        ns_rico["isDirectlyIncludedIn"],
        ns_rico["isIncludedIn"]
    }

    # Process each mapping row
    for _, row in mapping_df.iterrows():
        subj_col = str(row["Column Subject"]).strip() if pd.notna(row["Column Subject"]) else None
        obj_col = str(row["Column Object"]).strip() if pd.notna(row["Column Object"]) else None
        predicate_str = str(row["Predicate"]).strip() if pd.notna(row["Predicate"]) else None
        base_object = str(row["Object"]).strip() if pd.notna(row["Object"]) else None

        if not subj_col or not predicate_str:
            continue
        
        if not obj_col and not base_object:
            continue

        
        # Create a generic Subject URI based on its column name
        safe_subj_col = make_safe_uri_label(subj_col)
        subj_uri = URIRef(f"{BASE_NS}{safe_subj_col}_VALUE")

        # --- Predicate ---
        ns_pred = get_namespace(predicate_str)
        pred = ns_pred[predicate_str.split(":", 1)[1]] if ns_pred else URIRef(predicate_str) 
        
        final_obj_term = None

        if base_object:
            # The object is a fixed value 
            final_obj_term = detect_object_term(base_object, prefixes)
            
        elif obj_col:
            # The object comes from a column. 
            
            safe_obj_col = make_safe_uri_label(obj_col)

            if pred == RICO_HAS_IDENTIFIER_URI:
                final_obj_term = ns_ident[safe_obj_col]
            
            elif pred == RICO_HAS_APPELLATION_URI:
                final_obj_term = ns_appell[safe_obj_col]
                
            elif pred == RICO_HAS_TITLE_URI:
                final_obj_term = ns_title[safe_obj_col]

            elif pred in {RICO_LOCATION_URI, RICO_ASSOC_PLACE_URI}:
                final_obj_term = ns_place[safe_obj_col]
            
            elif pred == RICO_HAS_BEGIN_DATE:
                final_obj_term = ns_date[f"{safe_subj_col}_START_DATE"]
                
            elif pred == RICO_HAS_END_DATE:
                final_obj_term = ns_date[f"{safe_subj_col}_END_DATE"]
                
            elif pred == RICO_HAS_CREATION_DATE:
                final_obj_term = ns_date[f"{safe_subj_col}_CREATION_DATE"]

            elif pred in STRUCTURAL_URI_PREDICATES:
                final_obj_term = URIRef(f"{BASE_NS}{safe_obj_col}_VALUE")

            elif pred in {RICO_DATE_PREDICATE, RICO_EXPRESSED_DATE, RICO_NORMALIZED_DATE}:
                continue # Skip these helper predicates

            else:
                
                final_obj_term = Literal(f"{obj_col}_VALUE", datatype=XSD.string)
        
       
        if final_obj_term:
            
            if (subj_uri, RDF.type, None) not in g and (subj_uri, RDFS.label, None) not in g:
                 g.add((subj_uri, RDFS.label, Literal(f"instance based on '{subj_col}'")))
                 
            g.add((subj_uri, pred, final_obj_term))

# --- Serialize Output ---

print(f"\n✅ Mapping structure graph built successfully.")
print(f"Total triples: {len(g)}")
g.serialize(destination=output_path, format="turtle")
print(f" Saved mapping structure to {output_path}")