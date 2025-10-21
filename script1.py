import pandas as pd
import re
from rdflib import Graph, Namespace, URIRef, Literal
from rdflib.namespace import RDF, RDFS, XSD, OWL


# Utility functions to make uri's safe

def make_safe_uri_label(value):
    if not isinstance(value, str):
        value = str(value)
    clean = value.strip()
    clean = re.sub(r"[^\w\s-]", "", clean)  # remove punctuation
    clean = re.sub(r"\s+", "_", clean)      # spaces -> underscores
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
    if re.match(r"^\d{8}$", date_str):
        return f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"
    return date_str


# Configuration

mapping_path = "mapping1.xlsx"
instances_path = "instances.xlsx"
output_path = "output2.ttl"
num_sheets_to_process = 4
BASE_NS = "http://example.org/"


# Initialize RDF graph

g = Graph()
prefixes = {}

g.bind("rdf", RDF)
g.bind("rdfs", RDFS)
g.bind("xsd", XSD)
g.bind("owl", OWL)
prefixes.update({"rdf": RDF, "rdfs": RDFS, "xsd": XSD, "owl": OWL})

def get_namespace(term):
    if ":" in term and not term.startswith("http"):
        prefix, _ = term.split(":", 1)
        if prefix not in prefixes:
            prefixes[prefix] = Namespace(f"{BASE_NS}{prefix}#")
            g.bind(prefix, prefixes[prefix])
        return prefixes[prefix]
    return None


# Read Excel files (instances files limits the amount of sheets to be read)

mapping_excel = pd.ExcelFile(mapping_path)
instances_excel = pd.ExcelFile(instances_path)

instances_dfs = {name: instances_excel.parse(name) for name in instances_excel.sheet_names}
for df in instances_dfs.values():
    df.columns = df.columns.astype(str).str.strip()

if "rico" not in prefixes:
    prefixes["rico"] = Namespace(f"{BASE_NS}rico#")
    g.bind("rico", prefixes["rico"])
ns_rico = prefixes["rico"]


# Process mapping sheets

for sheet_name in mapping_excel.sheet_names[:num_sheets_to_process]:
    print(f"\nProcessing mapping sheet: {sheet_name}")
    mapping_df = mapping_excel.parse(sheet_name)
    mapping_df.columns = mapping_df.columns.astype(str).str.strip()

    required_cols = {"Subject", "Predicate", "Object", "Column Subject", "Column Object"}
    if not required_cols.issubset(mapping_df.columns):
        print(f"Missing required columns in sheet '{sheet_name}', skipping.")
        continue

    # Bind namespaces from mapping
    for col_name in ["Predicate", "Object"]:
        for val in mapping_df[col_name].dropna():
            val_str = str(val).strip()
            if ":" in val_str and not val_str.startswith("http"):
                get_namespace(val_str)

    
    # Row-level mapping
    
    for _, row in mapping_df.iterrows():
        subj_col = str(row["Column Subject"]).strip() if pd.notna(row["Column Subject"]) else None
        obj_col = str(row["Column Object"]).strip() if pd.notna(row["Column Object"]) else None
        predicate_str = str(row["Predicate"]).strip() if pd.notna(row["Predicate"]) else None
        base_object = str(row["Object"]).strip() if pd.notna(row["Object"]) else None

        if not subj_col or not predicate_str:
            continue

        # Skip "estremi cronologici" column (extracting only from "data")
        if subj_col.lower() == "estremi cronologici" or (obj_col and obj_col.lower() == "estremi cronologici"):
            continue

        ns_pred = get_namespace(predicate_str)
        pred = ns_pred[predicate_str.split(":", 1)[1]] if ns_pred else URIRef(predicate_str)

        for inst_name, instances_df in instances_dfs.items():
            if subj_col not in instances_df.columns:
                print(f"Column missing in instances sheet '{inst_name}': {subj_col}")
                continue

            for _, inst_row in instances_df.iterrows():
                subj_val = inst_row.get(subj_col)
                obj_val = inst_row.get(obj_col) if obj_col else base_object

                if pd.isna(subj_val) or pd.isna(obj_val) or str(subj_val).strip() == "":
                    continue

                safe_label = make_safe_uri_label(subj_val)
                subj_uri = URIRef(f"{BASE_NS}{safe_label}")
                obj_val_str = str(obj_val).strip()

                #Date handling
                start_date, end_date = parse_normalized_dates(obj_val_str)
                has_begin = ns_rico["hasBeginningDate"]
                has_end = ns_rico["hasEndDate"]
                has_date = ns_rico["hasDate"]
                normalized_pred = ns_rico["normalizedDateValue"]

                if start_date or end_date:
                    if end_date:
                        if (subj_uri, has_begin, Literal(normalize_to_xsd(start_date), datatype=XSD.date)) not in g:
                            g.add((subj_uri, has_begin, Literal(normalize_to_xsd(start_date), datatype=XSD.date)))
                        if (subj_uri, has_end, Literal(normalize_to_xsd(end_date), datatype=XSD.date)) not in g:
                            g.add((subj_uri, has_end, Literal(normalize_to_xsd(end_date), datatype=XSD.date)))
                    else:
                        if (subj_uri, has_date, Literal(normalize_to_xsd(start_date), datatype=XSD.date)) not in g:
                            g.add((subj_uri, has_date, Literal(normalize_to_xsd(start_date), datatype=XSD.date)))
                    continue  # skip normal object handling for dates

                # Only add normalizedDateValue if explicitly mapped
                if predicate_str == "rico:normalizedDateValue":
                    if (subj_uri, normalized_pred, Literal(obj_val_str)) not in g:
                        g.add((subj_uri, normalized_pred, Literal(obj_val_str)))
                    continue

                # Normal object handling
                if ":" in obj_val_str and not obj_val_str.startswith("http"):
                    prefix, local = obj_val_str.split(":", 1)
                    ns_obj = prefixes.get(prefix)
                    obj_term = ns_obj[local] if ns_obj else URIRef(obj_val_str)
                else:
                    obj_term = Literal(obj_val_str)

                if (subj_uri, pred, obj_term) not in g:
                    g.add((subj_uri, pred, obj_term))

# Output serialization
print("\nRDF graph built successfully")
print(f"Total triples: {len(g)}")
g.serialize(destination=output_path, format="turtle")
print(f"Saved RDF graph to {output_path}")
