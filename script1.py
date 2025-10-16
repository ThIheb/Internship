import pandas as pd
from rdflib import Graph, Namespace, URIRef, Literal
from rdflib.namespace import RDF, RDFS, XSD, OWL

# variables for configuration
mapping_path = "mapping1.xlsx"       
instances_path = "instances.xlsx"   
output_path = "output2.ttl"
num_sheets_to_process = 3  # number of excel sheets to map
BASE_NS = "http://example.org/"

# create rdf graph
g = Graph()
prefixes = {}

# standard namespaces and prefixes
g.bind("rdf", RDF)
g.bind("rdfs", RDFS)
g.bind("xsd", XSD)
g.bind("owl", OWL)
prefixes["rdf"] = RDF
prefixes["rdfs"] = RDFS
prefixes["xsd"] = XSD
prefixes["owl"] = OWL

# -auto-bind found namespaces
def get_namespace(term):
    if ":" in term and not term.startswith("http"):
        prefix, _ = term.split(":", 1)
        if prefix not in prefixes:
            prefixes[prefix] = Namespace(f"{BASE_NS}{prefix}#")
            g.bind(prefix, prefixes[prefix])
        return prefixes[prefix]
    return None

# read excel files
mapping_excel = pd.ExcelFile(mapping_path)
instances_excel = pd.ExcelFile(instances_path)

# Clean column headers for all instance sheets
instances_dfs = {name: instances_excel.parse(name) for name in instances_excel.sheet_names}
for name, df in instances_dfs.items():
    df.columns = df.columns.astype(str).str.strip()

# ensure rico namespace exists
if "rico" not in prefixes:
    prefixes["rico"] = Namespace(f"{BASE_NS}rico#")
    g.bind("rico", prefixes["rico"])
ns_rico = prefixes["rico"]

# --- Process first 3 mapping sheets ---
for sheet_name in mapping_excel.sheet_names[:num_sheets_to_process]:
    print(f"\n Processing mapping sheet: {sheet_name}")
    mapping_df = mapping_excel.parse(sheet_name)
    mapping_df.columns = mapping_df.columns.astype(str).str.strip()

    required_cols = {"Subject", "Predicate", "Object", "Column Subject", "Column Object"}
    if not required_cols.issubset(mapping_df.columns):
        print(f" Missing required columns in sheet '{sheet_name}', skipping.")
        continue

    # detect namespaces for the columns predicate and object
    for col_name in ["Predicate", "Object"]:
        for val in mapping_df[col_name].dropna():
            val_str = str(val).strip()
            if ":" in val_str and not val_str.startswith("http"):
                get_namespace(val_str)  # binds the prefix automatically

    for _, row in mapping_df.iterrows():
        subj_col = str(row["Column Subject"]).strip() if pd.notna(row["Column Subject"]) else None
        obj_col = str(row["Column Object"]).strip() if pd.notna(row["Column Object"]) else None
        predicate_str = str(row["Predicate"]).strip() if pd.notna(row["Predicate"]) else None
        base_object = str(row["Object"]).strip() if pd.notna(row["Object"]) else None

        if not subj_col or not predicate_str:
            continue

        # Create predicate URI using auto-bound namespaces
        ns_pred = get_namespace(predicate_str)
        pred = ns_pred[predicate_str.split(":", 1)[1]] if ns_pred else URIRef(predicate_str)

        # Iterate all instance sheets
        for inst_name, instances_df in instances_dfs.items():
            if subj_col not in instances_df.columns:
                print(f" Column missing in instances sheet '{inst_name}': {subj_col}")
                continue

            for _, inst_row in instances_df.iterrows():
                subj_val = inst_row.get(subj_col)
                if obj_col and obj_col in instances_df.columns:
                    obj_val = inst_row.get(obj_col)
                else:
                    obj_val = base_object  

                # if instance object is empty
                if pd.isna(obj_val) or obj_val == "":
                    obj_val = base_object

                if pd.isna(subj_val) or pd.isna(obj_val):
                    continue

                subj_uri = URIRef(f"{BASE_NS}{subj_val}")
                obj_val_str = str(obj_val).strip()

                # handling for spliting the start and end dates
                if ((obj_col and obj_col.lower() == "estremi cronologici") or
                    (subj_col and subj_col.lower() == "estremi cronologici")) and "-" in obj_val_str:
                    start, end = [d.strip() for d in obj_val_str.split("-", 1)]

                    # Force predicates to be rico:
                    has_begin = ns_rico["hasBeginningDate"]
                    has_end = ns_rico["hasEndDate"]

                    start_lit = Literal(start)
                    end_lit = Literal(end)

                    if (subj_uri, has_begin, start_lit) not in g:
                        g.add((subj_uri, has_begin, start_lit))
                    if (subj_uri, has_end, end_lit) not in g:
                        g.add((subj_uri, has_end, end_lit))

                    continue  

                
                if ":" in obj_val_str and not obj_val_str.startswith("http"):
                    prefix, local = obj_val_str.split(":", 1)
                    ns_obj = prefixes.get(prefix)
                    if ns_obj:
                        obj_term = ns_obj[local]
                    else:
                        obj_term = URIRef(obj_val_str)
                else:
                    obj_term = Literal(obj_val_str)

                if (subj_uri, pred, obj_term) not in g:
                    g.add((subj_uri, pred, obj_term))

print("\n RDF graph built successfully")
print(f"Total triples: {len(g)}")

# serialize to ttl file
g.serialize(destination=output_path, format="turtle")
print(f" Saved RDF graph to {output_path}")
