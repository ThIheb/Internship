import pandas as pd
from rdflib import Graph, Namespace, URIRef, Literal
from rdflib.namespace import RDF, RDFS, XSD, OWL

# --- Configuration ---
mapping_path = "mapping1.xlsx"
output_path = "output.ttl"
BASE_NS = "http://example.org/"

# --- Utility Functions ---
def make_safe_label(value):
    return str(value).strip().replace(" ", "_").replace(":", "_").upper()

# --- Graph Setup ---
g = Graph()
prefixes = {"rdf": RDF, "rdfs": RDFS, "xsd": XSD, "owl": OWL}
for pfx, ns in prefixes.items(): g.bind(pfx, ns)

# Define Namespaces
ns_rico = Namespace(f"{BASE_NS}rico#")
g.bind("rico", ns_rico)

ns_structure = Namespace(f"{BASE_NS}structure/")
g.bind("structure", ns_structure)

# Data Namespaces (for visualization purposes)
prefixes_dict = {
    "place": Namespace(f"{BASE_NS}place/"),
    "date": Namespace(f"{BASE_NS}date/"),
    "identifier": Namespace(f"{BASE_NS}identifier/"),
    "person": Namespace(f"{BASE_NS}person/"),
    "agent": Namespace(f"{BASE_NS}agent/"),
    "inst": Namespace(f"{BASE_NS}inst/"),
    "storageid": Namespace(f"{BASE_NS}storageid/"),
    "institution": Namespace(f"{BASE_NS}institution/"),
    "record": Namespace(f"{BASE_NS}Record/"),
    "recordset": Namespace(f"{BASE_NS}RecordSet/")
}
for k, v in prefixes_dict.items(): g.bind(k, v)

# Triggers
TEMP_BOX_ID = "temp:boxIdentifier"
TEMP_SENDER = "temp:hasSender"
TEMP_PROPAGATE = "temp:propagateSender"
TEMP_DATE = "temp:dateProcessing"

# --- Main Logic ---

print(f"Loading mapping rules from: {mapping_path}")
mapping_excel = pd.ExcelFile(mapping_path)

for sheet in mapping_excel.sheet_names:
    # 1. Filter excluded sheets
    if "immagin" in sheet.lower() or "image" in sheet.lower():
        continue

    print(f"🔹 Modeling structure for: {sheet}")
    mapping_df = mapping_excel.parse(sheet)
    
    # 2. Define the Generic Subject for this Sheet
    # e.g., structure:GENERIC_FASCICOLO
    sheet_clean = make_safe_label(sheet)
    subj_node = ns_structure[f"GENERIC_{sheet_clean}"]
    
    # Add a label so we know what this node represents
    g.add((subj_node, RDFS.label, Literal(f"Generic Representative of '{sheet}'")))

    # 3. Apply Class Logic (Record vs RecordSet)
    # This reflects the custom logic in the main script
    if sheet.lower() in ["serie", "sottoserie", "fascicolo", "fascicoli"]:
        g.add((subj_node, RDF.type, ns_rico["RecordSet"]))
        g.add((subj_node, RDFS.comment, Literal("Mapped as rico:RecordSet via Custom Logic")))
    elif sheet.lower() in ["documento", "documenti"]:
        g.add((subj_node, RDF.type, ns_rico["Record"]))
        g.add((subj_node, RDFS.comment, Literal("Mapped as rico:Record via Custom Logic")))

    # 4. Iterate Rules
    for _, row in mapping_df.iterrows():
        pred_str = str(row["Predicate"]).strip()
        obj_col = str(row["Column Object"]).strip()
        
        # --- A. CUSTOM LOGIC: BOX IDENTIFIER ---
        if pred_str == TEMP_BOX_ID:
            # Simulate the Instantiation -> Identifier -> Box structure
            inst_node = ns_structure[f"GENERIC_{sheet_clean}_INSTANTIATION"]
            box_node = prefixes_dict["storageid"]["EXAMPLE_BOX_ID"]
            type_node = ns_rico["IdentifierType"]

            # Subject -> Instantiation
            g.add((inst_node, RDF.type, ns_rico["Instantiation"]))
            g.add((inst_node, ns_rico["isOrWasInstantiationOf"], subj_node))
            
            # Instantiation -> Box Identifier
            g.add((inst_node, ns_rico["hasIdentifierType"], type_node))
            g.add((inst_node, ns_rico["hasOrHadIdentifier"], box_node))
            
            # Box Identifier Details
            g.add((box_node, RDF.type, ns_rico["Identifier"]))
            g.add((box_node, RDFS.label, Literal("Box [Number derived from Column]", datatype=XSD.string)))
            
            # Identifier Type Details
            g.add((type_node, RDFS.label, Literal("storage", datatype=XSD.string)))
            
            g.add((subj_node, RDFS.comment, Literal(f"Triggered by {TEMP_BOX_ID} on column '{obj_col}'")))

        # --- B. CUSTOM LOGIC: SENDER PROPAGATION ---
        elif pred_str == TEMP_SENDER or pred_str == TEMP_PROPAGATE:
            # Simulate Agent creation
            agent_node = prefixes_dict["agent"]["EXAMPLE_AGENT_FROM_COLUMN"]
            
            # Link to Subject
            g.add((subj_node, ns_rico["hasSender"], agent_node))
            
            # Agent Details
            g.add((agent_node, RDF.type, ns_rico["Agent"]))
            g.add((agent_node, ns_rico["hasOrHadName"], Literal(f"Name from '{obj_col}'")))
            g.add((agent_node, OWL.sameAs, URIRef("http://viaf.org/viaf/EXAMPLE_ID")))
            
            # Note about the conditional logic (Box 11+)
            if sheet.lower() == "fascicolo":
                institution_node = prefixes_dict["institution"]["EXAMPLE_INSTITUTION"]
                g.add((subj_node, RDFS.comment, Literal("Logic: If Box >= 11, creates institution:Node instead of agent:Node")))

        # --- C. CUSTOM LOGIC: DATE PROCESSING ---
        elif pred_str == "temp:dateProcessing" or "date" in pred_str.lower():
            # Simulate Date entity
            date_node = prefixes_dict["date"][f"DATE_FROM_{make_safe_label(obj_col)}"]
            
            # Determine predicate based on context (defaulting to dateOrDateRange for visual)
            g.add((subj_node, ns_rico["dateOrDateRange"], date_node))
            
            g.add((date_node, RDF.type, ns_rico["Date"]))
            g.add((date_node, ns_rico["normalizedDateValue"], Literal("YYYY-MM-DD", datatype=XSD.date)))
            g.add((date_node, ns_rico["expressedDate"], Literal(f"Value from '{obj_col}'")))

        # --- D. STANDARD MAPPING ---
        else:
            # Handle standard predicates
            if ":" in pred_str:
                pfx, local = pred_str.split(":", 1)
                # Resolve prefix
                if pfx == "rico": predicate = ns_rico[local]
                elif pfx == "rdfs": predicate = RDFS[local]
                elif pfx == "rdf": predicate = RDF[local]
                else: predicate = URIRef(f"{BASE_NS}{pfx}#{local}")
            else:
                predicate = URIRef(pred_str)

            # Determine Object
            # If it links to another structural layer (e.g. includes/includedIn)
            if local in ["includes", "isIncludedIn", "directlyIncludes", "isDirectlyIncludedIn"]:
                # Create a generic placeholder for the target
                target_label = f"TARGET_ENTITY_FROM_{make_safe_label(obj_col)}"
                object_node = ns_structure[target_label]
                g.add((subj_node, predicate, object_node))
            else:
                # Literal or Attribute
                g.add((subj_node, predicate, Literal(f"Data from '{obj_col}'")))

# --- Serialize ---
print(f"\n✅ Structure graph generated.")
print(f"Total triples: {len(g)}")
g.serialize(destination=output_path, format="turtle")
print(f" Saved structural model to {output_path}")