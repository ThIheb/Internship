import pandas as pd
from rdflib import Graph, Namespace, URIRef, Literal
from rdflib.namespace import RDF, RDFS, XSD, OWL

#Configuration
mapping_path = "mapping1.xlsx"
output_path = "output.ttl"
BASE_NS = "http://example.org/"

# Utility Functions
def make_safe_label(value):
    return str(value).strip().replace(" ", "_").replace(":", "").replace("-", "_").upper()


g = Graph()
prefixes = {"rdf": RDF, "rdfs": RDFS, "xsd": XSD, "owl": OWL}
for pfx, ns in prefixes.items(): g.bind(pfx, ns)

# Namespaces
ns_rico = Namespace(f"{BASE_NS}rico#")
g.bind("rico", ns_rico)

# Namespace for the generic structure nodes
ns_structure = Namespace(f"{BASE_NS}structure/")
g.bind("structure", ns_structure)

# Specific namespaces
ns_type = Namespace(f"{BASE_NS}type/")
g.bind("type", ns_type)

ns_inst = Namespace(f"{BASE_NS}inst/")
g.bind("inst", ns_inst)

ns_storageid = Namespace(f"{BASE_NS}storageid/")
g.bind("storageid", ns_storageid)

ns_ident = Namespace(f"{BASE_NS}internalIdentifier/")
g.bind("internalIdentifier", ns_ident)

# Other namespaces
prefixes_dict = {
    "place": Namespace(f"{BASE_NS}place/"),
    "date": Namespace(f"{BASE_NS}date/"),
    "person": Namespace(f"{BASE_NS}person/"),
    "agent": Namespace(f"{BASE_NS}agent/"),
    "institution": Namespace(f"{BASE_NS}institution/"),
    "record": Namespace(f"{BASE_NS}Record/"),
    "recordset": Namespace(f"{BASE_NS}RecordSet/")
}
for k, v in prefixes_dict.items(): g.bind(k, v)

#  STATIC TYPES
storage_type_uri = ns_type["StorageIdentifier"]
g.add((storage_type_uri, RDF.type, ns_rico["IdentifierType"]))
g.add((storage_type_uri, RDFS.label, Literal("Storage Identifier", datatype=XSD.string)))

internal_type_uri = ns_type["InternalIdentifier"]
g.add((internal_type_uri, RDF.type, ns_rico["IdentifierType"]))
g.add((internal_type_uri, RDFS.label, Literal("Internal Identifier", datatype=XSD.string)))


#Triggers
TEMP_BOX_ID = "temp:boxIdentifier"
TEMP_SENDER = "temp:propagateSender" 
TEMP_INTERMEDIATE = "temp:intermediateSender"
RICO_HAS_ID = "rico:hasOrHadIdentifier"
RICO_HAS_ID_SHORT = "rico:hasIdentifier"

#Main Logic

print(f"Loading mapping rules from: {mapping_path}")
mapping_excel = pd.ExcelFile(mapping_path)

for sheet in mapping_excel.sheet_names:
    if "immagin" in sheet.lower() or "image" in sheet.lower():
        continue

    print(f"🔹 Modeling structure for: {sheet}")
    mapping_df = mapping_excel.parse(sheet)
    
    sheet_clean = make_safe_label(sheet)
    subj_node = ns_structure[f"GENERIC_{sheet_clean}"]
    
    g.add((subj_node, RDFS.label, Literal(f"Generic Representative of '{sheet}'")))

    if sheet.lower() in ["serie", "sottoserie", "fascicolo", "fascicoli"]:
        g.add((subj_node, RDF.type, ns_rico["RecordSet"]))
    elif sheet.lower() in ["documento", "documenti"]:
        g.add((subj_node, RDF.type, ns_rico["Record"]))

    # Iterate Rules
    for _, row in mapping_df.iterrows():
        pred_str = str(row["Predicate"]).strip()
        obj_col = str(row["Column Object"]).strip()
        
        # BOX IDENTIFIER
        if pred_str == TEMP_BOX_ID:
            inst_node = ns_inst[f"GENERIC_INSTANTIATION_{sheet_clean}"]
            box_node = ns_storageid[f"GENERIC_BOX_ID_{sheet_clean}"]

            # Link Record to Instantiation
            g.add((inst_node, RDF.type, ns_rico["Instantiation"]))
            g.add((inst_node, ns_rico["isOrWasInstantiationOf"], subj_node))
            
            # Link Instantiation to Box Identifier
            g.add((inst_node, ns_rico["hasOrHadIdentifier"], box_node))
            
            # Link Box ID to the Static Type URI
            g.add((box_node, RDF.type, ns_rico["Identifier"]))
            g.add((box_node, RDFS.label, Literal("Box [Number]", datatype=XSD.string)))
            g.add((box_node, ns_rico["hasIdentifierType"], ns_type["StorageIdentifier"])) 
            
            g.add((subj_node, RDFS.comment, Literal(f"Structure triggered by {TEMP_BOX_ID}")))

        # INTERNAL IDENTIFIER
        elif pred_str in [RICO_HAS_ID, RICO_HAS_ID_SHORT] or "Identifier" in pred_str:
            id_node = ns_ident[f"GENERIC_INTERNAL_ID_{sheet_clean}"]
            g.add((subj_node, ns_rico["hasOrHadIdentifier"], id_node))
            g.add((id_node, RDF.type, ns_rico["Identifier"]))
            g.add((id_node, RDFS.label, Literal(f"Value from {obj_col}", datatype=XSD.string)))
            g.add((id_node, ns_rico["hasIdentifierType"], ns_type["InternalIdentifier"]))

        #SENDER PROPAGATION
        elif pred_str == TEMP_SENDER or pred_str == TEMP_INTERMEDIATE:
            agent_node = prefixes_dict["agent"]["GENERIC_AGENT"]
            
            g.add((subj_node, ns_rico["hasSender"], agent_node))
            g.add((agent_node, RDF.type, ns_rico["Agent"]))
            g.add((agent_node, RDFS.comment, Literal("Logic: Person (or Institution if Box > 10)")))

        # DATE LOGIC
        elif "date" in pred_str.lower() and "processing" in pred_str.lower():
            date_node = prefixes_dict["date"]["GENERIC_DATE"]
            g.add((subj_node, ns_rico["dateOrDateRange"], date_node))
            g.add((date_node, RDF.type, ns_rico["Date"]))
            g.add((date_node, ns_rico["normalizedDateValue"], Literal("YYYY-MM-DD", datatype=XSD.date)))

        #STANDARD MAPPING
        else:
            if ":" in pred_str:
                pfx, local = pred_str.split(":", 1)
                if pfx == "rico": predicate = ns_rico[local]
                elif pfx == "rdfs": predicate = RDFS[local]
                elif pfx == "rdf": predicate = RDF[local]
                else: predicate = URIRef(f"{BASE_NS}{pfx}#{local}")
            else:
                predicate = URIRef(pred_str)

            if local in ["includes", "isIncludedIn", "directlyIncludes", "isDirectlyIncludedIn"]:
                target_label = f"TARGET_ENTITY_FROM_{make_safe_label(obj_col)}"
                object_node = ns_structure[target_label]
                g.add((subj_node, predicate, object_node))
            elif "Place" in local or "place" in pred_str.lower():
                place_node = prefixes_dict["place"]["GENERIC_PLACE"]
                g.add((subj_node, predicate, place_node))
                g.add((place_node, RDF.type, ns_rico["Place"]))
            else:
                g.add((subj_node, predicate, Literal(f"Data from column '{obj_col}'")))

#Serialization
print(f"\n✅ Structure graph generated.")
print(f"Total triples: {len(g)}")
g.serialize(destination=output_path, format="turtle")
print(f" Saved structural model to {output_path}")