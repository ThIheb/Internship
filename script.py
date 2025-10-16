import pandas as pd
from rdflib import Graph, URIRef, Literal, Namespace
from rdflib.namespace import RDF, XSD

INPUT_FILE = "mapping1.xlsx"
OUTPUT_FILE = "output.ttl"
EX = Namespace("http://example.org/")
rico = Namespace("http://www.ica.org/standards/RiC/ontology#")
# Load the sheets
xls = pd.ExcelFile(INPUT_FILE)
sheet_names = xls.sheet_names
print(f"Found sheets: {sheet_names}")

# create RDF graph
g = Graph()

# Pre-bind standard namespaces
g.bind("rdf", RDF)
g.bind("xsd", XSD)
g.bind("rico", rico)

# prefixes collection
def extract_prefixes(df):
    prefixes = set()
    for col in ["Subject", "Predicate", "Object"]:
        if col not in df.columns:
            continue
        for val in df[col].dropna():
            val = str(val).strip()
            if ":" in val:
                prefixes.add(val.split(":", 1)[0])
    return prefixes

all_prefixes = set()
for name in sheet_names:
    df = pd.read_excel(INPUT_FILE, sheet_name=name)
    all_prefixes |= extract_prefixes(df)

# Known standard prefixes
KNOWN_PREFIXES = {"rdf", "xsd", "rico"}
prefix_map = {
    "rdf": RDF,
    "xsd": XSD,
    "rico": rico
}

# Bind unknown prefixes dynamically
for p in all_prefixes:
    if p not in KNOWN_PREFIXES:
        ns_uri = f"http://example.org/{p}#"
        prefix_map[p] = Namespace(ns_uri)
        g.bind(p, prefix_map[p])

print(f"Auto-detected prefixes: {list(prefix_map.keys())}")

# parse prefixed prefixes
def parse_prefixed(value):
    value = str(value).strip()
    if ":" in value:
        prefix, local = value.split(":", 1)
        if prefix in prefix_map:
            return prefix_map[prefix][local]
    return EX[value.replace(" ", "_")]

# parse sheets
for sheet in sheet_names:
    print(f"Processing sheet: {sheet}")
    df = pd.read_excel(INPUT_FILE, sheet_name=sheet)

    for _, row in df.iterrows():
        subj = row.get("Subject")
        pred = row.get("Predicate")
        obj = row.get("Object")

        if pd.isna(subj) or pd.isna(pred) or pd.isna(obj):
            continue

        subj_uri = parse_prefixed(subj)
        pred_uri = parse_prefixed(pred)

        obj_str = str(obj).strip()
        if obj_str.startswith("XSD:") or obj_str.startswith("xsd:"):
            obj_value = Literal(obj_str.split(":")[1])
        else:
            obj_value = parse_prefixed(obj_str)

        # RDFLib automatically ignores duplicate triples
        g.add((subj_uri, pred_uri, obj_value))

# serialize to ttl file
g.serialize(destination=OUTPUT_FILE, format="turtle")
print(f"\n RDF graph saved to {OUTPUT_FILE}")
print(f"Total unique triples: {len(g)}")
