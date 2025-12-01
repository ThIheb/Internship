# Documentation:



The Albini collection is mapped in accordance with a hierarchical structure that accommodates classes and subclasses to organize the documents within the collection. When mapped using the Records In Context (RiCo) ontology, the following classes and subclasses are established:

* Serie: the first layer of classes which contains sub-series (mapped as a rico:recordSet) → 4 series within the collection
* Sub-serie: the second layer of classes which contains Busta (mapped as a rico:recordSet) → 7 sub-series within the collection
* Busta: third layer of classes which contains Fascicolo (mapped as a rico:recordSet) → 36 Busta within the collection
* Fascicolo: fourth layer of classes which contains Documento (mapped as a rico:recordSet) → 581 Fascicolo within the collection
* Documento: final layer of recordSet (mapped as a rico:recordSet) → 1852 Documento within the collection
* Immagini: Single documents (images of documents) mapped as rico:recordResource → 6184 images within the collection



This mapping is accomplished using a python script that takes two excel files (A mapping file and an instances file), then generates RDF triples in the Turtle (.ttl) format. The script is designed to read and interpret column-to-predicate mappings and create URI instances for entities. Whenever possible the script enriches entities with external data from external ontologies. The script utilizes RICO (Records in Contexts) ontology to describe archival records, agents, places and their relationships.



## Overview:

The script performs the Following key actions:

1. **Reads mapping rules**: specific relationships defined in an Excel mapping file.
2. **Processes Instance Data**: Iterates through archival data present in the Instances Excel file, going through the different sheets of the document (Serie, Sottoserie, Fascicolo, Documento).
3. **Entity Extraction and Enrichment**: The automatic detection and creation for agents (People/Corporate Bodies), Dates and Places.
4. **Geo-Enrichment**: Connects place names to GeoNames data (via local cache or API) to retrieve longitude and latitude as well as feature classes.
5. **Hierarchy building**: Constructs an archival tree (RecordSets containing Records).
6. **Serialization**: Outputs to a Turtle (.ttl) RDF file.



## Dependencies:

The script relies on non-standard python libraries to run:

* Pandas: used for reading and processing the input file (Excel .xlsx)
* Rdflib: The core library for creating RDF graphs, namespaces and serialization
* Geonamescache: Used for a lookup of GeoNames IDs based on place names
* Requests: for any potential use of GeoNames API implementation



**External Services:**

* **GeoNames Account:** A valid username is required for querying the GeoNames API for coordinates and feature codes if they are not found in the local cache.





## Configuration:

The script relies on several file paths and constants defined at the beginning of the code:



* mapping\_path: with a default value of “mapping1.xlsx” → The excel file defining the column-to-predicate mappings.
* instances\_path: with a default value of “instances.xlsx” → The excel file containing the raw data instances to be transformed
* output\_path: with a default value of “output2.ttl” → The path where the final RDF graph will be saved
* BASE\_NS: with a default value of “http://example.org/” → The root URI for any locally generated entities and custom vocabularies
* User Credentials (GeoNames Username): required for usage of the live GeoNames API for extracting location features and long/lat



**Namespaces:**

The script initializes a BASE\_NS (http://example.org/) and bind the following RiC-O specific sub-namespaces for clean URI generation:

* rico: The main ontology (https://www.ica.org/standards/RiC/ontology#)
* place/ : For generated location entities
* date/ : For generated date entities
* identifier/ : For extracted IDs
* title/ : For document titles
* corporateBody/ : For senders that are corporateBody



## Mapping Logic:

The script processes the input data based on a sheet-by-sheet correspondence:

1. First, iterating through each sheet of the mapping\_path file
2. Second, attempting to find a matching sheet name in the instances\_path file
3. Finally, for each row in the mapping sheet, defining a subject, predicate and object rule



**Hierarchy and Typing:**

The script automatically assigns RDF classes based on the sheet name and ID structure:

* **RecordSet**(rico:RecordSet): Used for 'Serie' and 'Fascicolo'
* **Record**(rico:Record): Used for 'Documento'



## Entity Creation Rules:

The script handles subject-to-object mapping with specific logic. Instead of attaching data as simple strings (Literals), it creates separate nodes (entities) to allow for a richer data description. This logic is determined by the Predicate defined the "mapping1.xlsx".



1. **Places** (rico:isAssociatedWithPlace)

* **Logic:** it checks if the object is a URI or a string
* **Action:** creates a rico:place entity
* **Enrichment:** Cleans the name then searches geonamescache (local), if no name is found in the local cache then it searches the GeoNames API and adds "wgs84:lat", "wgs84:long" and owl:sameAs as well as feature codes to the place entity



2\. **Dates** (rico:date)

* **Logic:** it parses the input string using Regex
* **Supported formats:** YYYY, YYYYMMDD, YYYY-YYYY, YYYYMMDD-YYYYMMDD.
* **Action:** it splits ranges into start and end components

 	- Creates a unique URI based on {Subject}\_{DatePart} to make sure that dates are unique to the record they describe

 	- Adds rico:normalizedDateValue (typed as xsd:date or xsd:gYear)

 	- Adds rico:expressedDate (whenever an expressed date is found)



3\. **Identifiers** (rico:hasOrHadIdentifier)

* **Logic:** Identifiers are treated as distinct objects and not properties
* **Action:**

  * creates a URI in the identifier/ namespaces

 	- types it as rico:Identifier

 	- adds the value as an rdfs:label



4\. **Titles** (rico:hasOrHadTitle)

* **Logic:** The logic is activated when the mapping predicate is rico:hasOrHadTitle
* **Action:** The title text is converted into a safe URI suffix (e.g., "Verbali 1990" --> "Verbali\_1990), then a new node is created and assigned RDF class rico:Title



5\. **Agents** (rico:Person or rico:CorporateBody)

* **Logic:** if a "Box" (busta) number is >11, it treats the sender as a CorporateBody; otherise defaulting to Person (custom logic specific to this dataset)
* **Action:** it creates a rico:Agent entity
* **Special Rule:** Senders attached toe a Fascicolo are temporarily stored and then propogated down to every Documento contained within that folder during post-processing



6\. **Physical Insantiations (Boxes)**

* **Logic:** if a string such as "Busta 10" is detected
* **Action:** it creates a rico:instantiation entity linked to the record and a rico:Identifier of type "storage" for the physical location of each record it finds within the instances Excel file





GeoNames Materialization:

After the initial RDF generation loop is complete, the script performs a data enrichment step:

1. Identify Places: Queries the graph for all subjects that have been assigned rdf:type rico:Place.
2. Lookup ID: For each rico:Place entity, the script utilizes rdfs:label (the original place name) to find a corresponding GeoNames ID via the find\_geonames\_id\_by\_label function



* The function first attempts a lookup using the library geonamescache
* If the cache lookup fails, it uses a fallback features (e.g., ‘bologna’, ‘roma’, ‘paris’) for test cases



3\. Materialize Features: If an ID is found, the fetch\_and\_add\_geonames\_features function adds new triples to the graph:

* Linking the rico:place URI to the official GeoNames URI using owl:sameAs
* Stimulates fetching and adds static GeoNames features (gn:featureClass, gn:featureCode) using the GeoNames ontology namespace (gn)



## Utility Functions:

* make\_safe\_uri\_label(value): Cleans a string value by removing special characters and replacing spaces with underscores, ensuring its validity as a safe URI path segment.
* parse\_normalized\_dates(date\_str): Parses date strings in YYYYMMDD format (single dates or range) and returns start and end components.
* format\_date\_for\_xsd(date\_str): Converts an 8-digit date string into the ISO YYYY-MM-DD format with xsd:date compatibility
* detect\_object\_term(obj\_val\_str, prefixes): Determines the correct RDF term type (Literal, URIRef for VIAF/URL) for an object value
* find\_geonames\_id\_by\_label(label): hybrid searchers: looks in local cache first, then calls the API if the cache is missing returns a numeric ID
* fetch\_and\_add\_geonames\_features(g, place\_uri, geonames\_id, place\_label): takes a GeoName ID, fetches the metadata (Lat/Long/Class) and writes the triples directly to the graph
