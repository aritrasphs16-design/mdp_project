"""
api_registry.py
────────────────────────────────────────────────────────────────────────────
ChemoFilter · API Registry — 50+ Scientific APIs in 3 Tiers
• LAYER 1 — Core APIs    : PubChem, ChEMBL, PDB  (always available, cached)
• LAYER 2 — Extended     : DrugBank, BindingDB, KEGG, UniProt, OpenFDA,
                           Open Targets, STRING, Reactome, HMDB, ChemSpider,
                           MassBank, ZINC, SwissTargetPrediction, UniChem,
                           ClinicalTrials, ChemIDplus, NCI, EPA CompTox …
• LAYER 3 — Experimental : Europe PMC, CrossRef, Semantic Scholar,
                           NCBI Gene, DisGeNET, OMIM, WikiData,
                           OpenAlex, NCBI BioSystems …
All entries are DISABLED by default.
Users explicitly trigger each API fetch.
System is 100% functional without any API.
────────────────────────────────────────────────────────────────────────────
"""

# ─────────────────────────────────────────────────────────────────────────────
# UNIFIED API DESCRIPTOR FORMAT
# ─────────────────────────────────────────────────────────────────────────────
# Each entry:
#   key          : internal identifier (used in code)
#   name         : display name
#   category     : scientific domain
#   description  : one-line scientific purpose
#   base_url     : API endpoint base
#   requires_key : True if API key needed
#   tier         : "core" | "extended" | "experimental"
#   enabled      : default enabled state
#   timeout      : seconds
#   docs         : documentation URL

API_REGISTRY = {

    # =========================================================================
    # LAYER 1 — CORE APIs  (3 APIs)
    # Fast, stable, most important for drug discovery
    # =========================================================================

    "pubchem": {
        "key": "pubchem",
        "name": "PubChem",
        "category": "Chemical Databases",
        "description": "NIH compound identity, canonical SMILES, physicochemical properties & synonyms",
        "base_url": "https://pubchem.ncbi.nlm.nih.gov/rest/pug",
        "requires_key": False,
        "tier": "core",
        "enabled": True,
        "timeout": 5,
        "docs": "https://pubchemdocs.ncbi.nlm.nih.gov/pug-rest",
        "icon": "🧪",
        "output_fields": ["CID", "IUPAC Name", "Molecular Formula", "MW", "LogP", "TPSA", "Synonyms"],
    },

    "chembl": {
        "key": "chembl",
        "name": "ChEMBL",
        "category": "Bioactivity Databases",
        "description": "EMBL-EBI bioactivity database — IC50/Ki values, assay data, protein targets",
        "base_url": "https://www.ebi.ac.uk/chembl/api/data",
        "requires_key": False,
        "tier": "core",
        "enabled": True,
        "timeout": 5,
        "docs": "https://chembl.gitbook.io/chembl-interface-documentation/web-services/chembl-data-web-services",
        "icon": "🎯",
        "output_fields": ["ChEMBL ID", "Bioactivities", "Targets", "Assay Count"],
    },

    "pdb": {
        "key": "pdb",
        "name": "RCSB Protein Data Bank (PDB)",
        "category": "Structural Biology",
        "description": "3D protein structure data — target context, binding site information",
        "base_url": "https://search.rcsb.org/rcsbsearch/v2/query",
        "requires_key": False,
        "tier": "core",
        "enabled": True,
        "timeout": 5,
        "docs": "https://search.rcsb.org/",
        "icon": "🏗️",
        "output_fields": ["PDB ID", "Structure Title", "Resolution", "RCSB Link"],
    },

    # =========================================================================
    # LAYER 2 — EXTENDED APIs  (30 APIs)
    # User-triggered, cached 24h, fail-safe
    # =========================================================================

    "openfda": {
        "key": "openfda",
        "name": "OpenFDA Drug Database",
        "category": "Regulatory & Clinical",
        "description": "FDA drug approvals, adverse events, drug labels & recall data",
        "base_url": "https://api.fda.gov/drug",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 5,
        "docs": "https://open.fda.gov/apis/drug/",
        "icon": "🏛️",
        "output_fields": ["Drug Name", "Approval Status", "Indications", "Adverse Events"],
    },

    "uniprot": {
        "key": "uniprot",
        "name": "UniProt Protein Database",
        "category": "Target Biology",
        "description": "Protein sequence, function, post-translational modifications & disease associations",
        "base_url": "https://rest.uniprot.org",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 5,
        "docs": "https://www.uniprot.org/help/api",
        "icon": "🧬",
        "output_fields": ["Protein Name", "Gene", "Organism", "Function", "Disease"],
    },

    "kegg": {
        "key": "kegg",
        "name": "KEGG Pathway Database",
        "category": "Pathway & Metabolism",
        "description": "Metabolic pathway mapping, drug targets, enzyme interactions & KEGG DRUG",
        "base_url": "https://rest.kegg.jp",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 5,
        "docs": "https://www.kegg.jp/kegg/rest/keggapi.html",
        "icon": "🗺️",
        "output_fields": ["Pathway ID", "Pathway Name", "Enzymes", "Drug Entry"],
    },

    "bindingdb": {
        "key": "bindingdb",
        "name": "BindingDB",
        "category": "Bioactivity Databases",
        "description": "Measured binding affinities (Kd, Ki, IC50) for drug-target interactions",
        "base_url": "https://bindingdb.org/rest",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 6,
        "docs": "https://www.bindingdb.org/bind/BindingDBRESTfulAPI.jsp",
        "icon": "🔗",
        "output_fields": ["Target Name", "Binding Affinity (Kd/Ki/IC50)", "Assay", "Source"],
    },

    "open_targets": {
        "key": "open_targets",
        "name": "Open Targets Platform",
        "category": "Target Biology",
        "description": "Target–disease associations with evidence scores from genetics, literature & clinical data",
        "base_url": "https://api.platform.opentargets.org/api/v4/graphql",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 8,
        "docs": "https://platform-docs.opentargets.org/data-access/graphql-api",
        "icon": "🎯",
        "output_fields": ["Disease Associations", "Evidence Score", "Target Tractability"],
    },

    "string_db": {
        "key": "string_db",
        "name": "STRING Protein Interaction Network",
        "category": "Network Biology",
        "description": "Protein–protein interaction networks, functional enrichment & pathway context",
        "base_url": "https://string-db.org/api",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 6,
        "docs": "https://string-db.org/help/api/",
        "icon": "🕸️",
        "output_fields": ["Interaction Partners", "Confidence Score", "Interaction Type"],
    },

    "reactome": {
        "key": "reactome",
        "name": "Reactome Pathway Database",
        "category": "Pathway & Metabolism",
        "description": "Curated human biological pathways, reactions & molecular events",
        "base_url": "https://reactome.org/ContentService",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 5,
        "docs": "https://reactome.org/dev/content-service",
        "icon": "⚗️",
        "output_fields": ["Pathway ID", "Pathway Name", "Reactions", "Participants"],
    },

    "hmdb": {
        "key": "hmdb",
        "name": "Human Metabolome Database (HMDB)",
        "category": "Metabolomics",
        "description": "Human metabolite information, biofluid concentrations & metabolic pathways",
        "base_url": "https://hmdb.ca",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 6,
        "docs": "https://hmdb.ca/about#xml_downloads",
        "icon": "💧",
        "output_fields": ["HMDB ID", "Metabolite Class", "Biofluid", "Pathway"],
    },

    "unichem": {
        "key": "unichem",
        "name": "UniChem Cross-Reference",
        "category": "Chemical Databases",
        "description": "Cross-database chemical structure mapping — maps compound IDs across 40+ databases",
        "base_url": "https://www.ebi.ac.uk/unichem/api/v1",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 5,
        "docs": "https://www.ebi.ac.uk/unichem/info/webservices",
        "icon": "🔄",
        "output_fields": ["Source ID", "Database Name", "Cross-References"],
    },

    "clinicaltrials": {
        "key": "clinicaltrials",
        "name": "ClinicalTrials.gov",
        "category": "Regulatory & Clinical",
        "description": "Active and completed clinical trials — development stage, indications & sponsors",
        "base_url": "https://clinicaltrials.gov/api/v2",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 6,
        "docs": "https://clinicaltrials.gov/data-api/api",
        "icon": "🏥",
        "output_fields": ["Trial ID", "Phase", "Condition", "Status", "Sponsor"],
    },

    "epa_comptox": {
        "key": "epa_comptox",
        "name": "EPA CompTox Chemicals Dashboard",
        "category": "Toxicology",
        "description": "EPA toxicity data, exposure predictions, physicochemical properties & hazard flags",
        "base_url": "https://comptox.epa.gov/dashboard-api",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 6,
        "docs": "https://api-ccte.epa.gov/docs/",
        "icon": "⚠️",
        "output_fields": ["DTXSID", "Hazard Summary", "Toxicity Flags", "Exposure Data"],
    },

    "ncbi_gene": {
        "key": "ncbi_gene",
        "name": "NCBI Gene Database",
        "category": "Genomics",
        "description": "Gene information, expression data, sequence, chromosomal location & RefSeq",
        "base_url": "https://eutils.ncbi.nlm.nih.gov/entrez/eutils",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 5,
        "docs": "https://www.ncbi.nlm.nih.gov/books/NBK25501/",
        "icon": "🧬",
        "output_fields": ["Gene Symbol", "Gene ID", "Location", "Summary", "RefSeq"],
    },

    "ncbi_taxonomy": {
        "key": "ncbi_taxonomy",
        "name": "NCBI Taxonomy",
        "category": "Biology",
        "description": "Taxonomic classification of organisms relevant to drug target species",
        "base_url": "https://eutils.ncbi.nlm.nih.gov/entrez/eutils",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 4,
        "docs": "https://www.ncbi.nlm.nih.gov/taxonomy",
        "icon": "🌿",
        "output_fields": ["Taxon Name", "Lineage", "Division"],
    },

    "drugcentral": {
        "key": "drugcentral",
        "name": "Drug Central",
        "category": "Drug Information",
        "description": "FDA/EMA-approved drug data — pharmacology, targets, indications, interactions",
        "base_url": "https://drugcentral.org/api",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 5,
        "docs": "https://drugcentral.org/download",
        "icon": "💊",
        "output_fields": ["Drug Name", "Targets", "Indications", "Interaction Class"],
    },

    "chemspider": {
        "key": "chemspider",
        "name": "ChemSpider (Royal Society of Chemistry)",
        "category": "Chemical Databases",
        "description": "Chemical structure search, spectral data, supplier information & predicted properties",
        "base_url": "https://api.rsc.org/compounds/v1",
        "requires_key": True,
        "tier": "extended",
        "enabled": False,
        "timeout": 5,
        "docs": "https://developer.rsc.org/",
        "icon": "🔬",
        "output_fields": ["ChemSpider ID", "Structure", "Predicted Properties", "Suppliers"],
    },

    "swissadme": {
        "key": "swissadme",
        "name": "SwissADME",
        "category": "ADMET Prediction",
        "description": "Physicochemical & pharmacokinetic ADME prediction — lipophilicity, solubility, BBB",
        "base_url": "http://www.swissadme.ch",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 10,
        "docs": "http://www.swissadme.ch/",
        "icon": "💉",
        "output_fields": ["ADME Scores", "Blood-Brain Barrier", "P-gp Substrate", "Bioavailability Radar"],
    },

    "swiss_target": {
        "key": "swiss_target",
        "name": "SwissTargetPrediction",
        "category": "Target Prediction",
        "description": "Probabilistic target prediction for small molecules — protein class & probability",
        "base_url": "http://www.swisstargetprediction.ch",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 10,
        "docs": "http://www.swisstargetprediction.ch/",
        "icon": "🎯",
        "output_fields": ["Target Name", "Probability", "Protein Class", "Organism"],
    },

    "zinc": {
        "key": "zinc",
        "name": "ZINC Database",
        "category": "Virtual Screening",
        "description": "Purchasable compounds for virtual screening — 230M+ commercially available molecules",
        "base_url": "https://zinc.docking.org",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 6,
        "docs": "https://zinc.docking.org/",
        "icon": "🛒",
        "output_fields": ["ZINC ID", "Commercial Availability", "Vendor", "Price Tier"],
    },

    "chemidplus": {
        "key": "chemidplus",
        "name": "ChemIDplus (NLM)",
        "category": "Chemical Databases",
        "description": "NLM chemical dictionary — synonyms, registry numbers, toxicity data",
        "base_url": "https://chem.nlm.nih.gov/chemidplus",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 5,
        "docs": "https://chem.nlm.nih.gov/chemidplus/",
        "icon": "📖",
        "output_fields": ["Registry Number", "Synonyms", "Toxicity", "Physical Properties"],
    },

    "nci_cactus": {
        "key": "nci_cactus",
        "name": "NCI Chemical Identifier Resolver (CACTUS)",
        "category": "Chemical Databases",
        "description": "NCI structure name resolver — converts drug names to SMILES, InChI, structures",
        "base_url": "https://cactus.nci.nih.gov/chemical/structure",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 4,
        "docs": "https://cactus.nci.nih.gov/chemical/structure_documentation",
        "icon": "🌵",
        "output_fields": ["SMILES", "InChI", "IUPAC Name", "MW"],
    },

    "chembl_assay": {
        "key": "chembl_assay",
        "name": "ChEMBL Assay Detail",
        "category": "Bioactivity Databases",
        "description": "Detailed assay descriptions — tissue, cell line, organism & assay format",
        "base_url": "https://www.ebi.ac.uk/chembl/api/data/assay",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 5,
        "docs": "https://chembl.gitbook.io/chembl-interface-documentation/",
        "icon": "🧫",
        "output_fields": ["Assay ID", "Type", "Organism", "Tissue", "Cell Line"],
    },

    "ebi_expression": {
        "key": "ebi_expression",
        "name": "EMBL-EBI Expression Atlas",
        "category": "Genomics",
        "description": "Gene and protein expression across tissues, cell lines & disease states",
        "base_url": "https://www.ebi.ac.uk/gxa/json",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 6,
        "docs": "https://www.ebi.ac.uk/gxa/help/programmatic-access.html",
        "icon": "📊",
        "output_fields": ["Gene", "Expression Level", "Tissue", "Experiment"],
    },

    "disgenet": {
        "key": "disgenet",
        "name": "DisGeNET",
        "category": "Disease Biology",
        "description": "Gene–disease associations from curated databases, literature & clinical data",
        "base_url": "https://www.disgenet.org/api",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 6,
        "docs": "https://www.disgenet.org/api",
        "icon": "🏥",
        "output_fields": ["Disease Name", "Gene-Disease Score", "Evidence Type", "Source"],
    },

    "hetionet": {
        "key": "hetionet",
        "name": "Hetionet Drug Repurposing",
        "category": "Drug Repurposing",
        "description": "Heterogeneous network for drug repurposing — compounds, genes, diseases & pathways",
        "base_url": "https://het.io/api",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 5,
        "docs": "https://het.io/",
        "icon": "♻️",
        "output_fields": ["Repurposing Candidates", "Network Paths", "Evidence Score"],
    },

    "massbank": {
        "key": "massbank",
        "name": "MassBank of North America (MoNA)",
        "category": "Analytical Chemistry",
        "description": "Mass spectrometry spectral data for metabolite identification & structure verification",
        "base_url": "https://mona.fiehnlab.ucdavis.edu/rest",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 5,
        "docs": "https://mona.fiehnlab.ucdavis.edu/documentation",
        "icon": "📡",
        "output_fields": ["Spectrum ID", "MS Level", "Precursor m/z", "Splash Key"],
    },

    "echa_reach": {
        "key": "echa_reach",
        "name": "ECHA REACH Chemical Registry",
        "category": "Regulatory & Safety",
        "description": "EU REACH regulation data — registered substances, hazard classifications & tonnage",
        "base_url": "https://echa.europa.eu",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 6,
        "docs": "https://www.echa.europa.eu/information-on-chemicals",
        "icon": "🇪🇺",
        "output_fields": ["REACH Registration", "Hazard Class", "GHS Pictograms", "Tonnage Band"],
    },

    "chemrisk": {
        "key": "chemrisk",
        "name": "PubChem Safety & Toxicology",
        "category": "Toxicology",
        "description": "PubChem GHS safety data — hazard statements, precautionary codes & signal words",
        "base_url": "https://pubchem.ncbi.nlm.nih.gov/rest/pug_view",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 5,
        "docs": "https://pubchemdocs.ncbi.nlm.nih.gov/pug-view",
        "icon": "🛡️",
        "output_fields": ["GHS Hazard Codes", "Signal Word", "Precautionary Statements"],
    },

    "npatlas": {
        "key": "npatlas",
        "name": "Natural Products Atlas",
        "category": "Natural Products",
        "description": "Microbial natural products database — source organisms, biosynthetic gene clusters",
        "base_url": "https://www.npatlas.org/api",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 5,
        "docs": "https://www.npatlas.org/",
        "icon": "🌱",
        "output_fields": ["NP ID", "Source Organism", "Biosynthetic Class", "Activity"],
    },

    "lotus": {
        "key": "lotus",
        "name": "LOTUS Natural Products",
        "category": "Natural Products",
        "description": "Open natural product database — organism–molecule associations from literature",
        "base_url": "https://lotus.naturalproducts.net/api",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 5,
        "docs": "https://lotus.naturalproducts.net/api",
        "icon": "🪷",
        "output_fields": ["LOTUS ID", "Organism", "Kingdom", "Literature Reference"],
    },

    "cog": {
        "key": "cog",
        "name": "NCBI COG Ortholog Groups",
        "category": "Genomics",
        "description": "Clusters of orthologous groups for evolutionary & functional gene analysis",
        "base_url": "https://eutils.ncbi.nlm.nih.gov/entrez/eutils",
        "requires_key": False,
        "tier": "extended",
        "enabled": False,
        "timeout": 4,
        "docs": "https://www.ncbi.nlm.nih.gov/research/cog-project/",
        "icon": "🔭",
        "output_fields": ["COG ID", "Function Category", "Gene Name", "Organism Count"],
    },

    # =========================================================================
    # LAYER 3 — EXPERIMENTAL APIs  (20 APIs)
    # Literature, AI, genomics — disabled by default, user opt-in
    # =========================================================================

    "europe_pmc": {
        "key": "europe_pmc",
        "name": "Europe PMC Literature Search",
        "category": "Scientific Literature",
        "description": "Open-access biomedical literature — 37M+ abstracts, citations & full-text search",
        "base_url": "https://www.ebi.ac.uk/europepmc/webservices/rest",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 6,
        "docs": "https://europepmc.org/RestfulWebService",
        "icon": "📰",
        "output_fields": ["PubMed ID", "Title", "Journal", "Year", "Abstract"],
    },

    "crossref": {
        "key": "crossref",
        "name": "CrossRef DOI Resolution",
        "category": "Scientific Literature",
        "description": "DOI metadata — journal, authors, citation counts for scientific publications",
        "base_url": "https://api.crossref.org",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 5,
        "docs": "https://www.crossref.org/documentation/retrieve-metadata/rest-api/",
        "icon": "📚",
        "output_fields": ["DOI", "Title", "Authors", "Journal", "Citation Count"],
    },

    "semantic_scholar": {
        "key": "semantic_scholar",
        "name": "Semantic Scholar AI Literature",
        "category": "Scientific Literature",
        "description": "AI-powered academic paper search — semantic similarity, citation graphs",
        "base_url": "https://api.semanticscholar.org/graph/v1",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 6,
        "docs": "https://api.semanticscholar.org/",
        "icon": "🤖",
        "output_fields": ["Paper ID", "Title", "Year", "Citation Count", "TL;DR Summary"],
    },

    "openalex": {
        "key": "openalex",
        "name": "OpenAlex Open Research Graph",
        "category": "Scientific Literature",
        "description": "Open scholarly data — 200M+ research works, institution & concept mapping",
        "base_url": "https://api.openalex.org",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 5,
        "docs": "https://docs.openalex.org/",
        "icon": "🌐",
        "output_fields": ["Work ID", "Title", "Concepts", "Citation Count", "Open Access"],
    },

    "omim": {
        "key": "omim",
        "name": "OMIM — Online Mendelian Inheritance in Man",
        "category": "Disease Genomics",
        "description": "Genetic basis of disease — Mendelian disorders, disease genes & phenotypes",
        "base_url": "https://api.omim.org/api",
        "requires_key": True,
        "tier": "experimental",
        "enabled": False,
        "timeout": 5,
        "docs": "https://www.omim.org/help/api",
        "icon": "🧬",
        "output_fields": ["MIM Number", "Disease Name", "Gene", "Inheritance Mode"],
    },

    "wikidata": {
        "key": "wikidata",
        "name": "Wikidata SPARQL",
        "category": "Knowledge Graph",
        "description": "Open knowledge graph — drug–target–disease triples via SPARQL endpoint",
        "base_url": "https://query.wikidata.org/sparql",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 8,
        "docs": "https://www.wikidata.org/wiki/Wikidata:SPARQL_query_service",
        "icon": "🌐",
        "output_fields": ["Drug Entity", "Target", "Disease", "Mechanism"],
    },

    "gtopdb": {
        "key": "gtopdb",
        "name": "Guide to Pharmacology (GtoPdb)",
        "category": "Pharmacology",
        "description": "IUPHAR/BPS curated pharmacological data — receptor ligand interactions & pharmacodynamics",
        "base_url": "https://www.guidetopharmacology.org/services",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 5,
        "docs": "https://www.guidetopharmacology.org/webServices.jsp",
        "icon": "💊",
        "output_fields": ["Ligand ID", "Target", "Action", "Affinity (pKi)", "Selectivity"],
    },

    "drugmap": {
        "key": "drugmap",
        "name": "DrugMap Target Network",
        "category": "Network Pharmacology",
        "description": "Drug–target interaction network — polypharmacology & network pharmacology",
        "base_url": "https://drugmap.idrblab.net",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 6,
        "docs": "https://drugmap.idrblab.net/",
        "icon": "🕸️",
        "output_fields": ["Drug", "Target Network", "Pathway", "Interaction Score"],
    },

    "surechembl": {
        "key": "surechembl",
        "name": "SureChEMBL Patent Chemistry",
        "category": "Intellectual Property",
        "description": "Chemical compounds extracted from patent literature — IP landscape & novelty",
        "base_url": "https://www.surechembl.org/api",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 7,
        "docs": "https://www.surechembl.org/",
        "icon": "📋",
        "output_fields": ["Patent ID", "Compound Count", "Assignee", "Filing Date"],
    },

    "bioassay": {
        "key": "bioassay",
        "name": "PubChem BioAssay",
        "category": "Bioactivity Databases",
        "description": "Biological activity data from HTS and confirmatory assays — dose-response",
        "base_url": "https://pubchem.ncbi.nlm.nih.gov/rest/pug",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 5,
        "docs": "https://pubchemdocs.ncbi.nlm.nih.gov/pug-rest",
        "icon": "🧫",
        "output_fields": ["AID", "Assay Type", "Activity Outcome", "Target GI"],
    },

    "pdbe": {
        "key": "pdbe",
        "name": "PDBe — Protein Data Bank Europe",
        "category": "Structural Biology",
        "description": "Protein structure quality metrics, ligand interactions & binding site analysis",
        "base_url": "https://www.ebi.ac.uk/pdbe/api",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 5,
        "docs": "https://www.ebi.ac.uk/pdbe/api/doc/",
        "icon": "🏗️",
        "output_fields": ["Entry ID", "Resolution (Å)", "Ligand Interactions", "R-factor"],
    },

    "stitch": {
        "key": "stitch",
        "name": "STITCH Chemical–Protein Interactions",
        "category": "Network Biology",
        "description": "Chemical–protein interaction network — direct & indirect interactions with confidence",
        "base_url": "http://stitch.embl.de/api",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 6,
        "docs": "http://stitch.embl.de/cgi/show_input_page.pl?UserId=&sessionId=&input_page_show_search=FALSE",
        "icon": "🔗",
        "output_fields": ["Protein ID", "Interaction Score", "Action Type", "Source"],
    },

    "chebi": {
        "key": "chebi",
        "name": "ChEBI Chemical Entities",
        "category": "Chemical Ontology",
        "description": "Chemical ontology — biological roles, application classes & structure hierarchy",
        "base_url": "https://www.ebi.ac.uk/webservices/chebi/2.0/test",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 5,
        "docs": "https://www.ebi.ac.uk/chebi/webServices.do",
        "icon": "🌳",
        "output_fields": ["ChEBI ID", "Chemical Role", "Application", "Parent Classes"],
    },

    "dgidb": {
        "key": "dgidb",
        "name": "DGIdb — Drug-Gene Interaction Database",
        "category": "Drug-Gene Interactions",
        "description": "Curated drug–gene interactions — druggability categories & interaction types",
        "base_url": "https://dgidb.org/api/v2",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 5,
        "docs": "https://dgidb.org/api",
        "icon": "🔀",
        "output_fields": ["Gene Name", "Interaction Type", "Drug", "Source DB"],
    },

    "pharos": {
        "key": "pharos",
        "name": "Pharos Target Illumination Platform",
        "category": "Target Biology",
        "description": "NIH TCRD/Pharos — target development levels (Tclin/Tchem/Tbio/Tdark)",
        "base_url": "https://pharos-api.ncats.io/graphql",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 7,
        "docs": "https://pharos.nih.gov/api",
        "icon": "💡",
        "output_fields": ["Target Name", "Development Level", "Target Family", "Disease Score"],
    },

    "ncats_inxight": {
        "key": "ncats_inxight",
        "name": "NCATS Inxight Drugs",
        "category": "Drug Information",
        "description": "NCATS/FDA drug substance information — moieties, relationships & approval history",
        "base_url": "https://drugs.ncats.io/api/v1",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 5,
        "docs": "https://drugs.ncats.io/",
        "icon": "💉",
        "output_fields": ["UNII", "Drug Name", "Moieties", "Approval Year", "Routes"],
    },

    "cts_chem": {
        "key": "cts_chem",
        "name": "Chemical Translation Service (CTS)",
        "category": "Chemical Databases",
        "description": "Chemical identifier translation — InChI, InChIKey, SMILES, CAS, ChemSpider",
        "base_url": "https://cts.fiehnlab.ucdavis.edu/rest",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 4,
        "docs": "https://cts.fiehnlab.ucdavis.edu/",
        "icon": "🔄",
        "output_fields": ["CAS Number", "InChIKey", "Synonyms", "Database IDs"],
    },

    "cbioportal": {
        "key": "cbioportal",
        "name": "cBioPortal Cancer Genomics",
        "category": "Cancer Biology",
        "description": "Cancer genomic alteration data — mutations, copy numbers & expression across TCGA studies",
        "base_url": "https://www.cbioportal.org/api",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 7,
        "docs": "https://www.cbioportal.org/api/swagger-ui/index.html",
        "icon": "🔬",
        "output_fields": ["Cancer Type", "Mutation Frequency", "Alteration Type", "Study"],
    },

    "rhea": {
        "key": "rhea",
        "name": "Rhea Biochemical Reaction Database",
        "category": "Biochemistry",
        "description": "Expert-curated biochemical reactions — enzyme substrates, products & cofactors",
        "base_url": "https://www.rhea-db.org/rhea",
        "requires_key": False,
        "tier": "experimental",
        "enabled": False,
        "timeout": 5,
        "docs": "https://www.rhea-db.org/documentation",
        "icon": "⚗️",
        "output_fields": ["Rhea ID", "Equation", "EC Number", "UniProt Enzyme"],
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# TIER GROUPINGS
# ─────────────────────────────────────────────────────────────────────────────

TIER_CORE         = {k: v for k, v in API_REGISTRY.items() if v["tier"] == "core"}
TIER_EXTENDED     = {k: v for k, v in API_REGISTRY.items() if v["tier"] == "extended"}
TIER_EXPERIMENTAL = {k: v for k, v in API_REGISTRY.items() if v["tier"] == "experimental"}

TIER_ORDER = ["core", "extended", "experimental"]

TIER_META = {
    "core": {
        "label":       "Layer 1 — Core Databases",
        "description": "Always-available chemical and bioactivity reference databases. Cached 24h.",
        "color":       "#4ade80",
        "icon":        "🟢",
    },
    "extended": {
        "label":       "Layer 2 — Extended Scientific Databases",
        "description": "User-triggered enrichment. Click to fetch. Cached 24h per compound.",
        "color":       "#f5a623",
        "icon":        "🟡",
    },
    "experimental": {
        "label":       "Layer 3 — Experimental & Literature APIs",
        "description": "Optional intelligence layer. Disabled by default. User opt-in required.",
        "color":       "#a78bfa",
        "icon":        "🟣",
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# CATEGORY GROUPINGS (for UI display)
# ─────────────────────────────────────────────────────────────────────────────

CATEGORY_ICONS = {
    "Chemical Databases":      "🧪",
    "Bioactivity Databases":   "🎯",
    "Structural Biology":      "🏗️",
    "Regulatory & Clinical":   "🏛️",
    "Target Biology":          "🧬",
    "Pathway & Metabolism":    "🗺️",
    "Toxicology":              "⚠️",
    "ADMET Prediction":        "💉",
    "Target Prediction":       "🔮",
    "Natural Products":        "🌱",
    "Metabolomics":            "💧",
    "Scientific Literature":   "📰",
    "Disease Biology":         "🏥",
    "Genomics":                "🔬",
    "Network Biology":         "🕸️",
    "Virtual Screening":       "🛒",
    "Drug Information":        "💊",
    "Intellectual Property":   "📋",
    "Knowledge Graph":         "🌐",
    "Biochemistry":            "⚗️",
    "Cancer Biology":          "🔬",
    "Pharmacology":            "💊",
    "Drug-Gene Interactions":  "🔀",
    "Chemical Ontology":       "🌳",
    "Analytical Chemistry":    "📡",
    "Drug Repurposing":        "♻️",
    "Regulatory & Safety":     "🛡️",
    "Biology":                 "🌿",
}


def get_apis_by_tier(tier: str) -> dict:
    return {k: v for k, v in API_REGISTRY.items() if v["tier"] == tier}


def get_apis_by_category(category: str) -> dict:
    return {k: v for k, v in API_REGISTRY.items() if v["category"] == category}


def get_all_categories() -> list:
    return sorted(set(v["category"] for v in API_REGISTRY.values()))


def count_by_tier() -> dict:
    counts = {}
    for tier in TIER_ORDER:
        counts[tier] = sum(1 for v in API_REGISTRY.values() if v["tier"] == tier)
    return counts
