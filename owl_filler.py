# -*- coding: utf-8 -*-
"""
Created on Fri Nov 24 11:58:31 2023

@author: MCubaud
"""
import os
import pandas as pd
import owlready2  as or2
import re
import pylatexenc.latex2text

from metadata_enrichment import enrich_metadata

TRUE_VALUES = ["yes", "true", "1", "t", "y", "0b1", "√", "☑", "✔", "oui", "si"]
FALSE_VALUES = ["", None, "no", "false", "0", "f", "n", "w", "non"]

def is_date(string):
    try:
        pd.to_datetime(string, errors='raise')
        return True
    except (ValueError, TypeError):
        return False

def is_place_name(name):
    geolocator = Nominatim(user_agent="place_checker")
    try:
        location = geolocator.geocode(name, timeout=10)
        return location is not None
    except GeocoderUnavailable:
        print("error")
        return False

def decode_latex(latex_string):
    try:
        decoded_string = pylatexenc.latex2text.LatexNodes2Text().latex2text(latex_string)
        return decoded_string
    except Exception as e:
        try:
            decoded_string = pylatexenc.latex2text.latex2text(latex_string)
            return decoded_string
        except Exception as e2:
          print("Error:", e, e2)
          return latex_string

# --- Utility to parse grouped fields ---
def parse_grouped_field(field):
    if pd.isna(field): return []
    field = str(field).strip()
    if field.startswith("{") and field.endswith("}"):
        # Parse outermost groups
        return [item.strip("{} ") for item in re.split(r"}\s*;\s*{", field[1:-1])]
    else:
        return [item.strip() for item in re.split(r"\s?;\s?", field)]

def get_group_hierarchy(text):
    elements = re.split(r"\s?;\s?", text)
    group_numbers = []
    current_group = 0
    is_in_group = False
    for element in elements:
        group_numbers.append(current_group)
        is_in_group = (
             "(" in element or (is_in_group and ")" not in element)
        )
        if not is_in_group:
            current_group += 1

    return group_numbers


def article_metadata(onto, row):
    doi = row["doi"]
    article = onto["paper"](urllib.parse.quote(doi))
    article.doi = doi

    article.title = row["Title"]
    article.label = row["Title"]

    if not row.isna()["journal"]:
        journal = onto[row["type of publication"]](urllib.parse.quote(
            row["type of publication"]+
            "_"+
            row["journal"].replace(" ","_")
            ))
        journal.namePublisher.append(row["journal"])
        journal.label = row["journal"]
        article.isPublishedIn.append(journal)

    #affiliations
    affiliation_instances = []
    if not row.isna()["Affiliation Name"]:
        list_affiliation_names = re.split(r"\s?;\s?", row["Affiliation Name"])
        list_affiliation_addresses = re.split(r"\s?;\s?", row["Affiliation Address"])
        for i, affiliation_name in enumerate(list_affiliation_names):
            if affiliation_name not in ["", " "]:
                affiliation = onto["affiliation"](urllib.parse.quote(
                    affiliation_name.replace(" ","_").replace(",","")+"_"+
                    list_affiliation_addresses[i].replace(" ","_").replace(",","")
                    ))
                affiliation.label = affiliation_name
                affiliation.affiliationName.append(affiliation_name)
                affiliation.affiliationAddress.append(list_affiliation_addresses[i])
                affiliation_instances.append(affiliation)
            else:
                affiliation_instances.append(None)

    if not row.isna()["Authors"]:
        list_authors = row["Authors"].split(" and ")
        if len(affiliation_instances)==1:
            #Then we suppose all the authors have the same affiliation
            affiliation_instances = affiliation_instances * len(list_authors)
        for i, name_author in enumerate(list_authors):
            name_author = decode_latex(name_author)
            print(name_author)
            author = onto["author"](urllib.parse.quote(re.sub(r"\s?,\s?", "_", name_author)))
            author.label = name_author
            list_author_names = re.split(r"\s?,\s?", name_author)
            if len(list_author_names)==1:
                author.lastName = list_author_names[0]
            elif len(list_author_names)==2:
                if "." in list_author_names[1]:
                    #The order may not be what it is supposed to be
                    author.firstName = list_author_names[0]
                    author.lastName = list_author_names[1]
                else:
                    author.firstName = list_author_names[1]
                    author.lastName = list_author_names[0]
            else:
                name0 = list_author_names[0]
                name1 = " ".join(list_author_names[1:])
                if "." in name0:
                    author.firstName = name0
                    author.lastName = name1
                else:
                    author.firstName = name1
                    author.lastName = name0
            if len(affiliation_instances)!=0:
                if affiliation_instances[i] is not None:
                    author.hasAffiliation.append(affiliation_instances[i])
            article.hasAuthor.append(author)

    if not row.isna()["Year"]:
        article.year_date.append(int(row["Year"]))

    if not row.isna()["Keywords"]:
    #separated on , or ; with optional whitespaces around
        list_keywords = re.split("\s?;\s?|\s?,\s?", row["Keywords"])
        for keyword_name in list_keywords:
            keyword_name = keyword_name.strip()
            keyword = onto["keyword"](urllib.parse.quote("keyword_"+keyword_name.replace(" ","_")))
            keyword.label = keyword_name
            article.hasKeyword.append(keyword)

    if not row.isna()["Abstract"]:
        article.abstract = row["Abstract"]
    return article, doi

def per_class_metric_with_extra_info(row, metric, metric_type, process, doi, onto):
    # First split by study area blocks
    study_area_blocks = re.findall(r"\{.*?\}", row.fillna("")[metric])
    if not study_area_blocks:  # fallback to simple semicolon split
        study_area_blocks = re.split(r"\s?;\s?", row.fillna("")[metric])

    for j, block in enumerate(study_area_blocks):
        # Remove enclosing braces if present
        block = block.strip("{} ")
        parts = [part.strip() for part in re.split(r",\s?", block)]
        context_entity = None  # could be study area, dataset, algorithm, date
        # Try to detect the study area from the first part
        # e.g. "Fairfax: Non-residential: 0.78"
        for part in parts:
            fields = [f.strip() for f in part.split(":")]

            if len(fields) == 3:
                context_label, class_name, value = fields
            elif len(fields) == 2:
                class_name, value = fields
                context_label = None
            else:
                print(f"Unrecognized format: {part}")
                continue

            # Disambiguate context_label (study area, dataset, etc.)
            if context_label:
                context_label_lower = context_label.lower()
                if context_label_lower in row.fillna("")["algorithms"].lower():
                    context_entity = next(
                        (algo for algo in process.hasAlgorithm if algo.label[0].lower() == context_label_lower),
                        None
                    )
                elif is_date(context_label):
                    context_entity = context_label
                elif context_label_lower in row.fillna("")["Study Area name"].lower():
                    context_entity = next(
                        (sc for sc in process.hasStudyCase if context_label_lower in sc.label[0].lower()),
                        None
                    )
                elif context_label_lower in row.fillna("")["input data names"].lower():
                    context_entity = context_label  # you can further resolve to dataset object
                elif is_place_name(context_label):
                    context_entity = context_label  # maybe a subregion string
                else:
                    print(f"Cannot disambiguate '{context_label}' — storing as comment.")

            # Create metric
            algo_qual_assessment = onto[metric_type](
                metric.replace(" ", "_") + f"_{j}_{class_name}_{doi}_{int(time.time())}"
            )

            # Class and value
            lulc_class = onto.search_one(iri = f'*lulc_class_{class_name.strip().lower().replace(" ", "_").replace("-", "_")}')
            value = re.sub(r'(?<=\d),(?=\d)', '.', value)
            value = eval(value.replace("%", "/100"))
            algo_qual_assessment.assessedOnClass.append(lulc_class)
            algo_qual_assessment.value.append(value)

            # Link context_entity
            if hasattr(context_entity, "label"):  # likely an ontology object
                if context_entity in process.hasAlgorithm:
                    algo_qual_assessment.applyAccuracyAssessmentOn.append(context_entity)
                elif context_entity in process.hasStudyCase:
                    algo_qual_assessment.hasStudyCase.append(context_entity)
            elif isinstance(context_entity, str):
                if is_date(context_entity):
                    algo_qual_assessment.year_date.append(context_entity)
                elif context_entity in row["input data names"]:
                    algo_qual_assessment.hasValidationDataset.append(context_entity)
                else:
                    algo_qual_assessment.comment = context_entity


            process.hasAccuracyAlgorithm.append(algo_qual_assessment)



def create_article(onto, row):
    print(row)
    article, doi = article_metadata(onto, row)


    #Process
    #We suppose there is one and only one process by row
    defined_processes = [classe.name for classe in onto["process"].descendants()]
    if row["process type"].strip() in defined_processes:
        process = onto[row["process type"].strip()](urllib.parse.quote(
            "process_"+
            row["process type"].strip().replace(" ","_")+
            "_"+doi
            ))
    else:
        print(f"Unknown process type {row['process type']}")
        process = onto["process"](urllib.parse.quote(
            "process_"+
            row["process type"].strip().replace(" ","_")+
            "_"+doi
            ))
    process.label = f'{row["process type"].strip()} {row["Title"]}'
    print("\n________________\n",process.label,"\n________________\n")

    # --- Procedure ---
    if not row.isna()["procedure"]:
        list_procedure = [p.strip() for p in row["procedure"].split(";")]
        procedure_instances = []
        for i, procedure_name in enumerate(list_procedure):
            procedure_instance = onto["procedure"](urllib.parse.quote(
                "procedure_" + procedure_name.replace(" ","_") + "_" + doi
            ))
            procedure_instance.label = procedure_name
            procedure_instance.stepNumber.append(i + 1)
            process.hasProcedure.append(procedure_instance)
            procedure_instances.append(procedure_instance)
    else:
        procedure_instances = []



    # --- Algorithms ---
    algorithms = parse_grouped_field(row.get("algorithms", ""))
    if len(algorithms) == len(procedure_instances):  # Match per step
        for i, algo_group in enumerate(algorithms):
            for algorithm_name in [a.strip() for a in re.split(r"\s?;\s?", algo_group) if a.strip()]:
                algorithm = onto["algorithm"](urllib.parse.quote(algorithm_name.replace(" ","_")))
                algorithm.label = algorithm_name
                procedure_instances[i].hasAlgorithm.append(algorithm)
                process.hasAlgorithm.append(algorithm)  # optional global link
    else:  # Global fallback
        for algorithm_name in re.split(r"\s?;\s?", str(row.get("algorithms", ""))):
            if algorithm_name.strip():
                algorithm = onto["algorithm"](urllib.parse.quote(algorithm_name.replace(" ","_")))
                algorithm.label = algorithm_name
                process.hasAlgorithm.append(algorithm)

    # --- Tools ---
    tool_names = parse_grouped_field(row.get("tool used names", ""))
    tool_types = parse_grouped_field(row.get("tool used types", ""))
    tool_collab = parse_grouped_field(row.get("tool used is collaborative", ""))

    # Normalize tool lists
    max_tools = max(len(tool_names), len(tool_types), len(tool_collab))
    tool_names += ["{}"] * (max_tools - len(tool_names))
    tool_types += ["{}"] * (max_tools - len(tool_types))
    tool_collab += ["{}"] * (max_tools - len(tool_collab))

    if len(tool_names) == len(procedure_instances):  # Match per step
        for i in range(len(procedure_instances)):
            tools_step = [t.strip() for t in re.split(r"\s;\s?", tool_names[i]) if t.strip()]
            types_step = [t.strip() for t in re.split(r"\s;\s?", tool_types[i])] if tool_types[i] else []
            collab_step = [t.strip() for t in re.split(r"\s;\s?", tool_collab[i])] if tool_collab[i] else []

            for j, tool_name in enumerate(tools_step):
                tool_type = types_step[j] if j < len(types_step) else "tool"
                collab = collab_step[j] if j < len(collab_step) else None

                onto_class = onto[tool_type] if tool_type in ["annotation", "storage", "validation"] else onto["tool"]#(No class other tool)
                label = tool_name if tool_type in ["annotation", "storage", "validation", "other"] else f"{tool_name} ({tool_type})"

                tool = onto_class(urllib.parse.quote(f"tool_{tool_name.replace(' ','_')}"))
                tool.label = label
                if collab and collab.lower() in TRUE_VALUES:
                    tool.collaborativeTool.append(True)

                procedure_instances[i].isUsingTool.append(tool)
                process.isUsingTool.append(tool)  # optional global link
    else:
        for i, tool_name in enumerate(re.split(r"\s?;\s?", str(row.get("tool used names", "")))):
            if tool_name.strip():
                tool_type = tool_types[i] if i < len(tool_types) else "tool"
                collab = tool_collab[i] if i < len(tool_collab) else None

                onto_class = onto[tool_type] if tool_type in ["annotation", "storage", "validation"] else onto["tool"]
                label = tool_name if tool_type in ["annotation", "storage", "validation", "other"] else f"{tool_name} ({tool_type})"

                tool = onto_class(urllib.parse.quote(f"tool_{tool_name.replace(' ','_')}"))
                tool.label = label
                if collab and collab.lower() in TRUE_VALUES:
                    tool.collaborativeTool.append(True)

                process.isUsingTool.append(tool)


    #INPUT DATA
    defined_natures = [classe.name for classe in onto["spatial_data"].descendants()]

    list_inputs_is_training = None
    list_inputs_instances = []
    if not row.isna()["input data names"]:
        #Names
        list_inputs = re.split(r"\s?;\s?", row["input data names"])
        #Natures
        if not row.isna()["input data natures and resolution"]:
            list_inputs_nature = re.split(r"\s?;\s?", row["input data natures and resolution"])
            list_inputs_resolution = []
            for i in range(len(list_inputs_nature)):
                if ":" in list_inputs_nature[i]:
                    list_inputs_nature[i], resolution = re.split(r"\s?:\s?", list_inputs_nature[i])
                    list_inputs_resolution.append(resolution)
                else:
                    list_inputs_resolution.append(None)
        else:
            list_inputs_nature = [None] * len(list_inputs)
            list_inputs_resolution = [None] * len(list_inputs)
        #Date
        if not row.isna()["input data date"]:
                list_inputs_date = re.split(r"\s?;\s?",
                                           (row["input data date"]).lower()
                                           )
        else:
            list_inputs_date = [None] * len(list_inputs)
        #Is VGI?
        if not row.isna()["Input  is VGI "]:
            list_inputs_vgi = re.split(r"\s?;\s?",
                                       (row["Input  is VGI "]).lower()
                                       )
        else:
            list_inputs_vgi = ["None"] * len(list_inputs)
        #raster/vecter
        if not row.isna()["input data raster/points/lines/polygon"]:
            list_inputs_raster_vector = re.split("\s?;\s?",
                                       (row["input data raster/points/lines/polygon"]).lower()
                                       )
        else:
            list_inputs_raster_vector = [""] * len(list_inputs)

        #training, validation
        if not row.isna()["input is training, validation, both or neither"]:
            list_inputs_is_training = re.split(r"\s?;\s?",
                                       (row["input is training, validation, both or neither"]).lower()
                                       )
        else:
            list_inputs_is_training = [""] * len(list_inputs)
        #training size
        if not row.isna()["training dataset size"]:
            list_inputs_training_size = re.split(r"\s?;\s?",
                                       row["training dataset size"]
                                       )
        else:
            list_inputs_training_size = []
        #validation size
        if not row.isna()["validation dataset size"]:
            list_inputs_validation_size = re.split(r"\s?;\s?",
                                       row["validation dataset size"]
                                       )
        else:
            list_inputs_validation_size = []
        print(list_inputs_nature)
        index_training_dataset = 0
        index_validation_dataset = 0
        for i, input_name in enumerate(list_inputs):
            if list_inputs_nature[i].replace(" ", "_").lower() in [nature.lower() for nature in defined_natures]:
                n_defined_nature = [nature.lower() for nature in defined_natures].index(list_inputs_nature[i].replace(" ", "_").lower())
                nature = defined_natures[n_defined_nature]
            else:
                nature = "data"#Is it possible that non spatial data can be used ?
            print(nature)
            input_instance = onto[nature](urllib.parse.quote(
                input_name.replace(" ","_") + "_" + nature + "_" + doi
                ))
            input_instance.is_a.append(onto["input_data"])
            input_instance.label = input_name
            if list_inputs_resolution[i] is not None:
                input_instance.resolution.append(list_inputs_resolution[i])
            if list_inputs_vgi[i] is not None:
                input_instance.volunteered_geographic_information.append(
                    list_inputs_vgi[i].lower() in TRUE_VALUES
                    )
            if len(list_inputs_raster_vector)>i and list_inputs_raster_vector[i] is not None:
                input_instance.pixelGeometricRepresentation.append(
                    ("raster" in list_inputs_raster_vector[i]) or ("pixel" in list_inputs_raster_vector[i])
                    )
                input_instance.pointGeometricRepresentation.append(
                    "point" in list_inputs_raster_vector[i]
                    )
                input_instance.polygonGeometricRepresentation.append(
                    "polygon" in list_inputs_raster_vector[i]
                    )
                input_instance.lineGeometricRepresentation.append(
                    "line" in list_inputs_raster_vector[i]
                    )
            if len(list_inputs_date)>i and list_inputs_date[i] is not None and list_inputs_date[i]!='':
                if "period" in list_inputs_date[i]:
                    period = onto["period"](urllib.parse.quote(list_inputs_date[i].strip().replace(" ", "_").replace("-", "_").replace(",", "_")))
                    dates = re.split(r"\s?,\s?|\s?-\s?", list_inputs_date[i].replace("period(", "").replace(")", "") )
                    period.year_date.extend(dates)
                    input_instance.interval_date.append(
                        period
                        )
                if "[" in list_inputs_date[i]:
                    dates = re.split(r"\s?,\s?", list_inputs_date[i].replace("[", "").replace("]", "") )
                    input_instance.year_date.extend(dates)
                elif list_inputs_date[i].isnumeric():
                    input_instance.year_date.append(int(list_inputs_date[i]))

            process.hasInput.append(input_instance)
            list_inputs_instances.append(input_instance)
            if list_inputs_is_training[i] in ["training", "both"]:
                input_instance.is_a.append(onto["training_dataset"])
                process.hasTrainingDataset.append(input_instance)
                if(len(list_inputs_training_size)>index_training_dataset):
                    input_instance.datasetSize.append( list_inputs_training_size[index_training_dataset] )
                else:
                    print("It lacks a training dataset size")
                index_training_dataset +=1

            if list_inputs_is_training[i] in ["validation", "both"]:
                input_instance.is_a.append(onto["validation_dataset"])
                process.hasValidationDataset.append(input_instance)
                if(len(list_inputs_validation_size)>index_validation_dataset):
                    input_instance.datasetSize.append( list_inputs_validation_size[index_validation_dataset] )
                else:
                    print("It lacks a validation dataset size")
                index_validation_dataset +=1
            print(nature)
            if nature in ['land_use', 'land_cover', 'land_use_land_cover', 'building']:
                if not row.isna()["if classification, nomenclature classes"] or not row.isna()["if classification, nomenclature level"]:
                    if row.isna()["if classification, nomenclature level"]:
                        #we have to infer it by counting the pipes "|" in "if classification, nomenclature classes"
                        number_nomenclatures = row["if classification, nomenclature classes"].count("|") + 1
                        levels = [1]*number_nomenclatures#we suppose that they are all 1 level nomenclatures
                        print(levels)
                    else:
                        levels = re.split(r"\s?;\s?", row.fillna("")["if classification, nomenclature level"])
                        print(levels)
                        number_nomenclatures = len(levels)
                    lu_or_lc = {"land_use":"lu", "land_cover":"lc", "land_use_land_cover":"lulc", "building":"lu"}[nature]
                    if row.isna()["if classification, nomenclature name"]:
                        nomenclature_names = [f"{doi}_{lu_or_lc}_nomenclature_level_{levels[k]}_{k}" for k in range(number_nomenclatures)]
                    else:
                        nomenclature_names = re.split(r"\s?;\s?",  row.fillna("")["if classification, nomenclature name"])
                        nomenclature_names = [name if name!="" else f"{doi}_{lu_or_lc}_nomenclature_level_{levels[k]}" for k, name in enumerate(nomenclature_names)]
                    nomenclature_classes_groups = re.split(r"\s?\|\s?", row["if classification, nomenclature classes"])
                    print(nomenclature_classes_groups)
                    #We suppose that the nomenclature is hierarchical if there are two nomenclatures with increasing level
                    hierarchical_nomenclature = ( number_nomenclatures>1 and (int(levels[0])+1 == int(levels[1])) )
                    all_classes = []
                    for k in range(number_nomenclatures):
                        level = levels[k]
                        nomenclature_name = nomenclature_names[k]
                        print(nomenclature_name, f"level{level}_{'mixed_lu_and_lc' if lu_or_lc == 'lulc' else lu_or_lc}_nomenclature")
                        nomenclature_instance = onto[f"level{level}_{'mixed_lu_and_lc' if lu_or_lc == 'lulc' else lu_or_lc}_nomenclature"](urllib.parse.quote(
                            nomenclature_name.replace(" ", "_") + "_" + doi
                            ))
                        input_instance.hasNomenclature.append(nomenclature_instance)
                        nomenclature_instance.label = nomenclature_name
                        all_classes.append([])
                        nomenclature_k_classes = re.split(r"\s?;\s?", nomenclature_classes_groups[k].replace("(", "").replace(")", "") )
                        if hierarchical_nomenclature and k>0:
                            group_hierarchy = get_group_hierarchy(nomenclature_classes_groups[k])
                        for class_number, class_name in enumerate(nomenclature_k_classes):
                            class_instance = onto[f"{lu_or_lc}_class"](urllib.parse.quote(
                                "lulc_class_"+class_name.strip().lower().replace(" ", "_").replace("-", "_")))
                            class_instance.label = class_name
                            nomenclature_instance.hasLULCClass.append(class_instance)
                            if hierarchical_nomenclature and k>0:
                                mother_class = all_classes[k-1][group_hierarchy[class_number]]
                                class_instance.isALandUseOrLandCoverSubclassOf.append(mother_class)
                            all_classes[k].append(class_instance)

    all_classes = [nomenclature.hasLULCClass for nomenclature in list(set([nomenclature for input_instance in process.hasInput for nomenclature in input_instance.hasNomenclature]))]
    #OUTPUT DATA
    if not row.isna()["output data names"]:
        #Names
        list_outputs = re.split(r"\s?;\s?", row["output data names"])
        #Natures
        if not row.isna()["output data natures and resolution"]:
            list_outputs_nature = re.split(r"\s?;\s?", row["output data natures and resolution"])
            list_outputs_resolution = []
            for i in range(len(list_outputs_nature)):
                if ":" in list_outputs_nature[i]:
                    list_outputs_nature[i], resolution = re.split(r"\s?:\s?", list_outputs_nature[i])
                    list_outputs_resolution.append(resolution)
                else:
                    list_outputs_resolution.append(None)
        else:
            list_outputs_nature = [None] * len(list_outputs)
        #raster/vector
        if not row.isna()["output data raster/points/lines/polygon"]:
            list_outputs_raster_vector = re.split(r"\s?;\s?",
                                       row["output data raster/points/lines/polygon"]
                                       )
        else:
            list_outputs_raster_vector = [""] * len(list_outputs)


        for i, output_name in enumerate(list_outputs):
            nature_i = list_outputs_nature[i]
            if nature_i.replace(" ", "_").lower() not in [nature.lower() for nature in defined_natures]:
                nature_i = "data"
            else:
                n_defined_nature = [nature.lower() for nature in defined_natures].index(nature_i.replace(" ", "_").lower())
                nature_i = defined_natures[n_defined_nature]
            output_instance = onto[nature_i](urllib.parse.quote(
                output_name.replace(" ","_") + "_" + nature_i + "_" + doi
                ))
            output_instance.is_a.append(onto["output_data"])
            output_instance.label = output_name
            if list_outputs_resolution[i] is not None:
                output_instance.resolution.append(list_outputs_resolution[i])
            output_instance.pixelGeometricRepresentation.append(
                "raster" in list_outputs_raster_vector[i]
                )
            output_instance.pointGeometricRepresentation.append(
                "points" in list_outputs_raster_vector[i]
                )
            output_instance.polygonGeometricRepresentation.append(
                "polygon" in list_outputs_raster_vector[i]
                )
            output_instance.lineGeometricRepresentation.append(
                "lines" in list_outputs_raster_vector[i]
                )
            process.hasOutput.append(output_instance)


    #operator
    if not row.isna()["operator type"]:
        operators_types = re.split(r"\s?;\s?", row["operator type"])
        operators_infos = re.split(r"\s?;\s?", row["operator description"])
        for i, operator_type in enumerate(operators_types):
            if operator_type.strip() not in ["person", "computer"]:
                operator_type = "operator"
            operator = onto[operator_type](urllib.parse.quote(
                operator_type + str(i) + "_" + doi
                ))
            operator.label = operators_infos[i]
            process.hasOperator.append(
                operator
                )

    #Study case
    study_cases = []
    study_cases_instances = []
    if not row.isna()["Study Area name"]:
        study_cases = re.split(r"\s?;\s?", row["Study Area name"])
        if not row.isna()["belongs to country"]:
            countries = re.split(r"\s?;\s?", row["belongs to country"])
        else:
            countries = [""]*len(study_cases)
        extents_types = re.split(r"\s?;\s?", row["geographic extent type"])
        for i, study_area in enumerate(study_cases):
            extent_type = extents_types[i].lower()
            if extent_type == "state":
                extent_type = "local"
            elif extent_type not in ["local", "regional", "national", "global"]:
                extent_type = "geographic_extent"
            study_case = onto[extent_type](urllib.parse.quote(
                "study_case_"+study_area.replace(" ","_")
                ))
            study_case.label = study_area
            if "[" in countries[i]:
                list_countries_i = re.split(r"\s?,\s?", countries[i].replace("[","").replace("]","") )
                study_case.belongsToCountry.extend(list_countries_i)
            else :
                study_case.belongsToCountry.append( countries[i] )
            study_cases_instances.append(study_case)
            process.hasStudyCase.append(study_case)

    # Quality assessment
    computed = "computed"
    ## Global metrics
    list_global_quality_metrics = [
        'OA',
        'mF1',
        'mIoU',
        'kappa',
        'global recall (producer accuracy)',
        'global precision (user accuracy)'
        ]
    metrics_type = [
        "overall_accuracy",
        "f1_score",
        "intersection_over_union",
        "separated_kappa",
        "producer_accuracy",
        "user_accuracy"
        ]

    # Count the number of validation datasets and study areas
    num_validation_datasets = len(list_inputs_is_training) if list_inputs_is_training else 0
    num_study_areas = len(study_cases) if study_cases else 0

    for i, metric in enumerate(list_global_quality_metrics):
        if not row.isna()[metric]:
            metric_values = re.split(r"\s?;\s?", row[metric])
            for j, metric_value in enumerate(metric_values):
                algo_qual_assessment = onto[metrics_type[i]](urllib.parse.quote(
                    metric.replace(" ", "_")  + "_" + str(j) + "_" + str(doi) + "_" + str(int(time.time()))
                    ))
                #verify if the decimal numbers are written with dots and not with commas like in French
                metric_value = re.sub(r'(?<=\d),(?=\d)', '.', metric_value)
                if not is_number(metric_value.replace("%", "").strip()):#It is not simply the value of the metric
                    if "(" in metric_value:
                        metric_value, comment = re.split(r"\s?\(\s?", metric_value)
                        algo_qual_assessment.comment += comment.replace(")","").strip()
                    if ":" in metric_value:#if there are text: values
                        print(metric_value)
                        metric_text, metric_value = metric_value.split(":")
                        metric_text = metric_text.strip().lower()
                    else:#
                        metric_value, metric_text = split_by_decimal_token(metric_value)
                        metric_text = metric_text.strip().lower()
                        if metric_text.startswith("of"):
                            metric_text = metric_text[3:]
                    #In the excel filled by the other, metric_text can refer to a date, a study area, a dataset, an algorithm
                    if metric_text in row.fillna("")["algorithms"].lower():
                        print(metric_text,"is assumed to be the algorith used")
                        for algo in process.hasAlgorithm:
                            if algo.label[0].lower()==metric_text:
                                algo_qual_assessment.applyAccuracyAssessmentOn.append(algo)
                    elif is_date(metric_text):
                        print(metric_text,"is assumed to be the date")
                        algo_qual_assessment.year_date.append(metric_text)
                    elif metric_text in row.fillna("")["Study Area name"].lower():
                        print(metric_text,"is assumed to be the study case")
                        #find which study case it is
                        for study_case in process.hasStudyCase:
                            if study_case.label[0].lower() == metric_text:
                                algo_qual_assessment.hasStudyCase.append(study_case)
                    elif metric_text=="global" and len(process.hasStudyCase)==1:
                        algo_qual_assessment.hasStudyCase.append(process.hasStudyCase[0])
                    elif metric_text in row.fillna("")["input data names"].lower():
                        print(metric_text,"is assumed to be the dataset")
                        for input_data in process.hasInput:
                            if input_data.label[0].lower() == metric_text:
                                algo_qual_assessment.hasValidationDataset.append(input_data)
                    elif is_place_name(metric_text):
                        print(metric_text,"is assumed to be a subpart of the study area")
                        #we have to create an object study area
                        study_case = onto["geographic_extent"](urllib.parse.quote(
                            "study_case_"+metric_text.replace(" ","_")
                            ))
                        study_case.label = metric_text
                        algo_qual_assessment.hasStudyCase.append(study_case)
                    else:
                        print(f"Can't guess what '{metric_text}' refers to")
                        algo_qual_assessment.comment = metric_text
                elif num_validation_datasets == num_study_areas and num_validation_datasets == len(metric_values):
                    # Assume each validation dataset corresponds to a study area
                    algo_qual_assessment.hasValidationDataset.append(list_inputs_instances[j])
                    algo_qual_assessment.hasStudyCase.append(process.hasStudyCase[j])
                elif num_validation_datasets == len(metric_values):
                    # Assume each metric value corresponds to a validation dataset
                    algo_qual_assessment.hasValidationDataset.append(list_inputs_instances[j])
                elif num_study_areas == len(metric_values):
                    # Assume each metric value corresponds to a study area
                    algo_qual_assessment.hasStudyCase.append(study_cases_instances[j])
                else:
                    print(f"Mismatch in the number of validation datasets, study areas, and metric values for {metric}")

                algo_qual_assessment.value.append(
                    eval(metric_value.replace("%", "/100"))
                    )
                process.hasAccuracyAlgorithm.append(algo_qual_assessment)

    ## Per class metrics
    list_per_class_quality_metrics = [
        'per class binary accuracy',
        'per class F1 score',
        'per class IoU',
        'per class recall (producer accuracy)',
        'per class precision (user accuracy)'
        ]
    metrics_type = [
        "algorithm_quality_assessment",
        "f1_score",
        "intersection_over_union",
        "producer_accuracy",
        "user_accuracy"
        ]

    all_classes_flat = [class_i  for class_group in all_classes for class_i in class_group]
    for i, metric in enumerate(list_per_class_quality_metrics):
        if not row.isna()[metric]:
          if row[metric].lower().strip() == "computed":
            for lulc_class in all_classes_flat:
                algo_qual_assessment = onto[metrics_type[i]](urllib.parse.quote(
                    metric.replace(" ", "_") + "_" + str(doi) + "_" + str(int(time.time()))
                    ))
                algo_qual_assessment.assessedOnClass.append(lulc_class)
                process.hasAccuracyAlgorithm.append(algo_qual_assessment)
          else:
            if "{" in row[metric]:
                per_class_metric_with_extra_info(row, metric, metrics_type[i], process, doi, onto)
            else:
                metric_values = re.split(r"\s?;\s?", row[metric])
                for j, metric_value in enumerate(metric_values):
                    algo_qual_assessment = onto[metrics_type[i]](urllib.parse.quote(
                        metric.replace(" ", "_") + "_" + str(j) + "_" + str(doi) + "_" + str(int(time.time()))
                        ))
                    if ":" not in metric_value:

                        if len(all_classes_flat) == 1:
                            lulc_class = all_classes_flat[0]
                            value = metric_value
                        elif len(all_classes_flat) == len(metric_values):
                            lulc_class = all_classes_flat[j]
                            value = metric_value
                        else:
                            print( f"class name not provided for {metric} and cannot be infered")
                            break
                    else:
                        lulc_class_name, value = re.split(r"\s?:\s?", metric_value)
                        print(lulc_class_name)
                        lulc_class = onto.search_one(iri = "*lulc_class_" + lulc_class_name.strip().lower().replace(" ", "_").replace("-", "_"))
                        if not lulc_class:
                            print(f"LULC class {lulc_class_name} not found")
                            lulc_class = onto["lulc_class"](urllib.parse.quote(
                                "lulc_class_" + lulc_class_name.replace("-", "_").replace(" ", "_")))
                            lulc_class.label = lulc_class_name
                    #verify if the decimal numbers are written with dots and not with commas like in French
                    value = re.sub(r'(?<=\d),(?=\d)', '.', value)
                    if "(" in value:
                        value, comment = re.split(r"\s?\(\s?", value)
                        algo_qual_assessment.comment = comment.replace(")","").strip()
                    print(lulc_class, type(lulc_class))
                    algo_qual_assessment.assessedOnClass.append(lulc_class)
                    algo_qual_assessment.value.append(
                        eval(value.replace("%", "/100"))
                        )
                    print(num_validation_datasets, num_study_areas)
                    if num_validation_datasets == num_study_areas:
                        # Assume each validation dataset corresponds to a study area
                        algo_qual_assessment.hasValidationDataset.append(list_inputs_instances[j])
                        algo_qual_assessment.hasStudyCase.append(study_cases_instances[j])
                    elif num_validation_datasets == len(metric_values):
                        # Assume each metric value corresponds to a validation dataset
                        algo_qual_assessment.hasValidationDataset.append(list_inputs_instances[j])
                    elif num_study_areas == len(metric_values):
                        # Assume each metric value corresponds to a study area
                        algo_qual_assessment.hasStudyCase.append(study_cases_instances[j])
                    else:
                        print(f"Mismatch in the number of validation datasets, study areas, and metric values for {metric}")
                    process.hasAccuracyAlgorithm.append(algo_qual_assessment)

    if not row.isna()["user defined algorithm quality assessment metrics"]:
        other_metrics = re.split(r"\s?;\s?", row["user defined algorithm quality assessment metrics"])
        for j, metric in enumerate(other_metrics):
            if ":" in metric or "=" in metric:
                metric_name_and_class, metric_value = re.split(r"\s?:|=\s?", metric)
            else:
                metric_name_and_class = metric
                metric_value = None
            if "(" in metric_name_and_class:
                # Escape the parenthesis in the regex pattern
                metric_name, lulc_class_name = re.split(r"\s?\(\s?", metric_name_and_class.replace(")", ""))
                lulc_class_name = lulc_class_name.strip()
                print("onto.lulc_class_" + lulc_class_name.replace("-", "_").replace(" ", "_"))
                lulc_class = lulc_class = onto.search_one(iri = "*lulc_class_" + lulc_class_name.replace("-", "_").replace(" ", "_"))
                if not lulc_class:
                    print(f"LULC class {lulc_class_name} not found")
                    lulc_class = onto["lulc_class"]("lulc_class_" + lulc_class_name.replace("-", "_").replace(" ", "_"))
                    lulc_class.label = lulc_class_name
            else:
                metric_name = metric_name_and_class
            algo_qual_assessment = onto["algorithm_quality_assessment"](urllib.parse.quote(
                metric_name.replace(" ", "_") + "_" + str(j) + "_" + doi + "_" + str(int(time.time())))
                )
            algo_qual_assessment.label = metric_name
            if "(" in metric_name_and_class:
                algo_qual_assessment.assessedOnClass.append(lulc_class)
            if metric_value is not None:
                #verify if the decimal numbers are written with dots and not with commas like in French
                metric_value = re.sub(r'(?<=\d),(?=\d)', '.', metric_value)
                try:
                    algo_qual_assessment.value.append(
                            eval(metric_value.replace("%", "/100"))
                            )
                except:
                    algo_qual_assessment.value.append(
                            metric_value
                            )
            other_similar_metrics = [metric_name_and_class in metric for metric in other_metrics]
            if num_validation_datasets == num_study_areas:
                # Assume each validation dataset corresponds to a study area
                algo_qual_assessment.hasValidationDataset.append(list_inputs[j])
                algo_qual_assessment.hasStudyCase.append(study_cases[j])
            elif num_validation_datasets == len(other_similar_metrics):
                # Assume each metric value corresponds to a validation dataset
                algo_qual_assessment.hasValidationDataset.append(list_inputs[j])
            elif num_study_areas == len(other_similar_metrics):
                # Assume each metric value corresponds to a study area
                algo_qual_assessment.hasStudyCase.append(study_cases_instances[j])
            else:
                print(f"Mismatch in the number of validation datasets, study areas, and metric values for {metric}")
            process.hasAccuracyAlgorithm.append(algo_qual_assessment)

    #criterions
    if row.fillna("")["codeAvailability "].lower() not in(FALSE_VALUES):
        process.codeAvailability.append(row["codeAvailability "])
    if row.fillna("")["dataAvailability"].lower() not in(FALSE_VALUES):
        process.dataAvailability.append(row["dataAvailability"])
    if not row.isna()["challenge"]:
        process.challenge.extend(re.split(r"\s?;\s?", row["challenge"]))
    if not row.isna()["strength"]:
        process.strength.extend(re.split(r"\s?;\s?", row["strength"]))
    if not row.isna()["weakness"]:
        process.weaknesses.extend(re.split(r"\s?;\s?", row["weakness"]))

    article.hasProcess.append(process)


    return article

def get_excel_files(path: str):
    if os.path.isdir(path):
        # List all Excel files in the directory
        return [os.path.join(path, f) for f in os.listdir(path) if f.endswith(('.xls', '.xlsx', '.xlsm'))]
    elif os.path.isfile(path) and path.endswith(('.xls', '.xlsx', '.xlsm')):
        # If it's an Excel file, return a list containing only this file
        return [path]
    else:
        # If it's neither a directory nor an Excel file, return an empty list
        return []

#%%
if __name__ == "__main__":

    #path to the owl file defining the ontology
    owl_file_path = os.path.join(
        "lulc_review.owl"
        )
    
    onto = or2.get_ontology(owl_file_path).load()

    #path to an excel file describing articles to instantiate, or a folder of excel files
    excel_ontology_folder_path = "LULC_Ontology_example.xlsm" if len(sys.argv) <= 1 else sys.argv[1]
    
    list_excel_files_path = get_excel_files(excel_ontology_folder_path)
    
    for excel_ontology_file_path in list_excel_files_path:
        print("\n______________________________\n",excel_ontology_file_path)
        excel_ontology_file = pd.read_excel(excel_ontology_file_path, dtype=str, sheet_name="ontology_instanciation", header=[0,1])
        excel_ontology_file = enrich_metadata(excel_ontology_file)
        excel_ontology_file.to_csv("enriched_excel.csv", index=False)
        #break
        excel_ontology_file.columns = excel_ontology_file.columns.droplevel(0)
        excel_ontology_file.drop(index=excel_ontology_file.index[0], axis=0, inplace=True)#The purpose of this row is to help the user on how to fill each column
        display(excel_ontology_file)

        for i in range(1, len(excel_ontology_file)):
            row = excel_ontology_file.iloc[i]
            if ignore_error:#If the article cannot be instantiated, an error message is displayed, but the other papers of the folder can be instantiated
                try:
                    article = create_article(onto, row)
                except Exception as e:
                    print("\n--------------------------------\n", "Exception:\n",e, "\n--------------------------------\n")
            else:#Stop on error
                article = create_article(onto, row)
            #visualize_instance(article)

    onto.save(
        os.path.join(
            "lulc_review_instantiated.owl"
            )
        )
