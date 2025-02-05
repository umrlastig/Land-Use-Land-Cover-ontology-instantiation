# -*- coding: utf-8 -*-
"""
Created on Fri Nov 24 11:58:31 2023

@author: MCubaud
"""
import os
import pandas as pd
import numpy as np
import owlready2  as or2
import re
from typing import Union, Optional, Iterable
import pylatexenc.latex2text

TRUE_VALUES = ["yes", "true", "1", "t", "y", "0b1", "√", "☑", "✔"]
FALSE_VALUES = ["", None, "no", "false", "0", "f", "n", "w"]


def decode_latex(latex_string):
    try:
        decoded_string = pylatexenc.latex2text.LatexNodes2Text().latex2text(latex_string)
        return decoded_string
    except Exception as e:
        print("Error:", e)
        return latex_string
    
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
    article = onto["paper"](doi)
    article.doi = doi

    article.title = row["Title"]
    article.label = row["Title"]

    if not row.isna()["journal"]:
        journal = onto[row["type of publication"]](
            row["type of publication"]+
            "_"+
            row["journal"].replace(" ","_")
            )
        journal.namePublisher.append(row["journal"])
        journal.label = row["journal"]
        article.isPublishedIn.append(journal)

    #affiliations
    affiliation_instances = []
    if not row.isna()["Affiliation Name"]:
        list_affiliation_names = re.split("\s?;\s?", row["Affiliation Name"])
        list_affiliation_addresses = re.split("\s?;\s?", row["Affiliation Address"])
        for i, affiliation_name in enumerate(list_affiliation_names):
            affiliation = onto["affiliation"](
                affiliation_name.replace(" ","_").replace(",","")+"_"+
                list_affiliation_addresses[i].replace(" ","_").replace(",","")
                )
            affiliation.label = affiliation_name
            affiliation.affiliationName.append(affiliation_name)
            affiliation.affiliationAddress.append(list_affiliation_addresses[i])
            affiliation_instances.append(affiliation)

    if not row.isna()["Authors"]:
        list_authors = row["Authors"].split(" and ")
        if len(affiliation_instances)==1:
            #Then we suppose all the authors have the same affiliation
            affiliation_instances = affiliation_instances * len(list_authors)
        for i, name_author in enumerate(list_authors):
            name_author = decode_latex(name_author)
            author = onto["author"](re.sub("\s?,\s?", "_", name_author))
            author.label = name_author
            author.lastName, author.firstName = re.split("\s?,\s?", name_author)
            if len(affiliation_instances)!=0:
                author.hasAffiliation.append(affiliation_instances[i])
            article.hasAuthor.append(author)

    if not row.isna()["Year"]:
        article.year_date = int(row["Year"])
    
    if not row.isna()["Keywords"]:
    #separated on , or ; with optional whitespaces around
        list_keywords = re.split("\s?;\s?|\s?,\s?", row["Keywords"])
        for keyword_name in list_keywords:
            keyword_name = keyword_name.strip()
            keyword = onto["keyword"]("keyword_"+keyword_name.replace(" ","_"))
            keyword.label = keyword_name
            article.hasKeyword.append(keyword)

    if not row.isna()["Abstract"]:
        article.abstract = row["Abstract"]
    return article, doi

def create_article(onto, row):

    article, doi = article_metadata(onto, row)




    #Process
    #We suppose there is one and only one process by row
    process = onto[row["process type"].strip()](
            "process_"+
            row["process type"].strip().replace(" ","_")+
            "_"+doi
            )
    process.label = f'{row["process type"]} {row["Title"]}'

    #Procedure
    if not row.isna()["procedure"]:
        list_procedure = row["procedure"].split(";")
        for i, procedure_name in enumerate(list_procedure):
            procedure_name = procedure_name.strip()
            procedure_instance = onto["procedure"](
                "procedure_" + procedure_name.replace(" ","_") + "_" + doi
                )
            procedure_instance.label = procedure_name
            procedure_instance.stepNumber.append(i+1)
            process.hasProcedure.append(procedure_instance)


    #instruments
    #algorithms
    if not row.isna()["algorithms"]:
        for algorithm_name in re.split("\s?;\s?",row["algorithms"]):
            algorithm = onto["algorithm"](algorithm_name.replace(" ","_"))
            algorithm.label = algorithm_name
            process.hasInstrument.append(algorithm)
            
    #tools
    if not row.isna()["tool used names"]:
        tool_used_names = re.split("\s?;\s?",row["tool used names"])
        tool_used_types = re.split("\s?;\s?",row["tool used types"])
        if not row.isna()["tool used names"]:
            tool_used_collab = re.split("\s?;\s?", row["tool used is collaborative"])
        else:
            tool_used_collab = [None] * len(tool_used_names)
        for i, tool_name in enumerate(tool_used_names):
            tool = onto[tool_used_types[i]](
                f"tool_{tool_name.replace(' ','_')}")
            tool.label = tool_name
            if tool_used_collab[i] is not None:
                tool.collaborativeTool.append(
                    tool_used_collab[i].lower() in TRUE_VALUES
                    )
            process.isUsingTool.append(tool)


    
    #INPUT DATA
    defined_natures = [classe.name for classe in onto["spatial_data"].descendants()]

    if not row.isna()["input data names"]:
        #Names
        list_inputs = re.split("\s?;\s?", row["input data names"])
        #Natures
        if not row.isna()["input data natures and resolution"]:
            list_inputs_nature = re.split("\s?;\s?", row["input data natures and resolution"])
            list_inputs_resolution = []
            for i in range(len(list_inputs_nature)):
                if ":" in list_inputs_nature[i]:
                    list_inputs_nature[i], resolution = re.split("\s?:\s?", list_inputs_nature[i])
                    list_inputs_resolution.append(resolution)
                else:
                    list_inputs_resolution.append(None)
        else:
            list_inputs_nature = [None] * len(list_inputs)
            list_inputs_resolution = [None] * len(list_inputs)
        #Date
        if not row.isna()["input data date"]:
                list_inputs_date = re.split("\s?;\s?", 
                                           row["input data date"]
                                           )
        else:
            list_inputs_date = [None] * len(list_inputs)
        #Is VGI?
        if not row.isna()["Input  is VGI "]:
            list_inputs_vgi = re.split("\s?;\s?", 
                                       row["Input  is VGI "]
                                       )
        else:
            list_inputs_vgi = [None] * len(list_inputs)
        #raster/vecter
        if not row.isna()["input data raster/points/lines/polygon"]:
            list_inputs_raster_vector = re.split("\s?;\s?", 
                                       row["input data raster/points/lines/polygon"]
                                       )
        else:
            list_inputs_raster_vector = [""] * len(list_inputs)

        #training, validation
        if not row.isna()["input is training, validation, both or neither"]:
            list_inputs_is_training = re.split("\s?;\s?", 
                                       row["input is training, validation, both or neither"]
                                       )
        else:
            list_inputs_is_training = [""] * len(list_inputs)
        #training size
        if not row.isna()["training dataset size"]:
            list_inputs_training_size = re.split("\s?;\s?", 
                                       row["training dataset size"]
                                       )
        else:
            list_inputs_training_size = []
        #validation size
        if not row.isna()["validation dataset size"]:
            list_inputs_validation_size = re.split("\s?;\s?", 
                                       row["validation dataset size"]
                                       )
        else:
            list_inputs_validation_size = []

        index_training_dataset = 0
        index_validation_dataset = 0
        for i, input_name in enumerate(list_inputs):
            if list_inputs_nature[i] in defined_natures:
                nature = list_inputs_nature[i]
            else:
                nature = "data"#Is it possible that non spatial data can be used ?
            input_instance = onto[nature](
                input_name.replace(" ","_") + "_" + nature + "_" + doi
                )
            input_instance.is_a.append(onto["input_data"])
            input_instance.label = input_name
            if list_inputs_resolution[i] is not None:
                input_instance.minimum_mapping_unit.append(list_inputs_resolution[i])
            input_instance.volunteered_geographic_information.append(
                list_inputs_vgi[i].lower in TRUE_VALUES
                )
            input_instance.pixelGeometricRepresentation.append(
                "raster" in list_inputs_raster_vector[i]
                )
            input_instance.pointGeometricRepresentation.append(
                "points" in list_inputs_raster_vector[i]
                )
            input_instance.polygonGeometricRepresentation.append(
                "polygon" in list_inputs_raster_vector[i]
                )
            input_instance.lineGeometricRepresentation.append(
                "lines" in list_inputs_raster_vector[i]
                )
            if list_inputs_date[i] is not None and list_inputs_date[i]!='':
                if "period" in list_inputs_date[i]:
                    period = onto["period"]
                    dates = re.split("\s?,\s?|\s?-\s?", list_inputs_date[i].replace("period(", "").replace(")", "") )
                    period.year_date.extend(dates)
                    input_instance.interval_date.append(
                        period
                        )
                if "[" in list_inputs_date[i]:
                    dates = re.split("\s?,\s?", list_inputs_date[i].replace("[", "").replace("]", "") )
                    input_instance.year_date.extend(dates)
                elif list_inputs_date[i].isnumeric():
                    input_instance.year_date = int(list_inputs_date[i])
                    
            process.hasInput.append(input_instance)
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
                if nature in ['land_use', 'land_cover', 'land_use_land_cover']:
                    if not row.isna()["if classification, nomenclature level"]:
                        levels = re.split("\s?;\s?", row["if classification, nomenclature level"])
                        number_nomenclatures = len(levels)
                        lu_or_lc = {"land_use":"lu", "land_cover":"lc", "land_use_land_cover":"lulc"}[nature_i]
                        nomenclature_names = re.split("\s?;\s?",  row["if classification, nomenclature name"])
                        nomenclature_names = [name if name!="" else f"{doi}_{lu_or_lc}_nomenclature_level_{levels[k]}" for k, name in enumerate(nomenclature_names)]
                        nomenclature_classes_groups = re.split("\s?|\s?", row["if classification, nomenclature classes"])
                        #We suppose that the nomenclature is hierarchical if there are two nomenclatures with increasing level
                        hierarchical_nomenclature = ( number_nomenclatures>1 and (int(levels[0])+1 == int(levels[1])) )
                        all_classes = []
                        for k in range(number_nomenclatures):
                            level = levels[k]
                            nomenclature_name = nomenclature_names[k]
                            nomenclature_instance = onto[f"level{level}_{lu_or_lc}_nomenclature"](
                                nomenclature_name.replace(" ", "_") + "_" + doi
                                )
                            input_instance.hasNomenclature.append(nomenclature_instance)
                            nomenclature_instance.label = nomenclature_name
                            all_classes.append([])
                            nomenclature_k_classes = re.split("\s?;\s?", nomenclature_classes_groups[k].replace("(", "").replace(")", "") )
                            if hierarchical_nomenclature and k>0:
                                group_hierarchy = get_group_hierarchy(nomenclature_classes_groups[k])
                            for class_number, class_name in enumerate(nomenclature_k_classes):
                                class_instance = onto[f"{lu_or_lc}_class"]("lulc_class_"+class_name.replace(" ", "_"))
                                class_instance.label = class_name
                                nomenclature_instance.hasLULCClass.append(class_instance)
                                if hierarchical_nomenclature and k>0:
                                    mother_class = all_classes[k-1][group_hierarchy[class_number]]
                                    class_instance.isALandUseOrLandCoverSubclassOf.append(mother_class)
                                all_classes[k].append(class_instance)
            

    #OUTPUT DATA
    if not row.isna()["output data names"]:
        #Names
        list_outputs = re.split("\s?;\s?", row["output data names"])
        #Natures
        if not row.isna()["output data natures and resolution"]:
            list_outputs_nature = re.split("\s?;\s?", row["output data natures and resolution"])
            list_outputs_resolution = []
            for i in range(len(list_outputs_nature)):
                if ":" in list_outputs_nature[i]:
                    list_outputs_nature[i], resolution = re.split("\s?:\s?", list_outputs_nature[i])
                    list_outputs_resolution.append(resolution)
                else:
                    list_inputs_resolution.append(None)
        else:
            list_outputs_nature = [None] * len(list_outputs)

        #raster/vector
        if not row.isna()["output data raster/points/lines/polygon"]:
            list_outputs_raster_vector = re.split("\s?;\s?", 
                                       row["output data raster/points/lines/polygon"]
                                       )
        else:
            list_outputs_raster_vector = [""] * len(list_outputs)
        
        
        for i, output_name in enumerate(list_outputs):
            nature_i = list_outputs_nature[i]
            if nature_i not in defined_natures:
                nature_i = "data"
            output_instance = onto[nature_i](
                output_name.replace(" ","_") + "_" + nature + "_" + doi
                )
            output_instance.is_a.append(onto["output_data"])
            output_instance.label = output_name
            if list_outputs_resolution[i] is not None:
                output_instance.minimum_mapping_unit.append(list_outputs_resolution[i])
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
        operators_types = re.split("\s?;\s?", row["operator type"])
        operators_infos = re.split("\s?;\s?", row["operator description"])
        for i, operator_type in enumerate(operators_types):
            operator = onto[operator_type](
                operator_type + str(i) + "_" + doi
                )
            operator.label = operators_infos[i]
            process.hasOperator.append(
                operator
                )

    #Study case
    if not row.isna()["Study Area"]:
        study_cases = re.split("\s?;\s?", row["Study Area name"])
        countries = re.split("\s?;\s?", row["belongs to country"])
        for i, study_area in enumerate(study_cases):
            study_case = onto[row["geographic extent type"]](
                "study_case_"+study_area.replace(" ","_")
                )
            study_case.label = study_area
            if "[" in countries[i]:
                list_countries_i = re.split("\s?,\s?", countries[i].replace("[","").replace("]","") )
                study_case.belongsToCountry.extend(list_countries_i)
            else :    
                study_case.belongsToCountry.append( countries[i] )
        
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
        "algorithm_quality_assessment",
        "f1_score",
        "intersection_over_union",
        "algorithm_quality_assessment"
        "recall",
        "precision"
        ]
    for i, metric in enumerate(list_global_quality_metrics):
        if not row.isna()[metric]:
            metric_values = re.split("\s?;\s?", row[metric])
            for j, metric_value in enumerate(metric_values):
                algo_qual_assessment = onto[metrics_type[i]](
                    metric.replace(" ", "_")  + "_" + j + "_" + doi
                    )
                algo_qual_assessment.value.append(
                    eval( metric_value.replace("%","/100") )
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
        "algorithm_quality_assessment",
        "recall",
        "precision"
        ]

    for i, metric in enumerate(list_per_class_quality_metrics):
        if not row.isna()[metric]:
            metric_values = re.split("\s?;\s?", row[metric])
            for j, metric_value in enumerate(metric_values):

                algo_qual_assessment = onto[metrics_type[i]](
                    metric.replace(" ", "_") + "_" + j + "_" + doi
                    )
                lulc_class_name, value = re.split("\s?:\s?", metric_value)
                lulc_class = eval("onto.lulc_class_"+lulc_class_name.replace(" ", "_"))
                algo_qual_assessment.assessedOnClass.append(lulc_class)
                algo_qual_assessment.value.append(
                    eval( value.replace("%","/100") )
                    )
                process.hasAccuracyAlgorithm.append(algo_qual_assessment)
    
    if not row.isna()["user defined algorithm quality assessment metrics"]:
        other_metrics = re.split("\s?;\s?", row["user defined algorithm quality assessment metrics"])
        for j, metric in enumerate(other_metrics):
            metric_name_and_class, metric_value = re.split("\s?:\s?", metric)
            if "(" in metric_name_and_class:
                metric_name, lulc_class_name = re.split("\s?(\s?", metric_name_and_class.replace(")", ""))
                lulc_class = eval("onto.lulc_class_"+lulc_class_name.replace(" ", "_"))
            else:
                metric_name = metric_name_and_class
            algo_qual_assessment = onto["algorithm_quality_assessment"](
                metric_name.replace(" ", "_") + "_" + j + "_" + doi
                )
            algo_qual_assessment.label = metric_name
            if "(" in metric_name_and_class:
                algo_qual_assessment.assessedOnClass.append(lulc_class)
            algo_qual_assessment.value.append(
                    eval( value.replace("%","/100") )
                    )
            process.hasAccuracyAlgorithm.append(algo_qual_assessment)

    #criterions
    if row["codeAvailability "].lower not in(FALSE_VALUES):
        process.codeAvailability.append(row["codeAvailability "])
    if row["dataAvailability"].lower not in(FALSE_VALUES):
        process.dataAvailability.append(row["dataAvailability"])
    if row["challenge"] is not None:
        process.chalenge.extend(re.split("\s?;\s?", row["challenge"]))
    if row["strength"] is not None:
        process.strength.extend(re.split("\s?;\s?", row["strength"]))
    if row["weakness"] is not None:
        process.weakness.extend(re.split("\s?;\s?", row["weakness"]))
    
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
  
    owl_file_path = os.path.join(
        "lulc.owl"
        )
    
    onto = or2.get_ontology(owl_file_path).load()
    
    excel_ontology_folder_path = os.path.join(
        "LULC_Ontology_example.xlsm"
        )
    
    list_excel_files_path = get_excel_files(excel_ontology_folder_path)
    
    for excel_ontology_file_path in list_excel_files_path:
        excel_ontology_file = pd.read_excel(excel_ontology_file_path, dtype=str, sheet_name="ontology_instanciation", header=1)
        excel_ontology_file.drop(index=excel_ontology_file.index[0], axis=0, inplace=True)#The purpose of this row is to help the user on how to fill each column
        
        for i in range(len(excel_ontology_file)):
            row = excel_ontology_file.iloc[i]
            article = create_article(onto, row)
    
    onto.save(
        os.path.join(
            "lulc_instanciated.owl"
            )
        )