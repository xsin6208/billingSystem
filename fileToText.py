from __future__ import annotations
from google.api_core.client_options import ClientOptions
from google.cloud import documentai_v1
from pathlib import Path


from typing import Any, Dict, List

from typing import Optional

from google.api_core.client_options import ClientOptions
from google.cloud import documentai  # type: ignore


def process_document_sample(
    project_id: str,
    location: str,
    processor_id: str,
    file_path: str,
    mime_type: str,
    field_mask: Optional[str] = None,
    processor_version_id: Optional[str] = None,
) -> None:
    # You must set the `api_endpoint` if you use a location other than "us".
    opts = ClientOptions(api_endpoint=f"{location}-documentai.googleapis.com")

    client = documentai.DocumentProcessorServiceClient(client_options=opts)

    if processor_version_id:
        # The full resource name of the processor version, e.g.:
        # `projects/{project_id}/locations/{location}/processors/{processor_id}/processorVersions/{processor_version_id}`
        name = client.processor_version_path(
            project_id, location, processor_id, processor_version_id
        )
    else:
        # The full resource name of the processor, e.g.:
        # `projects/{project_id}/locations/{location}/processors/{processor_id}`
        name = client.processor_path(project_id, location, processor_id)

    # Read the file into memory
    with open(file_path, "rb") as image:
        image_content = image.read()

    # Load binary data
    raw_document = documentai.RawDocument(content=image_content, mime_type=mime_type)

    # For more information: https://cloud.google.com/document-ai/docs/reference/rest/v1/ProcessOptions
    # Optional: Additional configurations for processing.
    process_options = documentai.ProcessOptions(
        # Process only specific pages
        individual_page_selector=documentai.ProcessOptions.IndividualPageSelector(
            pages=[1]
        )
    )

    # Configure the process request
    request = documentai.ProcessRequest(
        name=name,
        raw_document=raw_document,
        field_mask=field_mask,
        #process_options=process_options,
    )

    result = client.process_document(request=request)

    # For a full list of `Document` object attributes, reference this page:
    # https://cloud.google.com/document-ai/docs/reference/rest/v1/Document
    document = result.document
    return document
    # for ent in document.entities:
    #     print(ent.type_)
    #     for prop in ent.properties:
    #         print(f"  {prop.type_}: {prop.mention_text}")
    
    # # Read the text recognition output from the processor
    # print("The document contains the following text:")
    # print(document.entities)



def entity_value(entity: documentai.Document.Entity) -> str:
    """Best-effort value string for an entity."""
    # mention_text is usually the extracted text for the entity
    if getattr(entity, "mention_text", None):
        return entity.mention_text.strip()
    # Some processors populate normalized_value (dates, numbers, etc.)
    nv = getattr(entity, "normalized_value", None)
    if nv and str(nv).strip():
        return str(nv).strip()
    return ""

def entity_to_obj(entity: documentai.Document.Entity) -> Any:
    """
    Convert an entity into:
      - a plain string value if it has no properties
      - a dict or list of dicts if it has nested properties / repeats
    """
    # If no nested properties, return the extracted value
    if not getattr(entity, "properties", None):
        return entity_value(entity)

    # If it has properties, build a structured object
    obj: Dict[str, Any] = {}
    for prop in entity.properties:
        key = prop.type_
        val = entity_to_obj(prop)

        # Handle repeated keys (e.g., multiple TestNumber entries)
        if key in obj:
            if not isinstance(obj[key], list):
                obj[key] = [obj[key]]
            obj[key].append(val)
        else:
            obj[key] = val

    # Sometimes the parent entity itself also has a useful value
    # (optional) include it under a special key if present
    parent_val = entity_value(entity)
    if parent_val and "__value__" not in obj:
        obj["__value__"] = parent_val

    return obj

def extract_patients(document: documentai.Document, patient_entity_name: str = "Patient") -> List[Dict[str, Any]]:
    """
    Return a list of patients as dictionaries, based on repeating top-level entity type.
    """
    patients: List[Dict[str, Any]] = []
    for ent in getattr(document, "entities", []) or []:
        if ent.type_ == patient_entity_name:
            parsed = entity_to_obj(ent)
            # Ensure each patient becomes a dict
            if isinstance(parsed, dict):
                patients.append(parsed)
            else:
                # If schema is odd and patient has no properties, store as a dict anyway
                patients.append({"__value__": parsed})
    return patients

# def extract_data(pdf_path):
#     # TODO(developer): Uncomment these variables before running the sample.
#     project_id = "project-9bb7cf59-486c-4126-a39"
#     location = "eu"
#     processor_id = "509bb47d15290df7"
#     file_path = pdf_path
#     mime_type = "application/pdf" # Refer to https://cloud.google.com/document-ai/docs/file-types for supported file types
#     field_mask = "text,entities,pages.pageNumber"  # Optional. The fields to return in the Document object.
#     # processor_version_id = "YOUR_PROCESSOR_VERSION_ID" # Optional. Processor version to use

#     document = process_document_sample(project_id, location, processor_id, file_path, mime_type, field_mask)
#     patients = extract_patients(document, patient_entity_name="Patient")
#     return patients

# [END documentai_process_document]

def extract_data(pdf_path):
    # TODO(developer): Uncomment these variables before running the sample.
    project_id = "project-9bb7cf59-486c-4126-a39"
    location = "eu"
    processor_id = "50e239fc1af622b3"
    file_path = pdf_path
    mime_type = "application/pdf" # Refer to https://cloud.google.com/document-ai/docs/file-types for supported file types
    field_mask = "text,entities,pages.pageNumber"  # Optional. The fields to return in the Document object.
    # processor_version_id = "YOUR_PROCESSOR_VERSION_ID" # Optional. Processor version to use

    document = process_document_sample(project_id, location, processor_id, file_path, mime_type, field_mask)
    patients = extract_patients(document, patient_entity_name="Patient")
    return patients


# pdf_path = Path(r"C:\Users\xsing\OneDrive\Desktop\Billing\Scan_20251214 (9).pdf")
# print(extract_data(pdf_path))