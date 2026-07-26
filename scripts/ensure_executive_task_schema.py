import os
import sys
import time

import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from auth import get_client, get_plugin_headers, get_token, load_env


TABLE_SCHEMA = "new_ExecutiveTask"
TABLE_LOGICAL = "new_executivetask"
SOLUTION_NAME = "WescoDashboard"

STANDARD_COLUMNS = {
    "new_TaskKey": "string",
    "new_GroupKey": "string",
    "new_AssignedTo": "string",
    "new_AssigneeEmail": "string",
    "new_ProjectNumber": "string",
    "new_Location": "string",
    "new_DateAssigned": "string",
    "new_DueDate": "string",
    "new_CommunicationMethod": "string",
    "new_Priority": "string",
    "new_Status": "string",
    "new_CompletedAt": "string",
    "new_Photo": "file",
    "new_Document": "file",
    "new_PhotoName": "string",
    "new_DocumentName": "string",
}


def metadata_headers(token, include_solution=False):
    headers = get_plugin_headers("dv-metadata", token)
    headers.update(
        {
            "Accept": "application/json",
            "Content-Type": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
        }
    )
    if include_solution:
        headers["MSCRM.SolutionName"] = SOLUTION_NAME
    return headers


def label(text):
    return {
        "@odata.type": "Microsoft.Dynamics.CRM.Label",
        "LocalizedLabels": [
            {
                "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
                "Label": text,
                "LanguageCode": 1033,
            }
        ],
    }


def get_table_metadata(base_url, token):
    url = (
        f"{base_url}/api/data/v9.2/"
        f"EntityDefinitions(LogicalName='{TABLE_LOGICAL}')"
        "?$select=LogicalName,SchemaName,EntitySetName,PrimaryIdAttribute"
    )
    response = requests.get(url, headers=metadata_headers(token), timeout=60)
    if response.status_code == 404:
        return None
    response.raise_for_status()
    return response.json()


def get_existing_columns(base_url, token):
    url = (
        f"{base_url}/api/data/v9.2/"
        f"EntityDefinitions(LogicalName='{TABLE_LOGICAL}')/Attributes"
        "?$select=LogicalName"
    )
    response = requests.get(url, headers=metadata_headers(token), timeout=60)
    response.raise_for_status()
    return {item["LogicalName"].lower() for item in response.json().get("value", [])}


def ensure_completed_at_string(client, base_url, token):
    incompatible = []
    for logical_name in ("new_dateassigned", "new_duedate", "new_completedat"):
        url = (
            f"{base_url}/api/data/v9.2/"
            f"EntityDefinitions(LogicalName='{TABLE_LOGICAL}')/Attributes"
            f"(LogicalName='{logical_name}')"
            "?$select=MetadataId,AttributeType"
        )
        response = requests.get(url, headers=metadata_headers(token), timeout=60)
        if response.status_code == 404:
            incompatible.append(logical_name)
            continue
        response.raise_for_status()
        if response.json().get("AttributeType") != "String":
            incompatible.append(logical_name)

    if not incompatible:
        print("Reusing ISO date text columns.")
        return

    delete_url = (
        f"{base_url}/api/data/v9.2/"
        f"EntityDefinitions(LogicalName='{TABLE_LOGICAL}')"
    )
    delete_response = requests.delete(
        delete_url,
        headers=metadata_headers(token, include_solution=True),
        timeout=120,
    )
    if not delete_response.ok:
        raise RuntimeError(
            f"Table reset failed ({delete_response.status_code}): {delete_response.text}"
        )
    print(
        "Removed empty table containing incompatible date metadata: "
        + ", ".join(incompatible)
    )
    time.sleep(30)
    info = client.tables.create(
        TABLE_SCHEMA,
        STANDARD_COLUMNS,
        solution=SOLUTION_NAME,
        primary_column="new_TaskTitle",
        display_name="Executive Task",
    )
    print(f"Recreated table: {info['table_schema_name']}")
    time.sleep(20)


def ensure_instructions_memo(base_url, token, existing_columns):
    if "new_instructions" in existing_columns:
        print("Reusing column: new_instructions")
        return
    url = (
        f"{base_url}/api/data/v9.2/"
        f"EntityDefinitions(LogicalName='{TABLE_LOGICAL}')/Attributes"
    )
    payload = {
        "@odata.type": "Microsoft.Dynamics.CRM.MemoAttributeMetadata",
        "AttributeType": "Memo",
        "AttributeTypeName": {"Value": "MemoType"},
        "SchemaName": "new_Instructions",
        "DisplayName": label("Instructions"),
        "Description": label("Task instructions and completion expectations."),
        "RequiredLevel": {"Value": "None"},
        "MaxLength": 10000,
        "Format": "TextArea",
        "ImeMode": "Disabled",
        "IsLocalizable": False,
    }
    response = requests.post(
        url,
        headers=metadata_headers(token, include_solution=True),
        json=payload,
        timeout=120,
    )
    if not response.ok:
        raise RuntimeError(
            f"Instructions column creation failed ({response.status_code}): {response.text}"
        )
    print("Created column: new_instructions")


def ensure_alternate_key(client):
    key_name = "new_ExecutiveTaskKey"
    existing = client.tables.get_alternate_keys(TABLE_SCHEMA)
    if any(key.schema_name.lower() == key_name.lower() for key in existing):
        print(f"Reusing alternate key: {key_name}")
        return
    key = client.tables.create_alternate_key(
        TABLE_SCHEMA,
        key_name,
        ["new_taskkey"],
        display_name="Executive Task Key",
    )
    print(f"Created alternate key: {key.schema_name} ({key.status})")


def main():
    load_env()
    base_url = os.environ["DATAVERSE_URL"].rstrip("/")
    client = get_client("dv-metadata")
    token = get_token()

    table = client.tables.get(TABLE_LOGICAL)
    if table:
        print(f"Reusing table: {TABLE_LOGICAL}")
    else:
        info = client.tables.create(
            TABLE_SCHEMA,
            STANDARD_COLUMNS,
            solution=SOLUTION_NAME,
            primary_column="new_TaskTitle",
            display_name="Executive Task",
        )
        print(f"Created table: {info['table_schema_name']}")
        time.sleep(20)

    existing_columns = get_existing_columns(base_url, token)
    missing = {
        schema: column_type
        for schema, column_type in STANDARD_COLUMNS.items()
        if schema.lower() not in existing_columns
    }
    if missing:
        created = client.tables.add_columns(TABLE_SCHEMA, missing)
        print("Created columns: " + ", ".join(created))
        time.sleep(15)
        existing_columns = get_existing_columns(base_url, token)

    ensure_completed_at_string(client, base_url, token)
    existing_columns = get_existing_columns(base_url, token)
    ensure_instructions_memo(base_url, token, existing_columns)
    time.sleep(10)
    ensure_alternate_key(client)

    metadata = get_table_metadata(base_url, token)
    print("Verified table:")
    print(f"  Logical name: {metadata['LogicalName']}")
    print(f"  Entity set: {metadata['EntitySetName']}")
    print(f"  Primary ID: {metadata['PrimaryIdAttribute']}")


if __name__ == "__main__":
    main()
