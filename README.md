# JSON2Dataverse
Quickly create entities, attributes, keys and relationships in Dataverse from a JSON file.
- [JSON2Dataverse](#json2dataverse)
  - [Setup](#setup)
    - [Example app.config](#example-appconfig)
  - [Configuring .JSON file](#configuring-json-file)
    - [Structure](#structure)
    - [Data types](#data-types)
  - [Usage](#usage)

## Setup
Before you can run the script, you will need to setup the connection between it and your Dataverse instance:

 1. Create `./Excel2Dataverse/app.config` file
 2. Add [`connectionString`](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/xrm-tooling/use-connection-strings-xrm-tooling-connect#connection-string-parameters) with the name `dynamics` to the file. Or copy the example one below to get started 

### Example app.config

> [!NOTE]
> Replace: `YOUR_ENVIRONMENT_URL` with your Dataverse enironment URL and, `YOUR_AZURE_APPID` with the Microsoft Entra SSO App ID

```xml
<?xml version='1.0' encoding='utf-8'?>
<configuration>
  <connectionStrings>
    <add name="dynamics"
     providerName="Microsoft.Xrm.Sdk.Client"
     connectionString="AuthType = OAuth;
     Url = YOUR_ENVIRONMENT_URL;
     AppId = YOUR_AZURE_APPID;
     RedirectUri = http://localhost:64636;
     LoginPrompt = Auto;
     RequireNewInstance = True" />
  </connectionStrings>
</configuration>
```

## Configuring .JSON file
The JSON file is used as a instructionset for the script.

### Structure
```json
[
  {
    "Name": "ENTITY_NAME",
    // Optional
    // Can be ATTRIBUTE_DATATYPE: "string", "int", "datetime"
    "PrimaryKey": "ATTRIBUTE_NAME",   
    "Members": {
      // First member is primary attribute
      // Can be ATTRIBUTE_DATATYPE: "string"
      "ATTRIBUTE_NAME": "ATTRIBUTE_DATATYPE",
      "ATTRIBUTE_NAME": "ATTRIBUTE_DATATYPE"
    },
    // Optional
    "ManyToOne": {
      "RELATIONSHIPNAME": "ENTITY_NAME_RELATING_TO"
    },
    // Optional
    "ManyToMany": {
      "RELATIONSHIPNAME": "ENTITY_NAME_RELATING_TO"
    }
  }
]
```

### Data types
| ATTRIBUTE_DATATYPE | Dataverse type     |
|--------------------|--------------------|
| enum               | Single line text   |
| string             | Single line text   |
| boolean            | Yes/No             |
| datetime           | Date and time      |
| integer            | Whole number       |
| long               | Whole number (Big) |
| decimal            | Decimal            |
| binary             | File               |

## Usage
Once you have the script setup and the JSON file created, you can run the script to build the Dataverse entities. This will first create the entities, attributes and keys; then create the relationships (Many-to-one and then Many-to-many).
