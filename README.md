# JSON2Dataverse

 - Setup
 - Configuring .JSON file
 - Usage

## Setup
 1. Create `./Excel2Dataverse/app.config` file
 2. Add [`connectionString`](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/xrm-tooling/use-connection-strings-xrm-tooling-connect#connection-string-parameters) called 'dynamics' to the file. Or copy the example one below to get started (Replacing 'YOUR_ENVIRONMENT_URL' with your Dataverse enironment URL, and 'YOUR_AZURE_APPID' with the Microsoft Entra SSO App ID)
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
JSON schema is as follows...
```json
[
  {
    "Name": "ENTITY_NAME",
    "Members": {
      "ATTRIBUTE_NAME": "ATTRIBUTE_DATATYPE"
    },
    "ManyToOne": {
      "RELATIONSHIPNAME": "ENTITY_NAME_RELATING_TO"
    },
    "ManyToMany": {
      "RELATIONSHIPNAME": "ENTITY_NAME_RELATING_TO"
    }
  }
]
```

## Usage
Follow the guidance on [Setup](#setup) and [Configuring .JSON file](#configuring-json-file). Once you have the script setup and the JSON schema created, you can run the script to build the Dataverse tables. This will first create the entities and attributes, then create the relationships (Many-to-one first, then Many-to-many).
