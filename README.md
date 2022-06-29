# WebhookApp

### How to deploy webhook-app

Use Microsoft Azure Portal to create a new ResourceGroup for desired region
https://docs.microsoft.com/en-us/azure/azure-resource-manager/management/manage-resource-groups-portal

### or use Azure cli
```
az group create --name <resourceGroupName> --location <regionName>
```

### Deploy ARM template by using Azure portal
1. Open "Deploy a custom template"
2. Click on "Build your own template in the editor"
3. Load azure-deployment-template.json file
4. Populate parameters
* select resource group 
* populate userId to grant this user access to secret value (UserId could be extracted from Azure Users->Object Id)

 5. Start deployment

### or use Azure cli
Update parameters json file to set userId, update app-name, db name if requied

run following command with azure cli

```
az deployment group create --resource-group <resourceGroupName> --template-file azure-deployment-template.json --parameters azure-deployment-parameters.json
```

## After Deployment

After deployment is done extract functionHostName from output parameters in the Azure Deployments-> Outputs 
or from cli command outputs

from Azure Key Vaults extract a value of created secret WebHookAuth which should be in a following format (user:password)

## Webhook url
Create webhook url as below and set it as a value of callback url in the management portal

```
https://user:password@<function-host-name>/api/webhook
``` 


## Report Generation

To generate a report:
1. open web url, 
2. enter user and password 
3. choose From and To dates 
4. click "Generate Report" button

```
https://<function-host-name>/api/home
```

Or use command line to generate a report. URL contains the date range and should be constructed as "/api/report/{from date yyyy-mm-dd}/{to date yyyy-mm-dd}",

```
curl -v -u user:password https://<function-host-name>/api/report/2020-01-20/2022-06-26 --output reportName.xlsx
```

