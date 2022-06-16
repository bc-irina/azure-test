# webhookApp

### How to deploy webhook-app
Use Microsoft Azure Portal to create a new ResourceGroup for desired region
https://docs.microsoft.com/en-us/azure/azure-resource-manager/management/manage-resource-groups-portal

### or use Azure cli
```
az group create --name <resourceGroupName> --location <regionName>
```

### Deploy ARM template by using the Azure portal
-Deploy a custom template-> Build your own template in the editor
then load azure-deployment-template.json file
select serource group and populate userId to grant this user access to secret value
start deployment

### or use Azure cli
set valid userId in the parameters json file

```
az deployment group create --resource-group <resourceGroupName> --template-file azure-deployment-template.json --parameters azure-deployment-parameters.json
```


## Console command to generate a report
```
curl -v -u user:password https://<host-name>/api/report/2020-01-01/2022-06-06 --output report.xlsx
```

## UI report gereration 
update report.html file with a valid host-name and auhtorisation details (username: password)
open report.html in the browser 
```
        var HOST_NAME="https://<host-name>"
        var USER="user"
        var PASSWORD="password"

```
