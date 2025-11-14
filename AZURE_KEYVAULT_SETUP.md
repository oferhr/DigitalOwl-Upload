# Azure Key Vault Setup Guide

This document explains how to configure Azure Key Vault for the DigitalOwl-Upload application.

## Overview

The application now uses **Azure Key Vault** to securely store the Digital Owl API key instead of storing it in a Word document. This significantly improves security by:

- Encrypting secrets at rest
- Providing centralized secret management
- Enabling secret rotation without code changes
- Supporting audit logging of secret access
- Preventing secrets from being committed to source control

## Prerequisites

- Azure subscription
- Azure CLI installed (for local development)
- Appropriate permissions to create Azure resources

## Setup Steps

### 1. Create an Azure Key Vault

Using Azure CLI:

```bash
# Login to Azure
az login

# Set your subscription (if you have multiple)
az account set --subscription "YOUR_SUBSCRIPTION_ID"

# Create a resource group (if you don't have one)
az group create --name "rg-digitalowl" --location "eastus"

# Create the Key Vault
az keyvault create \
  --name "kv-digitalowl-prod" \
  --resource-group "rg-digitalowl" \
  --location "eastus"
```

Or using Azure Portal:
1. Navigate to Azure Portal (https://portal.azure.com)
2. Click "Create a resource"
3. Search for "Key Vault"
4. Fill in the required details:
   - Name: `kv-digitalowl-prod`
   - Resource Group: Create new or use existing
   - Region: Choose your preferred region
5. Click "Review + Create"

### 2. Add the API Key Secret to Key Vault

Using Azure CLI:

```bash
# Store your Digital Owl API key in Key Vault
az keyvault secret set \
  --vault-name "kv-digitalowl-prod" \
  --name "DigitalOwlApiKey" \
  --value "YOUR_ACTUAL_API_KEY_HERE"
```

Or using Azure Portal:
1. Navigate to your Key Vault
2. Click "Secrets" in the left menu
3. Click "+ Generate/Import"
4. Fill in:
   - Name: `DigitalOwlApiKey`
   - Value: Your Digital Owl API key
5. Click "Create"

### 3. Configure Authentication

The application uses `DefaultAzureCredential` which supports multiple authentication methods:

#### For Production (Recommended: Managed Identity)

If running on an Azure VM, App Service, or Container Instance:

```bash
# Enable system-assigned managed identity on your Azure VM/App Service
az vm identity assign --name "your-vm-name" --resource-group "rg-digitalowl"

# Grant the managed identity access to Key Vault
az keyvault set-policy \
  --name "kv-digitalowl-prod" \
  --object-id "MANAGED_IDENTITY_OBJECT_ID" \
  --secret-permissions get list
```

#### For Local Development

**Option A: Azure CLI (Recommended)**

```bash
# Login with Azure CLI
az login

# The application will automatically use your Azure CLI credentials
```

**Option B: Service Principal with Environment Variables**

```bash
# Create a service principal
az ad sp create-for-rbac --name "sp-digitalowl-upload"

# This will output:
# {
#   "appId": "YOUR_CLIENT_ID",
#   "password": "YOUR_CLIENT_SECRET",
#   "tenant": "YOUR_TENANT_ID"
# }

# Grant the service principal access to Key Vault
az keyvault set-policy \
  --name "kv-digitalowl-prod" \
  --spn "YOUR_CLIENT_ID" \
  --secret-permissions get list

# Set environment variables (Windows)
setx AZURE_CLIENT_ID "YOUR_CLIENT_ID"
setx AZURE_CLIENT_SECRET "YOUR_CLIENT_SECRET"
setx AZURE_TENANT_ID "YOUR_TENANT_ID"

# Set environment variables (PowerShell)
$env:AZURE_CLIENT_ID = "YOUR_CLIENT_ID"
$env:AZURE_CLIENT_SECRET = "YOUR_CLIENT_SECRET"
$env:AZURE_TENANT_ID = "YOUR_TENANT_ID"
```

### 4. Update App.config

Update the `App.config` file with your Key Vault details:

```xml
<appSettings>
  <!-- Azure Key Vault Configuration -->
  <add key="KeyVaultUrl" value="https://kv-digitalowl-prod.vault.azure.net/"/>
  <add key="KeyVaultSecretName" value="DigitalOwlApiKey"/>

  <!-- Other settings... -->
  <add key="baseUrl" value="https://api.il.digitalowl.app"/>
</appSettings>
```

### 5. Install NuGet Packages

The following packages are required (already added to packages.config):

```bash
dotnet add package Azure.Identity
dotnet add package Azure.Security.KeyVault.Secrets
```

Or using Package Manager Console in Visual Studio:

```powershell
Install-Package Azure.Identity
Install-Package Azure.Security.KeyVault.Secrets
```

## Verification

To verify the setup is working:

1. Run the application
2. Check the logs for the message:
   ```
   Successfully retrieved secret 'DigitalOwlApiKey' from Azure Key Vault
   ```

If you see authentication errors, check:
- You are logged in with `az login` (for local development)
- The Managed Identity has been granted access (for production)
- Environment variables are set correctly (if using service principal)
- The Key Vault URL and secret name in App.config are correct

## Security Best Practices

### 1. Access Control
- Use Managed Identity for production workloads
- Grant least-privilege access (only `get` and `list` permissions for secrets)
- Regularly review access policies

### 2. Secret Rotation
To rotate the API key:

```bash
# Update the secret in Key Vault
az keyvault secret set \
  --vault-name "kv-digitalowl-prod" \
  --name "DigitalOwlApiKey" \
  --value "NEW_API_KEY_HERE"

# Restart the application to pick up the new key
```

### 3. Audit Logging
Enable diagnostic logging for Key Vault:

```bash
az monitor diagnostic-settings create \
  --name "kv-audit-logs" \
  --resource "/subscriptions/YOUR_SUBSCRIPTION_ID/resourceGroups/rg-digitalowl/providers/Microsoft.KeyVault/vaults/kv-digitalowl-prod" \
  --logs '[{"category": "AuditEvent", "enabled": true}]' \
  --workspace "YOUR_LOG_ANALYTICS_WORKSPACE_ID"
```

### 4. Networking (Optional)
For enhanced security, restrict Key Vault access to specific networks:

```bash
az keyvault update \
  --name "kv-digitalowl-prod" \
  --default-action Deny

az keyvault network-rule add \
  --name "kv-digitalowl-prod" \
  --ip-address "YOUR_PUBLIC_IP/32"
```

## Troubleshooting

### Error: "AuthenticationFailedException"

**Cause**: No valid credentials found

**Solution**:
- For local dev: Run `az login`
- For production: Ensure Managed Identity is enabled and has Key Vault access
- Check environment variables if using service principal

### Error: "Forbidden" (403)

**Cause**: Insufficient permissions

**Solution**:
```bash
# Grant access to the identity
az keyvault set-policy \
  --name "kv-digitalowl-prod" \
  --object-id "YOUR_IDENTITY_OBJECT_ID" \
  --secret-permissions get list
```

### Error: "SecretNotFound"

**Cause**: Secret name doesn't match

**Solution**: Verify the secret name in Key Vault matches `KeyVaultSecretName` in App.config

## Migration from Word Document

The old method of storing the API key in `key1.docx` is now **deprecated** for security reasons.

To migrate:
1. Copy your API key from `key1.docx`
2. Store it in Azure Key Vault (see step 2 above)
3. Update `App.config` with Key Vault details
4. Delete or secure the old `key1.docx` file
5. Remove the old configuration: `<add key="keyFile" value="..."/>`

## References

- [Azure Key Vault Documentation](https://docs.microsoft.com/en-us/azure/key-vault/)
- [DefaultAzureCredential Documentation](https://docs.microsoft.com/en-us/dotnet/api/azure.identity.defaultazurecredential)
- [Managed Identity Documentation](https://docs.microsoft.com/en-us/azure/active-directory/managed-identities-azure-resources/)

## Support

For issues related to:
- **Azure Key Vault**: Contact Azure Support
- **Application issues**: Check application logs in the working directory
- **Digital Owl API**: Contact Digital Owl support
