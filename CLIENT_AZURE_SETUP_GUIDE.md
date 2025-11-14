# Azure Key Vault Setup Instructions for DigitalOwl-Upload

**For:** Client Azure Account Administrator
**Purpose:** Secure storage of DigitalOwl API Key
**Time Required:** 15-20 minutes
**Prerequisites:** Azure account with Owner or Contributor permissions

---

## Overview

This guide will walk you through creating an Azure Key Vault and storing your DigitalOwl API key securely. This replaces the previous insecure method of storing the API key in a Word document.

**Why this is important:**
- ✅ Your API key will be encrypted and secure
- ✅ Access to the key can be controlled and audited
- ✅ No risk of accidentally sharing the key in files
- ✅ Easy to rotate the key without changing code

---

## Step 1: Sign in to Azure Portal

1. Open your web browser and go to: **https://portal.azure.com**
2. Sign in with your Azure account credentials
3. Once signed in, you should see the Azure Portal dashboard

---

## Step 2: Create a Resource Group (If You Don't Have One)

A Resource Group is a container that holds related Azure resources.

### Option A: Using Azure Portal (Recommended for most users)

1. In the Azure Portal, click the **☰ menu** (hamburger icon) on the top left
2. Click **"Resource groups"**
3. Click **"+ Create"** at the top
4. Fill in the details:
   - **Subscription:** Select your Azure subscription
   - **Resource group name:** `rg-digitalowl-prod`
   - **Region:** Choose the region closest to your location (e.g., "East US", "West Europe")
5. Click **"Review + Create"**
6. Click **"Create"**

### Option B: Using Azure CLI (For technical users)

```bash
# Open PowerShell or Command Prompt
az login

# Create the resource group
az group create --name "rg-digitalowl-prod" --location "eastus"
```

**✅ Verification:** You should see your new resource group in the Resource Groups list.

---

## Step 3: Create the Key Vault

### Option A: Using Azure Portal (Recommended)

1. In the Azure Portal search bar at the top, type **"Key vaults"**
2. Click **"Key vaults"** from the results
3. Click **"+ Create"** at the top

4. **Basics Tab:**
   - **Subscription:** Select your Azure subscription
   - **Resource group:** Select `rg-digitalowl-prod` (or the one you created)
   - **Key vault name:** `kv-digitalowl-prod`
     - ⚠️ **IMPORTANT:** This name must be globally unique across all of Azure
     - If `kv-digitalowl-prod` is taken, try: `kv-digitalowl-prod-[yourcompany]` or `kv-digitalowl-[random-number]`
     - **Write down this name - you'll need it later!**
   - **Region:** Same region as your resource group
   - **Pricing tier:** **Standard** (this is sufficient for most needs)
   - **Days to retain deleted vaults:** 90 (default is fine)
   - **Purge protection:** Disabled (can enable for extra protection)

5. Click **"Next: Access configuration"**

6. **Access Configuration Tab:**
   - **Permission model:** Select **"Vault access policy"**
   - Leave other settings as default

7. Click **"Next: Networking"**

8. **Networking Tab:**
   - **Connectivity method:** Select **"Public endpoint (all networks)"**
     - Note: For enhanced security in production, you can restrict this later

9. Click **"Next: Tags"** (optional, can skip)

10. Click **"Review + Create"**

11. Review your settings and click **"Create"**

12. Wait for deployment to complete (usually 1-2 minutes)

13. Click **"Go to resource"**

### Option B: Using Azure CLI

```bash
# Replace with your chosen Key Vault name if needed
az keyvault create \
  --name "kv-digitalowl-prod" \
  --resource-group "rg-digitalowl-prod" \
  --location "eastus"
```

**✅ Verification:** You should see your Key Vault in the Azure Portal under Key vaults.

---

## Step 4: Add Your DigitalOwl API Key to the Key Vault

### Option A: Using Azure Portal (Recommended)

1. In your Key Vault page, look at the left menu
2. Under **"Objects"**, click **"Secrets"**
3. Click **"+ Generate/Import"** at the top

4. Fill in the secret details:
   - **Upload options:** **Manual**
   - **Name:** `DigitalOwlApiKey`
     - ⚠️ **IMPORTANT:** Use this exact name (it's case-sensitive)
   - **Value:** Paste your actual DigitalOwl API key here
     - This is the key that was previously stored in the Word document
     - Example format: `Bearer_abc123xyz...` or just the token itself
   - **Content type:** (leave empty or type "DigitalOwl API Token")
   - **Set activation date:** No (leave unchecked)
   - **Set expiration date:** No (leave unchecked)
   - **Enabled:** Yes (checked)

5. Click **"Create"**

6. You should see your secret `DigitalOwlApiKey` in the list

### Option B: Using Azure CLI

```bash
# Replace YOUR_API_KEY_HERE with your actual API key
az keyvault secret set \
  --vault-name "kv-digitalowl-prod" \
  --name "DigitalOwlApiKey" \
  --value "YOUR_API_KEY_HERE"
```

**✅ Verification:** Click on the secret name, then click the current version. You should see the value (masked by default for security).

---

## Step 5: Get Your Key Vault URL

You need this URL to configure the application.

### Using Azure Portal:

1. In your Key Vault page, look at the **"Overview"** section
2. Find **"Vault URI"** on the right side
3. Copy this URL - it will look like: `https://kv-digitalowl-prod.vault.azure.net/`

### Using Azure CLI:

```bash
az keyvault show --name "kv-digitalowl-prod" --query properties.vaultUri --output tsv
```

**📝 Write down your Vault URI - you'll need it for the application configuration!**

---

## Step 6: Grant Access to the Application

The application needs permission to read secrets from your Key Vault. The best method depends on where the application runs.

### Option A: If Running on Your Computer (Development/Testing)

You'll authenticate using your Azure account:

1. Install Azure CLI if not already installed:
   - Download from: https://aka.ms/installazurecliwindows
   - Run the installer

2. Open PowerShell or Command Prompt and run:
   ```bash
   az login
   ```

3. A browser window will open - sign in with your Azure account

4. Grant yourself access to the Key Vault:
   ```bash
   # Get your Azure user object ID
   az ad signed-in-user show --query id --output tsv

   # Grant yourself access (replace YOUR_OBJECT_ID with the ID from above)
   az keyvault set-policy \
     --name "kv-digitalowl-prod" \
     --object-id "YOUR_OBJECT_ID" \
     --secret-permissions get list
   ```

### Option B: If Running on an Azure Virtual Machine (Production)

Use Managed Identity (most secure):

1. In Azure Portal, go to your Virtual Machine

2. Click **"Identity"** in the left menu

3. Under **"System assigned"** tab:
   - Set **Status** to **"On"**
   - Click **"Save"**
   - Click **"Yes"** to confirm

4. Copy the **Object ID** that appears

5. Go back to your Key Vault

6. Click **"Access policies"** in the left menu

7. Click **"+ Create"**

8. **Permissions Tab:**
   - Under **Secret permissions**, check:
     - ✅ Get
     - ✅ List
   - Click **"Next"**

9. **Principal Tab:**
   - Paste the VM's Object ID in the search box
   - Select your VM from the results
   - Click **"Next"**

10. **Application Tab:** (skip, click **"Next"**)

11. **Review + create Tab:**
    - Click **"Create"**

### Option C: Using Service Principal (For CI/CD or Automation)

1. Create a Service Principal:
   ```bash
   az ad sp create-for-rbac --name "sp-digitalowl-app" --skip-assignment
   ```

2. Save the output (you'll need these values):
   ```json
   {
     "appId": "12345678-1234-1234-1234-123456789abc",
     "displayName": "sp-digitalowl-app",
     "password": "your-client-secret",
     "tenant": "87654321-4321-4321-4321-cba987654321"
   }
   ```

3. Grant the Service Principal access to Key Vault:
   ```bash
   az keyvault set-policy \
     --name "kv-digitalowl-prod" \
     --spn "12345678-1234-1234-1234-123456789abc" \
     --secret-permissions get list
   ```

4. Set these environment variables on the machine running the application:
   - `AZURE_CLIENT_ID` = appId from step 2
   - `AZURE_CLIENT_SECRET` = password from step 2
   - `AZURE_TENANT_ID` = tenant from step 2

---

## Step 7: Update Application Configuration

1. Open the application's `App.config` file

2. Find the Azure Key Vault section:
   ```xml
   <!-- Azure Key Vault Configuration -->
   <add key="KeyVaultUrl" value="https://your-keyvault-name.vault.azure.net/"/>
   <add key="KeyVaultSecretName" value="DigitalOwlApiKey"/>
   ```

3. Replace `https://your-keyvault-name.vault.azure.net/` with your actual Vault URI from Step 5

4. Keep `KeyVaultSecretName` as `DigitalOwlApiKey` (unless you used a different name)

5. Save the file

**Example:**
```xml
<!-- Azure Key Vault Configuration -->
<add key="KeyVaultUrl" value="https://kv-digitalowl-prod.vault.azure.net/"/>
<add key="KeyVaultSecretName" value="DigitalOwlApiKey"/>
```

---

## Step 8: Test the Setup

1. Run the DigitalOwl-Upload application

2. Check the log file for these messages:
   - ✅ `Attempting to retrieve secret 'DigitalOwlApiKey' from Azure Key Vault: https://kv-digitalowl-prod.vault.azure.net/`
   - ✅ `Successfully retrieved secret 'DigitalOwlApiKey' from Azure Key Vault`

3. If you see errors, proceed to the Troubleshooting section below

---

## Troubleshooting

### Error: "AuthenticationFailedException"

**Problem:** The application cannot authenticate to Azure.

**Solutions:**
- **If running locally:** Make sure you ran `az login` in PowerShell/Command Prompt
- **If running on Azure VM:** Make sure Managed Identity is enabled (Step 6, Option B)
- **If using Service Principal:** Check that environment variables are set correctly

### Error: "ForbiddenByPolicy" or "Access Denied" (403)

**Problem:** The identity doesn't have permission to access the Key Vault.

**Solution:**
1. Go to your Key Vault in Azure Portal
2. Click **"Access policies"** in the left menu
3. Verify there's a policy with **Get** and **List** permissions for:
   - Your user account (if running locally), OR
   - Your VM's Managed Identity (if running on Azure VM), OR
   - Your Service Principal (if using that method)
4. If missing, add the access policy using Step 6 above

### Error: "SecretNotFound"

**Problem:** The secret name doesn't match.

**Solution:**
1. Go to your Key Vault
2. Click **"Secrets"**
3. Verify the secret name is exactly `DigitalOwlApiKey` (case-sensitive)
4. If different, either:
   - Rename the secret in Key Vault, OR
   - Update `KeyVaultSecretName` in App.config to match

### Error: "VaultNotFound"

**Problem:** The Key Vault URL is incorrect.

**Solution:**
1. Go to your Key Vault in Azure Portal
2. Copy the **Vault URI** from the Overview page
3. Update `KeyVaultUrl` in App.config with the exact URL

### The application can't find my Azure credentials

**For local development:**
```bash
# Re-run Azure login
az login

# Verify you're logged in
az account show
```

**For Azure VM:**
- Make sure Managed Identity is enabled
- Restart the VM after enabling Managed Identity
- Wait 5 minutes for the identity to propagate

---

## Security Best Practices

### ✅ Do's:

1. **Use Managed Identity** for production (most secure)
2. **Enable soft-delete** on the Key Vault (prevents accidental deletion)
3. **Enable purge protection** for production Key Vaults
4. **Regularly review access policies** (remove unused access)
5. **Enable diagnostic logging** to track who accesses secrets
6. **Use separate Key Vaults** for dev/test/production environments

### ❌ Don'ts:

1. **Don't share the API key** via email, chat, or documents anymore
2. **Don't grant more permissions than needed** (only Get and List for secrets)
3. **Don't use the same Key Vault** for multiple unrelated applications
4. **Don't disable Azure security recommendations** without good reason

---

## How to Rotate the API Key

When you need to change the API key (recommended every 90 days):

1. Get your new API key from DigitalOwl

2. Go to your Key Vault in Azure Portal

3. Click **"Secrets"** → **"DigitalOwlApiKey"**

4. Click **"+ New Version"** at the top

5. Paste the new API key value

6. Click **"Create"**

7. Restart the DigitalOwl-Upload application

**That's it!** No code changes needed - the application will automatically use the new key.

---

## Cost Information

Azure Key Vault pricing (as of 2024):

- **Key Vault:** ~$0.03 per 10,000 transactions
- **Secret storage:** Free for first 10 secrets, then minimal cost
- **Expected monthly cost:** Less than $1 for typical usage

This is a small price for the security benefits!

---

## Summary Checklist

Use this checklist to verify everything is set up correctly:

- [ ] Created Resource Group: `rg-digitalowl-prod`
- [ ] Created Key Vault: `kv-digitalowl-prod` (or your custom name)
- [ ] Added secret: `DigitalOwlApiKey` with your API key value
- [ ] Copied Vault URI: `https://kv-digitalowl-prod.vault.azure.net/`
- [ ] Granted access (Managed Identity, Azure CLI login, or Service Principal)
- [ ] Updated `App.config` with correct `KeyVaultUrl`
- [ ] Tested application and saw success messages in logs
- [ ] Deleted or secured the old `key1.docx` file
- [ ] Application is running successfully

---

## Quick Reference

**Your Key Vault Details:** (Fill this out for your records)

| Setting | Value |
|---------|-------|
| Key Vault Name | `kv-digitalowl-prod` |
| Vault URI | `https://kv-digitalowl-prod.vault.azure.net/` |
| Secret Name | `DigitalOwlApiKey` |
| Resource Group | `rg-digitalowl-prod` |
| Region | (e.g., East US) |

---

## Need Help?

If you encounter issues not covered in the Troubleshooting section:

1. **Check application logs** in the working directory for detailed error messages
2. **Review Azure Key Vault documentation:** https://docs.microsoft.com/en-us/azure/key-vault/
3. **Contact Azure Support** through the Azure Portal if it's an Azure-specific issue
4. **Contact your development team** with the error logs

---

## Next Steps After Setup

Once the Key Vault is working:

1. ✅ Delete the old `key1.docx` file (or move it to secure archive)
2. ✅ Update your documentation to reference the Key Vault
3. ✅ Set a reminder to rotate the API key in 90 days
4. ✅ Consider enabling Azure Monitor alerts for Key Vault access failures

**Congratulations!** Your DigitalOwl API key is now stored securely in Azure Key Vault. 🎉
