# Azure Deployment Guide - Finest Travel Voucher System

This guide walks you through deploying the Voucher System to Azure App Service using Docker containers.

## Prerequisites

1. **Azure Account** - [Create one here](https://azure.microsoft.com/free/)
2. **Azure CLI** - [Install instructions](https://docs.microsoft.com/cli/azure/install-azure-cli)
3. **Docker Desktop** (for local testing) - [Download](https://www.docker.com/products/docker-desktop)

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────┐
│                      Azure Cloud                             │
│  ┌─────────────────┐    ┌──────────────────────────────┐   │
│  │ Azure Container │───▶│     Azure App Service        │   │
│  │    Registry     │    │  (Docker Container)          │   │
│  │  (ACR)          │    │  - FastAPI App               │   │
│  └─────────────────┘    │  - LibreOffice (PDF conv.)   │   │
│           ▲             └──────────────────────────────┘   │
│           │                          │                      │
│  ┌────────┴────────┐                 ▼                      │
│  │ GitHub Actions  │         ┌──────────────┐              │
│  │ (CI/CD)         │         │   Users      │              │
│  └─────────────────┘         │   Browser    │              │
│           ▲                  └──────────────┘              │
└───────────┼─────────────────────────────────────────────────┘
            │
     ┌──────┴──────┐
     │   GitHub    │
     │ Repository  │
     └─────────────┘
```

## Step 1: Create Azure Resources

Open a terminal and run:

```bash
# Login to Azure
az login

# Set variables (customize these)
RESOURCE_GROUP="finest-travel-rg"
LOCATION="westeurope"
ACR_NAME="finesttravel"  # Must be globally unique, lowercase
APP_NAME="finest-travel-vouchers"  # Must be globally unique
APP_SERVICE_PLAN="finest-travel-plan"

# Create Resource Group
az group create --name $RESOURCE_GROUP --location $LOCATION

# Create Azure Container Registry (ACR)
az acr create \
  --resource-group $RESOURCE_GROUP \
  --name $ACR_NAME \
  --sku Basic \
  --admin-enabled true

# Get ACR credentials (save these!)
az acr credential show --name $ACR_NAME

# Create App Service Plan (B1 is minimum for containers)
az appservice plan create \
  --name $APP_SERVICE_PLAN \
  --resource-group $RESOURCE_GROUP \
  --sku B2 \
  --is-linux

# Create Web App for Containers
az webapp create \
  --resource-group $RESOURCE_GROUP \
  --plan $APP_SERVICE_PLAN \
  --name $APP_NAME \
  --deployment-container-image-name $ACR_NAME.azurecr.io/vouchersystem:latest

# Configure Web App to use ACR
az webapp config container set \
  --name $APP_NAME \
  --resource-group $RESOURCE_GROUP \
  --docker-custom-image-name $ACR_NAME.azurecr.io/vouchersystem:latest \
  --docker-registry-server-url https://$ACR_NAME.azurecr.io

# Enable ACR pull for App Service
az webapp identity assign \
  --name $APP_NAME \
  --resource-group $RESOURCE_GROUP

# Configure continuous deployment from ACR
az webapp deployment container config \
  --name $APP_NAME \
  --resource-group $RESOURCE_GROUP \
  --enable-cd true
```

## Step 2: Build and Push Docker Image (First Time)

```bash
# Login to ACR
az acr login --name $ACR_NAME

# Build the Docker image
docker build -t $ACR_NAME.azurecr.io/vouchersystem:latest .

# Push to ACR
docker push $ACR_NAME.azurecr.io/vouchersystem:latest

# Verify image is in ACR
az acr repository list --name $ACR_NAME --output table
```

## Step 3: Configure GitHub Actions (Automated Deployments)

### 3.1 Create Azure Service Principal

```bash
# Create service principal for GitHub Actions
az ad sp create-for-rbac \
  --name "github-actions-vouchersystem" \
  --role contributor \
  --scopes /subscriptions/$(az account show --query id -o tsv)/resourceGroups/$RESOURCE_GROUP \
  --sdk-auth

# Save the JSON output - you'll need it for GitHub secrets
```

### 3.2 Add GitHub Secrets

Go to your GitHub repository → Settings → Secrets and variables → Actions

Add these secrets:

| Secret Name | Value |
|-------------|-------|
| `AZURE_CREDENTIALS` | The full JSON from service principal creation |
| `ACR_USERNAME` | ACR admin username (from `az acr credential show`) |
| `ACR_PASSWORD` | ACR admin password |

### 3.3 Update Workflow Variables

Edit `.github/workflows/azure-deploy.yml` and update:
- `AZURE_WEBAPP_NAME`: Your App Service name
- `REGISTRY_NAME`: Your ACR name (without .azurecr.io)

## Step 4: Configure App Settings

```bash
# Set environment variables
az webapp config appsettings set \
  --name $APP_NAME \
  --resource-group $RESOURCE_GROUP \
  --settings \
    WEBSITES_PORT=8000 \
    VOUCHER_TEMPLATE_PATH=/app/templates/_Voucher\ blank.docx
```

## Step 5: Test Deployment

```bash
# Get your app URL
echo "https://$APP_NAME.azurewebsites.net"

# Check logs if issues
az webapp log tail --name $APP_NAME --resource-group $RESOURCE_GROUP
```

## Estimated Costs

| Resource | SKU | Estimated Monthly Cost |
|----------|-----|------------------------|
| App Service Plan | B2 (2 cores, 3.5 GB RAM) | ~$55/month |
| Container Registry | Basic | ~$5/month |
| **Total** | | **~$60/month** |

> 💡 **Cost Optimization**: For lower usage, B1 (~$14/month) may work, but B2 is recommended for LibreOffice PDF conversion performance.

## Troubleshooting

### Container won't start
```bash
# Check container logs
az webapp log download --name $APP_NAME --resource-group $RESOURCE_GROUP

# Enable detailed logging
az webapp log config --name $APP_NAME --resource-group $RESOURCE_GROUP \
  --docker-container-logging filesystem
```

### PDF conversion failing
LibreOffice needs sufficient memory. Ensure you're using B2 or higher plan.

### File upload issues
Default max upload is 100MB. To increase:
```bash
az webapp config appsettings set \
  --name $APP_NAME \
  --resource-group $RESOURCE_GROUP \
  --settings WEBSITES_CONTAINER_START_TIME_LIMIT=600
```

## Alternative: Quick Deploy with Azure CLI

For quick testing without GitHub Actions:

```bash
# One-command deployment (after ACR setup)
az acr build --registry $ACR_NAME --image vouchersystem:latest .

# Restart app to pull new image
az webapp restart --name $APP_NAME --resource-group $RESOURCE_GROUP
```

## Security Recommendations

1. **Enable HTTPS Only**
   ```bash
   az webapp update --name $APP_NAME --resource-group $RESOURCE_GROUP --https-only true
   ```

2. **Add IP Restrictions** (if internal use only)
   ```bash
   az webapp config access-restriction add \
     --name $APP_NAME \
     --resource-group $RESOURCE_GROUP \
     --rule-name "OfficeOnly" \
     --priority 100 \
     --ip-address YOUR_OFFICE_IP/32
   ```

3. **Enable Authentication** (optional)
   Azure AD authentication can be enabled via Azure Portal → App Service → Authentication

## Support

For issues with deployment, check:
- Azure Portal → App Service → Diagnose and solve problems
- Application Insights (if configured)
- Container logs via Azure CLI

