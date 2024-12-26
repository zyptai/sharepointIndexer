// Copyright (c) 2024 ZyptAI, tim.barrow@zyptai.com
// Proprietary and confidential to ZyptAI  
// File: services/graphService.js
// Purpose: Handles all Microsoft Graph API interactions for SharePoint access.

require('isomorphic-fetch');
const { Client } = require('@microsoft/microsoft-graph-client');
const axios = require('axios');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const { ClientSecretCredential } = require('@azure/identity');
const configService = require('../utils/configService');
const keyVaultService = require('./keyVaultService');
const { logMessage, logError } = require('../utils/loggingService');

/**
 * Initialize Microsoft Graph client with authentication
 * @returns {Promise<Client>} Authenticated Graph client
 */
async function initializeGraphClient() {
    try {
        logMessage(null, "Initializing Graph client");

        const tenantId = await configService.getSetting('SHAREPOINT_BOT_TENANT_ID');
        const clientId = await configService.getSetting('SHAREPOINT_BOT_ID');
        const clientSecret = await keyVaultService.getSecret('SECRET-SHAREPOINT-BOT-PASSWORD');

        logMessage(null, "Retrieved Graph authentication config:", {
            hasTenantId: !!tenantId,
            hasClientId: !!clientId,
            hasSecret: !!clientSecret
        });

        if (!tenantId || !clientId || !clientSecret) {
            throw new Error("Missing required Graph authentication configuration");
        }

        // Create credential using client secret
        const credential = new ClientSecretCredential(
            tenantId,
            clientId,
            clientSecret
        );

        // Create authentication provider
        const authProvider = new TokenCredentialAuthenticationProvider(credential, {
            scopes: ['https://graph.microsoft.com/.default']
        });

        // Initialize Graph client
        const client = Client.initWithMiddleware({
            authProvider: authProvider,
            defaultVersion: 'v1.0'
        });

        logMessage(null, "Graph client initialized successfully");
        return client;
    } catch (error) {
        logError(null, error, { operation: 'initializeGraphClient' });
        throw new Error(`Failed to initialize Graph client: ${error.message}`);
    }
}

/**
 * Get SharePoint site information
 * @param {Object} context - Azure Functions context
 * @param {Client} graphClient - Initialized Graph client
 * @param {string} tenantName - SharePoint tenant name
 * @param {string} sitePath - Site path
 * @returns {Promise<Object>} Site information
 */
async function getSiteInfo(context, graphClient, tenantName, sitePath) {
    try {
        const siteUrl = `/sites/${tenantName}.sharepoint.com:/sites/${sitePath}`;
        logMessage(context, "Fetching site information", { siteUrl });
        
        const site = await graphClient.api(siteUrl).get();
        logMessage(context, "Site information fetched", {
            siteId: site.id,
            siteName: site.displayName
        });
        
        return site;
    } catch (error) {
        logError(context, error, { 
            operation: 'getSiteInfo',
            tenantName,
            sitePath 
        });
        throw new Error(`Failed to fetch site info: ${error.message}`);
    }
}

/**
 * Get Documents library information
 * @param {Object} context - Azure Functions context
 * @param {Client} graphClient - Initialized Graph client
 * @param {string} siteId - SharePoint site ID
 * @returns {Promise<Object>} Drive information
 */
async function getDriveInfo(context, graphClient, siteId) {
    try {
        const drivesUrl = `/sites/${siteId}/drives`;
        logMessage(context, "Fetching drives", { drivesUrl });
        
        const drives = await graphClient.api(drivesUrl).get();
        const documentLibrary = drives.value.find(drive => 
            drive.name === 'Documents' || drive.name === 'Shared Documents'
        );
        
        if (!documentLibrary) {
            throw new Error("Documents library not found in site");
        }
        
        logMessage(context, "Documents drive found", { 
            driveId: documentLibrary.id,
            driveName: documentLibrary.name 
        });
        
        return documentLibrary;
    } catch (error) {
        logError(context, error, { 
            operation: 'getDriveInfo',
            siteId 
        });
        throw new Error(`Failed to fetch drive info: ${error.message}`);
    }
}

/**
 * Get file metadata and content
 * @param {Object} context - Azure Functions context
 * @param {Client} graphClient - Initialized Graph client
 * @param {string} siteId - SharePoint site ID
 * @param {string} driveId - Drive ID
 * @param {string} filePath - File path
 * @returns {Promise<{metadata: Object, content: Buffer}>} File data
 */
async function getFileMetadata(context, graphClient, siteId, driveId, filePath) {
    try {
        const encodedFilePath = encodeURIComponent(filePath).replace(/%2F/g, '/');
        const fileUrl = `/sites/${siteId}/drives/${driveId}/root:/${encodedFilePath}`;

        logMessage(context, "Fetching file metadata", { fileUrl });

        // Get file metadata
        const file = await graphClient.api(fileUrl).get();
        
        if (!file['@microsoft.graph.downloadUrl']) {
            throw new Error("Download URL not found in file metadata");
        }

        logMessage(context, "File metadata fetched", {
            fileName: file.name,
            fileSize: file.size,
            fileId: file.id,
            mimeType: file.file?.mimeType
        });

        // Download file content
        logMessage(context, "Starting file download", {
            downloadUrl: file['@microsoft.graph.downloadUrl']
        });

        const response = await axios.get(file['@microsoft.graph.downloadUrl'], {
            responseType: 'arraybuffer',
            maxContentLength: Infinity,
            maxBodyLength: Infinity
        });

        logMessage(context, "File download complete", {
            downloadedSize: response.data.length
        });

        return {
            metadata: file,
            content: response.data
        };
    } catch (error) {
        logError(context, error, { 
            operation: 'getFileMetadata',
            siteId,
            driveId,
            filePath 
        });
        throw new Error(`Failed to fetch file: ${error.message}`);
    }
}

module.exports = {
    initializeGraphClient,
    getSiteInfo,
    getDriveInfo,
    getFileMetadata
};