// Copyright (c) 2024 ZyptAI, tim.barrow@zyptai.com
// Proprietary and confidential to ZyptAI
// File: services/graphService.js
// Purpose: Handles all Microsoft Graph API interactions for SharePoint access.
//          Provides methods for site, drive, and file operations.

const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
const { ClientSecretCredential } = require('@azure/identity');
const axios = require('axios');
const { getRequiredConfig } = require('../utils/configService');
const { logMessage, logError } = require('../utils/loggingService');
require('isomorphic-fetch'); // Required for MS Graph Client
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
const { ClientSecretCredential } = require("@azure/identity");  // Add this line

/**
 * Initializes a Microsoft Graph client with proper authentication
 * @returns {Promise<Client>} Authenticated Graph client
 * @throws {Error} If authentication fails
 */
async function initializeGraphClient() {
    const config = await getRequiredConfig();
    
    logMessage(null, "Initializing Graph client with config", {
        tenantId: config.graphAuth.tenantId,
        clientId: config.graphAuth.clientId,
        clientSecret: config.graphAuth.clientSecret ? "present" : "missing"
    });
    
    // Create credential from configured values
    const credential = new ClientSecretCredential(
        config.graphAuth.tenantId,
        config.graphAuth.clientId,
        config.graphAuth.clientSecret
    );
    
    // Create auth provider with required Graph API scope
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
        scopes: ['https://graph.microsoft.com/.default']
    });
    
    return Client.initWithMiddleware({ authProvider });
}

/**
 * Retrieves SharePoint site information
 * @param {Object} context - Azure Functions context
 * @param {Client} graphClient - Initialized Graph client
 * @param {string} tenantName - SharePoint tenant name
 * @param {string} sitePath - Path to the SharePoint site
 * @returns {Promise<Object>} Site information including ID and name
 * @throws {Error} If site cannot be found or accessed
 */
async function getSiteInfo(context, graphClient, tenantName, sitePath) {
    const siteUrl = `/sites/${tenantName}.sharepoint.com:/sites/${sitePath}`;
    logMessage(context, "Fetching site information", { siteUrl });
    
    try {
        const site = await graphClient.api(siteUrl).get();
        logMessage(context, "Site information fetched", {
            siteId: site.id,
            siteName: site.displayName
        });
        return site;
    } catch (error) {
        logError(context, error, { 
            operation: 'getSiteInfo',
            siteUrl,
            tenantName,
            sitePath 
        });
        throw new Error(`Failed to fetch site info: ${error.message}`);
    }
}

/**
 * Retrieves the Documents library drive from a SharePoint site
 * @param {Object} context - Azure Functions context
 * @param {Client} graphClient - Initialized Graph client
 * @param {string} siteId - SharePoint site ID
 * @returns {Promise<Object>} Drive information for the Documents library
 * @throws {Error} If Documents library cannot be found or accessed
 */
async function getDriveInfo(context, graphClient, siteId) {
    const drivesUrl = `/sites/${siteId}/drives`;
    logMessage(context, "Fetching drives", { drivesUrl });
    
    try {
        const drives = await graphClient.api(drivesUrl).get();
        const documentLibrary = drives.value.find(drive => drive.name === 'Documents');
        
        if (!documentLibrary) {
            throw new Error("Documents library not found in site");
        }
        
        logMessage(context, "Documents drive found", { driveId: documentLibrary.id });
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
 * Retrieves metadata and downloads a specific file from SharePoint
 * @param {Object} context - Azure Functions context
 * @param {Client} graphClient - Initialized Graph client
 * @param {string} siteId - SharePoint site ID
 * @param {string} driveId - Drive ID
 * @param {string} filePath - Path to the file within the drive
 * @returns {Promise<{metadata: Object, content: Buffer}>} File metadata and content
 * @throws {Error} If file cannot be found or accessed
 */
async function getFileMetadata(context, graphClient, siteId, driveId, filePath) {
    const encodedFilePath = encodeURIComponent(filePath).replace(/%2F/g, '/');
    const fileUrl = `/sites/${siteId}/drives/${driveId}/root:/${encodedFilePath}`;

    try {
        // Get file metadata
        const file = await graphClient.api(fileUrl).get();
        logMessage(context, "File metadata fetched", {
            fileName: file.name,
            fileSize: file.size,
            fileId: file.id,
            mimeType: file.file ? file.file.mimeType : 'Unknown'
        });

        // Download file content
        logMessage(context, "Starting file download", {
            downloadUrl: file['@microsoft.graph.downloadUrl'],
            expectedSize: file.size
        });

        const response = await axios.get(file['@microsoft.graph.downloadUrl'], {
            responseType: 'arraybuffer',
            maxContentLength: Infinity,
            maxBodyLength: Infinity
        });

        logMessage(context, "File download complete", {
            downloadedSize: response.data.length,
            matched: response.data.length === file.size
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
            filePath,
            errorDetails: error.response?.data ? JSON.stringify(error.response.data) : 'No response data'
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