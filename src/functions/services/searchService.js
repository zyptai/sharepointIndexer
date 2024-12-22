// Copyright (c) 2024 ZyptAI, tim.barrow@zyptai.com
// Proprietary and confidential to ZyptAI
// File: services/searchService.js
// Purpose: Manages Azure Cognitive Search operations for document indexing.

const { SearchClient, AzureKeyCredential } = require("@azure/search-documents");
const configService = require('../utils/configService');
const keyVaultService = require('./keyVaultService');
const { logMessage, logError } = require('../utils/loggingService');
const { validateDocument } = require('../models/documentModel');

/**
 * Initializes an Azure Cognitive Search client
 * @returns {Promise<SearchClient>} Initialized search client
 */
async function initializeSearchClient() {
    try {
        // Get configuration settings
        const endpoint = await configService.getSetting('AZURE_SEARCH_ENDPOINT');
        const indexName = await configService.getSetting('SEARCH_INDEX_NAME');
        
        // Get secret from Key Vault
        const apiKey = await keyVaultService.getSecret('SECRET-AZURE-SEARCH-KEY');

        logMessage(null, "Initializing Search client with config:", {
            endpoint,
            indexName,
            hasApiKey: !!apiKey
        });
        
        if (!endpoint || !apiKey || !indexName) {
            throw new Error("Missing required search configuration");
        }
        
        return new SearchClient(
            endpoint,
            indexName,
            new AzureKeyCredential(apiKey)
        );
    } catch (error) {
        logError(null, error, { operation: 'initializeSearchClient' });
        throw new Error(`Search client initialization failed: ${error.message}`);
    }
}

/**
 * Deletes all existing documents for a given file URL
 * @param {Object} context - Azure Functions context
 * @param {SearchClient} searchClient - Initialized search client
 * @param {string} fileUrl - URL of the file whose documents should be deleted
 * @throws {Error} If deletion fails
 */
async function deleteExistingDocuments(context, searchClient, fileUrl) {
    logMessage(context, `Deleting existing documents`, { fileUrl });
    
    try {
        const results = await searchClient.search('', { 
            filter: `fileUrl eq '${fileUrl}'`,
            select: ['docId']
        });
        
        const documentsToDelete = [];
        for await (const result of results.results) {
            documentsToDelete.push({ docId: result.document.docId });
        }

        if (documentsToDelete.length > 0) {
            await searchClient.deleteDocuments(documentsToDelete);
            logMessage(context, `Deletion complete`, {
                documentsDeleted: documentsToDelete.length,
                fileUrl
            });
        } else {
            logMessage(context, "No existing documents found to delete", { fileUrl });
        }
    } catch (error) {
        logError(context, error, { 
            operation: 'deleteExistingDocuments',
            fileUrl 
        });
        throw new Error(`Failed to delete existing documents: ${error.message}`);
    }
}

/**
 * Uploads documents to the search index
 * @param {Object} context - Azure Functions context
 * @param {SearchClient} searchClient - Initialized search client
 * @param {Array<Object>} documents - Array of documents to upload
 * @throws {Error} If upload fails
 */
async function uploadDocuments(context, searchClient, documents) {
    try {
        logMessage(context, "Validating documents before upload", { 
            documentCount: documents.length 
        });

        // Validate all documents before upload
        documents.forEach(doc => validateDocument(doc));

        const result = await searchClient.uploadDocuments(documents);
        
        const failedDocs = result.results.filter(r => !r.succeeded);
        if (failedDocs.length > 0) {
            throw new Error(`Failed to upload ${failedDocs.length} documents`);
        }

        logMessage(context, "Document upload complete", {
            uploadedCount: documents.length,
            failedCount: 0
        });
    } catch (error) {
        logError(context, error, { 
            operation: 'uploadDocuments',
            documentCount: documents.length 
        });
        throw new Error(`Failed to upload documents: ${error.message}`);
    }
}

/**
 * Performs a vector search in the index
 * @param {Object} context - Azure Functions context
 * @param {SearchClient} searchClient - Initialized search client
 * @param {Array<number>} vector - Vector to search with
 * @param {number} [top=5] - Number of results to return
 * @returns {Promise<Array>} Search results
 */
async function vectorSearch(context, searchClient, vector, top = 5) {
    try {
        const searchResults = await searchClient.search(null, {
            vector: {
                value: vector,
                fields: ["descriptionVector"],
                k: top
            },
            select: ["docId", "docTitle", "description", "fileUrl"],
            orderBy: ["@search.score desc"]
        });

        const results = [];
        for await (const result of searchResults.results) {
            results.push(result.document);
        }

        return results;
    } catch (error) {
        logError(context, error, { 
            operation: 'vectorSearch',
            vectorLength: vector.length,
            top 
        });
        throw error;
    }
}

module.exports = {
    initializeSearchClient,
    deleteExistingDocuments,
    uploadDocuments,
    vectorSearch
};