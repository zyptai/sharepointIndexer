// Copyright (c) 2024 ZyptAI, tim.barrow@zyptai.com
// Proprietary and confidential to ZyptAI
// File: services/openAiService.js
// Purpose: Manages interactions with Azure OpenAI service for generating embeddings.

const { OpenAIClient, AzureKeyCredential } = require("@azure/openai");
const configService = require('../utils/configService');
const keyVaultService = require('./keyVaultService');
const { logMessage, logError } = require('../utils/loggingService');

// Constants for retry logic
const MAX_RETRIES = 3;
const BASE_DELAY = 1000; // 1 second base delay for exponential backoff

let openAIClient = null;

/**
 * Initialize or get the OpenAI client
 * @returns {Promise<OpenAIClient>}
 */
async function getOpenAIClient() {
    if (!openAIClient) {
        try {
            // Get configuration settings
            const endpoint = await configService.getSetting('AZURE_OPENAI_ENDPOINT');
            const apiKey = await keyVaultService.getSecret('SECRET-AZURE-OPENAI-API-KEY');

            logMessage(null, "Creating new OpenAI client with config:", {
                endpoint,
                hasApiKey: !!apiKey
            });

            if (!endpoint || !apiKey) {
                throw new Error("Missing required OpenAI configuration");
            }

            openAIClient = new OpenAIClient(
                endpoint,
                new AzureKeyCredential(apiKey)
            );
            
            logMessage(null, "OpenAI client created successfully");
        } catch (error) {
            logError(null, error, { operation: 'initializeOpenAIClient' });
            throw new Error(`Failed to initialize OpenAI client: ${error.message}`);
        }
    }
    return openAIClient;
}

/**
 * Generates embeddings for given text content with retry logic
 * @param {Object} context Azure Functions context
 * @param {string} text Text to generate embeddings for
 * @param {number} [retry=0] Current retry attempt
 * @returns {Promise<Array<number>>} Vector embedding
 * @throws {Error} If embedding generation fails after all retries
 */
async function generateEmbedding(context, text, retry = 0) {
    try {
        const client = await getOpenAIClient();
        const embeddingDeployment = await configService.getSetting('AZURE_OPENAI_EMBEDDING_DEPLOYMENT');

        logMessage(context, "Generating embeddings", {
            textLength: text.length,
            attempt: retry + 1,
            embeddingDeployment
        });

        const result = await client.getEmbeddings(embeddingDeployment, [text]);

        if (!result.data || !result.data[0].embedding) {
            throw new Error("No embedding returned from Azure OpenAI");
        }

        logMessage(context, "Embeddings generated successfully", {
            vectorLength: result.data[0].embedding.length
        });
        
        return result.data[0].embedding;
    } catch (error) {
        logError(context, error, {
            operation: 'generateEmbedding',
            textLength: text?.length,
            retry: retry
        });
        
        if (retry < MAX_RETRIES) {
            const delay = BASE_DELAY * Math.pow(2, retry);
            logMessage(context, `Retrying embedding generation after ${delay}ms`, {
                error: error.message,
                retry: retry + 1
            });
            
            await new Promise(resolve => setTimeout(resolve, delay));
            return generateEmbedding(context, text, retry + 1);
        }

        throw error;
    }
}

/**
 * Processes a batch of texts for embedding generation
 * @param {Object} context Azure Functions context
 * @param {Array<string>} texts Array of texts to process
 * @returns {Promise<Array<Array<number>>>} Array of embeddings
 */
async function generateEmbeddingBatch(context, texts) {
    logMessage(context, "Starting batch embedding generation", {
        batchSize: texts.length
    });

    const embeddings = [];
    for (let i = 0; i < texts.length; i++) {
        try {
            logMessage(context, `Processing batch item ${i + 1}/${texts.length}`);
            const embedding = await generateEmbedding(context, texts[i]);
            embeddings.push(embedding);
        } catch (error) {
            logError(context, error, {
                operation: 'generateEmbeddingBatch',
                failedIndex: i
            });
            throw error;
        }
    }

    return embeddings;
}

module.exports = {
    generateEmbedding,
    generateEmbeddingBatch
};