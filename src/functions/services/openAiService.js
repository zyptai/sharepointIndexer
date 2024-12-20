// Copyright (c) 2024 ZyptAI, tim.barrow@zyptai.com
// Proprietary and confidential to ZyptAI
// File: services/openAiService.js
// Purpose: Manages interactions with Azure OpenAI service for generating embeddings.
//          Handles rate limiting and error recovery for embedding operations.

const { OpenAIClient, AzureKeyCredential } = require("@azure/openai");
const { getRequiredConfig } = require('../utils/configService');
const { logMessage, logError } = require('../utils/loggingService');

/**
 * Maximum retries for embedding generation
 * @constant {number}
 */
const MAX_RETRIES = 3;

/**
 * Base delay for exponential backoff (in ms)
 * @constant {number}
 */
const BASE_DELAY = 1000;

let openAIClient = null;

/**
 * Initialize or get the OpenAI client
 * @returns {Promise<OpenAIClient>}
 */
async function getOpenAIClient() {
    if (!openAIClient) {
        const config = await getRequiredConfig();
        
        console.log("Creating new OpenAI client with config:", {
            endpoint: config.openai.endpoint,
            apiKey: config.openai.apiKey,
            embeddingDeployment: config.openai.embeddingDeployment,
            apiKeyLength: config.openai.apiKey?.length
        });

        if (!config.openai.endpoint || !config.openai.apiKey) {
            throw new Error(`Missing OpenAI configuration - endpoint: ${!!config.openai.endpoint}, apiKey: ${!!config.openai.apiKey}`);
        }

        try {
            openAIClient = new OpenAIClient(
                config.openai.endpoint,
                new AzureKeyCredential(config.openai.apiKey)
            );
            
            console.log("OpenAI client created successfully");
        } catch (error) {
            console.error("Failed to create OpenAI client:", error);
            throw error;
        }
    } else {
        console.log("Returning existing OpenAI client");
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
        const config = await getRequiredConfig();
        
        console.log("Generating embedding with config:", {
            endpoint: config.openai.endpoint,
            apiKey: config.openai.apiKey,
            embeddingDeployment: config.openai.embeddingDeployment,
            textLength: text.length,
            attempt: retry + 1
        });

        const client = await getOpenAIClient();

        logMessage(context, "Generating embeddings", {
            textLength: text.length,
            attempt: retry + 1
        });

        console.log("Making API call to OpenAI with deployment:", config.openai.embeddingDeployment);
        const result = await client.getEmbeddings(
            config.openai.embeddingDeployment,
            [text]
        );

        if (!result.data || !result.data[0].embedding) {
            throw new Error("No embedding returned from Azure OpenAI");
        }

        logMessage(context, "Embeddings generated successfully", {
            vectorLength: result.data[0].embedding.length
        });
        
        return result.data[0].embedding;
    } catch (error) {
        console.error("Error in generateEmbedding:", error);
        
        if (retry < MAX_RETRIES) {
            const delay = BASE_DELAY * Math.pow(2, retry);
            logMessage(context, `Retrying embedding generation after ${delay}ms`, {
                error: error.message,
                retry: retry + 1
            });
            
            await new Promise(resolve => setTimeout(resolve, delay));
            return generateEmbedding(context, text, retry + 1);
        }

        logError(context, error, {
            operation: 'generateEmbedding',
            textLength: text.length,
            finalAttempt: retry + 1
        });
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
            console.log(`Processing batch item ${i + 1}/${texts.length}`);
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