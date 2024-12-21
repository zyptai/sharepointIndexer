// Copyright (c) 2024 ZyptAI, tim.barrow@zyptai.com
// Proprietary and confidential to ZyptAI

const { processSharePointFile } = require('./utils/fileProcessors');
const { logMessage, logError } = require('./utils/loggingService');
const { getRequiredConfig } = require('./utils/configService');  // ✓ correct
const { TextDecoder } = require('util');

/**
 * Parses the request body from various possible formats
 */
async function parseRequestBody(body) {
    if (typeof body === 'string') {
        return JSON.parse(body);
    }
    
    if (body && typeof body.getReader === 'function') {
        const reader = body.getReader();
        const decoder = new TextDecoder();
        let result = '';
        
        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            result += decoder.decode(value, { stream: true });
        }
        
        return JSON.parse(result);
    }
    
    return body;
}

/**
 * HTTP trigger handler
 */
async function httpHandler(request, context) {
    try {
        logMessage(context, "Received indexing request");
        
        // Parse request body
        const requestBody = await parseRequestBody(request.body);
        logMessage(context, "Parsed request body", { body: requestBody });
        
        // Get file URL from query params, body, or default config
        const fileUrl = request.query.fileUrl || 
                       requestBody?.fileUrl || 
                       await getRequiredConfig('DEFAULT_SHAREPOINT_FILE_PATH');
        
        if (!fileUrl) {
            throw new Error("No file URL provided and no default URL configured");
        }
        
        logMessage(context, "Processing file", { fileUrl });
        const result = await processSharePointFile(context, fileUrl);
        
        return { 
            status: 200,
            body: {
                message: result,
                fileUrl: fileUrl
            }
        };
    } catch (error) {
        logError(context, error);
        
        return { 
            status: 500, 
            body: {
                error: "Internal Server Error",
                message: error.message
            }
        };
    }
}

/**
 * Queue trigger handler
 */
async function queueHandler(queueItem, context) {
    try {
        logMessage(context, "Processing queue item", { item: queueItem });
        
        const fileUrl = typeof queueItem === 'string' ? 
                       queueItem : 
                       queueItem?.fileUrl;
        
        if (!fileUrl) {
            throw new Error("No file URL found in queue item");
        }
        
        await processSharePointFile(context, fileUrl);
        logMessage(context, "Queue item processing complete", { fileUrl });
    } catch (error) {
        logError(context, error, { queueItem });
        throw error; // Allows the queue to handle retry logic
    }
}

module.exports = {
    httpHandler,
    queueHandler
};