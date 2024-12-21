// Copyright (c) 2024 ZyptAI, tim.barrow@zyptai.com
// Proprietary and confidential to ZyptAI
// File: services/fileProcessing.js
// Purpose: Handles extraction and processing of content from various file types.
//          Supports DOCX, XLSX, PDF, PPTX, CSV, and TXT files.

const mammoth = require('mammoth');
const Excel = require('exceljs');
const pdfParse = require('pdf-parse');
const path = require('path');
const axios = require('axios');
const { logMessage, logError } = require('./loggingService');
const { generateEmbedding } = require('../services/openAiService');
const { createSearchDocument, validateDocument } = require('../models/documentModel');
const { initializeGraphClient, getSiteInfo, getDriveInfo, getFileMetadata } = require('../services/graphService');
const { initializeSearchClient, deleteExistingDocuments } = require('../services/searchService');

/**
 * Controls the maximum size of content chunks for processing
 * @constant {number}
 */
const MAX_CHUNK_SIZE = 2000;

/**
 * Splits content into chunks while preserving sentence boundaries
 * @param {Object} context - Azure Functions context
 * @param {string} content - Text content to chunk
 * @param {number} [maxChunkSize=2000] - Maximum size of each chunk
 * @returns {Array<string>} Array of content chunks
 */
function chunkContent(context, content, maxChunkSize = MAX_CHUNK_SIZE) {
    logMessage(context, "Starting content chunking", {
        contentLength: content.length,
        maxChunkSize
    });

    const chunks = [];
    let currentChunk = "";
    
    // Split into sentences using regex
    const sentences = content.match(/[^.!?]+[.!?]+|\s+/g) || [];

    for (const sentence of sentences) {
        if ((currentChunk + sentence).length > maxChunkSize && currentChunk.length > 0) {
            chunks.push(currentChunk.trim());
            currentChunk = "";
        }
        currentChunk += sentence;
    }

    if (currentChunk.trim().length > 0) {
        chunks.push(currentChunk.trim());
    }

    logMessage(context, "Content chunking complete", {
        totalChunks: chunks.length,
        averageChunkSize: Math.round(chunks.reduce((acc, chunk) => acc + chunk.length, 0) / chunks.length)
    });

    return chunks;
}

/**
 * Processes chunks and creates search documents
 * @param {Object} context - Azure Functions context
 * @param {Array<string>} chunks - Content chunks
 * @param {Object} fileInfo - File metadata
 * @param {Object} searchClient - Search client instance
 * @param {string} fileUrl - Original file URL
 * @returns {Promise<Array>} Processed documents
 */
async function processChunks(context, chunks, fileInfo, searchClient, fileUrl) {
    const documents = [];
    
    for (let i = 0; i < chunks.length; i++) {
        try {
            const embedding = await generateEmbedding(context, chunks[i]);
            
            const document = createSearchDocument({
                fileId: fileInfo.id,
                chunkIndex: i + 1,
                fileInfo: {
                    ...fileInfo,
                    webUrl: fileUrl
                },
                content: chunks[i],
                embedding,
                totalChunks: chunks.length
            });

            validateDocument(document);
            documents.push(document);
            
            logMessage(context, `Processed chunk ${i + 1}/${chunks.length}`, {
                docId: document.docId,
                contentLength: chunks[i].length
            });
        } catch (error) {
            logError(context, error, {
                operation: 'processChunks',
                chunkIndex: i,
                fileUrl
            });
            throw error;
        }
    }

    await searchClient.uploadDocuments(documents);
    return documents;
}

/**
 * Extracts text content based on file type
 * @param {Object} context - Azure Functions context
 * @param {string} fileExtension - File extension
 * @param {Buffer} buffer - File content buffer
 * @returns {Promise<string>} Extracted text content
 */
async function extractTextContent(context, fileExtension, buffer) {
    logMessage(context, "Starting text extraction", { fileExtension });

    try {
        switch (fileExtension.toLowerCase()) {
            case '.docx':
                const result = await mammoth.extractRawText({ buffer });
                return result.value;
                
            case '.xlsx': {
                const workbook = new Excel.Workbook();
                await workbook.xlsx.load(buffer);
                let content = '';
                workbook.worksheets.forEach(worksheet => {
                    worksheet.eachRow((row) => {
                        content += row.values.slice(1).join(' ') + '\n';
                    });
                });
                return content;
            }

            case '.pdf':
                return (await pdfParse(buffer)).text;

            case '.txt':
                return buffer.toString('utf8');

            default:
                throw new Error(`Unsupported file format: ${fileExtension}`);
        }
    } catch (error) {
        logError(context, error, {
            operation: 'extractTextContent',
            fileExtension
        });
        throw error;
    }
}

/**
 * Main file processing function
 * @param {Object} context - Azure Functions context
 * @param {string} fileUrl - URL of the file to process
 * @returns {Promise<string>} Processing result message
 */
async function processSharePointFile(context, fileUrl) {
    try {
        // If no context is provided, create a minimal context for logging
        const loggingContext = context || { 
            log: (msg) => console.log(msg) 
        }; 

        logMessage(loggingContext, "Starting file processing", { fileUrl });

        // Initialize graph client using the updated service
        const graphClient = await initializeGraphClient();

        // Parse URL components
        const url = new URL(fileUrl);
        const tenantName = url.hostname.split('.')[0];
        const sitePath = url.pathname.split('/sites/')[1].split('/')[0];
        const filePath = decodeURIComponent(url.pathname.split('/Shared%20Documents/')[1].split('?')[0]);

        logMessage(loggingContext, "URL Components", { 
            tenantName, 
            sitePath, 
            filePath 
        });

        // Get site information
        const site = await getSiteInfo(loggingContext, graphClient, tenantName, sitePath);
        logMessage(loggingContext, "Retrieved site info", { siteId: site.id });

        // Get drive information (document library)
        const drive = await getDriveInfo(loggingContext, graphClient, site.id);
        logMessage(loggingContext, "Retrieved drive info", { driveId: drive.id });

        // Get file metadata and content
        const { metadata, content } = await getFileMetadata(
            loggingContext, 
            graphClient, 
            site.id, 
            drive.id, 
            filePath
        );
        logMessage(loggingContext, "Retrieved file", { 
            fileName: metadata.name,
            fileSize: content.length 
        });

        // Extract text content based on file type
        const fileExtension = path.extname(metadata.name).toLowerCase();
        const textContent = await extractTextContent(loggingContext, fileExtension, content);
        logMessage(loggingContext, "Extracted text content", { 
            contentLength: textContent.length 
        });

        // Split content into chunks
        const chunks = chunkContent(loggingContext, textContent);
        logMessage(loggingContext, "Content chunked", { 
            numberOfChunks: chunks.length 
        });

        // Initialize search client
        const searchClient = await initializeSearchClient();

        // Delete any existing documents for this file
        await deleteExistingDocuments(loggingContext, searchClient, fileUrl);

        // Process and index chunks
        const documents = await processChunks(
            loggingContext,
            chunks,
            metadata,
            searchClient,
            fileUrl
        );

        logMessage(loggingContext, "File processing complete", {
            fileUrl,
            chunksProcessed: documents.length
        });

        return `Successfully processed ${metadata.name} into ${documents.length} chunks`;
    } catch (error) {
        logError(context || console, error, { 
            operation: 'processSharePointFile', 
            fileUrl 
        });
        throw error;
    }
}

module.exports = {
    processSharePointFile,
    chunkContent,
    extractTextContent,
    MAX_CHUNK_SIZE
};