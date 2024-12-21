// Copyright (c) 2024 ZyptAI, tim.barrow@zyptai.com
// Proprietary and confidential to ZyptAI
// File: services/fileProcessing.js
// Purpose: Handles extraction and processing of content from various file types.
//          Supports DOCX, XLSX, PDF, PPTX, CSV, and TXT files.

const mammoth = require('mammoth');
const Excel = require('exceljs');
const pdfParse = require('pdf-parse');
const { Readable } = require('stream');
const csv = require('csv-parser');
const unzipper = require('unzipper');
const xml2js = require('xml2js');
const path = require('path');
const axios = require('axios');
const { logMessage, logError } = require('./loggingService');
const { generateEmbedding } = require('../services/openAiService');
const { createSearchDocument, validateDocument } = require('../models/documentModel');
const { initializeGraphClient, getSiteInfo, getDriveInfo, getFileMetadata } = require('../services/graphService');
const { initializeSearchClient } = require('../services/searchService');
const { getRequiredConfig } = require('./configService');

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
 * Parses SharePoint URL into its components
 * @param {string} fileUrl - SharePoint URL
 * @returns {Object} URL components
 */
function parseSharePointUrl(fileUrl) {
    try {
        const url = new URL(fileUrl);
        console.log('URL parsed:', url);
        
        const tenantName = url.hostname.split('.')[0];
        console.log('Tenant:', tenantName);
        
        const sitePath = url.pathname.split('/sites/')[1].split('/')[0];
        console.log('Site path:', sitePath);
        
        const filePath = decodeURIComponent(url.pathname.split('/Shared%20Documents/')[1].split('?')[0]);
        console.log('File path:', filePath);
        
        return { tenantName, sitePath, filePath };
    } catch (error) {
        console.error('Error parsing SharePoint URL:', error);
        console.error('Original URL:', fileUrl);
        throw error;
    }
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
        const embedding = await generateEmbedding(context, chunks[i]);
        
        const document = createSearchDocument({
            fileId: fileInfo.id,
            chunkIndex: i + 1,
            fileInfo,
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
    switch (fileExtension.toLowerCase()) {
        case '.docx':
            return (await mammoth.extractRawText({ buffer })).value;
            
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

        case '.pptx':
            const zip = await unzipper.Open.buffer(buffer);
            let text = '';
            let slideCounter = 1;

            for (const file of zip.files) {
                if (file.path.startsWith('ppt/slides/slide')) {
                    const content = await file.buffer();
                    const parser = new xml2js.Parser();
                    const result = await parser.parseStringPromise(content);

                    if (result && result['p:sld'] && result['p:sld']['p:cSld']) {
                        const slideContent = extractTextFromSlide(result['p:sld']['p:cSld'][0]);
                        if (slideContent.trim()) {
                            text += `Slide ${slideCounter}: ${slideContent}\n\n`;
                            slideCounter++;
                        }
                    }
                }
            }
            return text.trim();

        case '.csv':
            return new Promise((resolve) => {
                let content = '';
                Readable.from(buffer)
                    .pipe(csv())
                    .on('data', (row) => { content += Object.values(row).join(' ') + '\n'; })
                    .on('end', () => { resolve(content); });
            });

        case '.txt':
            return buffer.toString('utf8');

        default:
            throw new Error(`Unsupported file format: ${fileExtension}`);
    }
}

/**
 * Helper function to extract text from PowerPoint slide XML
 * @private
 * @param {Object} slide - Slide XML object
 * @returns {string} Extracted text content
 */
function extractTextFromSlide(slide) {
    let text = '';
    if (slide && slide['p:spTree'] && slide['p:spTree'][0] && slide['p:spTree'][0]['p:sp']) {
        for (const shape of slide['p:spTree'][0]['p:sp']) {
            if (shape['p:txBody'] && shape['p:txBody'][0] && shape['p:txBody'][0]['a:p']) {
                for (const paragraph of shape['p:txBody'][0]['a:p']) {
                    if (paragraph['a:r'] && paragraph['a:r'][0] && paragraph['a:r'][0]['a:t']) {
                        text += paragraph['a:r'][0]['a:t'][0] + ' ';
                    }
                }
            }
        }
    }
    return text.trim();
}

/**
 * Main file processing function
 * @param {Object} context - Azure Functions context
 * @param {string} fileUrl - URL of the file to process
 * @returns {Promise<string>} Processing result message
 */
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
        
        const config = await getRequiredConfig();

        // Initialize graph client
        const credential = new ClientSecretCredential(
            config.graphAuth.tenantId,
            config.graphAuth.clientId,
            config.graphAuth.clientSecret
        );
        
        const authProvider = new TokenCredentialAuthenticationProvider(credential, {
            scopes: ['https://graph.microsoft.com/.default']
        });

        const graphClient = Client.initWithMiddleware({ authProvider });

        // Parse URL components
        const url = new URL(fileUrl);
        const tenantName = url.hostname.split('.')[0];
        const sitePath = url.pathname.split('/sites/')[1].split('/')[0];
        const filePath = decodeURIComponent(url.pathname.split('/Shared%20Documents/')[1].split('?')[0]);

        logMessage(loggingContext, "URL Components", { tenantName, sitePath, filePath });

        // Rest of your function...
    } catch (error) {
        logError(context || console, error, { operation: 'processSharePointFile', fileUrl });
        throw error;
    }
}

module.exports = {
    processSharePointFile,
    chunkContent,
    extractTextContent,
    parseSharePointUrl
};