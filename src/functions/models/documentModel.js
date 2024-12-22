// Copyright (c) 2024 ZyptAI, tim.barrow@zyptai.com
// Proprietary and confidential to ZyptAI
// File: models/documentModel.js
// Purpose: Defines the document structure and provides validation for search documents.
//          Creates consistent document objects for Azure Cognitive Search.

const path = require('path');

/**
 * Creates a document chunk for search indexing
 * @param {Object} params Document creation parameters
 * @param {string} params.fileId Original file ID
 * @param {number} params.chunkIndex Index of this chunk
 * @param {Object} params.fileInfo File metadata from SharePoint
 * @param {string} params.content Chunk text content
 * @param {Array<number>} params.embedding Vector embedding of content
 * @param {number} params.totalChunks Total number of chunks
 * @returns {Object} Formatted search document
 */
function createSearchDocument({
    fileId,
    chunkIndex,
    fileInfo,
    content,
    embedding,
    totalChunks
}) {
    if (!fileId || !fileInfo || !content || !embedding) {
        throw new Error('Missing required parameters for document creation');
    }

    // Convert lastmodified to proper ISO string if it isn't already
    const lastModified = new Date(fileInfo.lastModifiedDateTime).toISOString();

    return {
        // Unique identifier for this chunk
        docId: `${fileId}-${chunkIndex}`,
        
        // File metadata
        docTitle: fileInfo.name,
        filename: fileInfo.name,
        filetype: path.extname(fileInfo.name).toLowerCase(),
        fileUrl: fileInfo.webUrl,
        lastmodified: lastModified,
        
        // Chunk information
        description: content,
        chunkindex: parseInt(chunkIndex), // Convert to integer for Edm.Int32
        totalChuncks: parseInt(totalChunks), // Match schema spelling and type
        
        // Vector embedding
        descriptionVector: embedding
    };
}

/**
 * Validates a search document before indexing
 * @param {Object} document Search document to validate
 * @returns {boolean} True if document is valid
 * @throws {Error} If document is invalid
 */
function validateDocument(document) {
    const requiredFields = [
        'docId',
        'docTitle',
        'description',
        'filename',
        'filetype',
        'lastmodified',
        'chunkindex',
        'totalChuncks',
        'descriptionVector',
        'fileUrl'
    ];

    const missingFields = requiredFields.filter(field => !document[field]);
    
    if (missingFields.length > 0) {
        throw new Error(`Invalid document: missing required fields: ${missingFields.join(', ')}`);
    }

    // Type validations
    if (!Array.isArray(document.descriptionVector)) {
        throw new Error('Invalid document: descriptionVector must be an array');
    }

    if (typeof document.chunkindex !== 'number') {
        throw new Error('Invalid document: chunkindex must be a number');
    }

    if (typeof document.totalChuncks !== 'number') {
        throw new Error('Invalid document: totalChuncks must be a number');
    }

    if (isNaN(new Date(document.lastmodified).getTime())) {
        throw new Error('Invalid document: lastmodified must be a valid date');
    }

    return true;
}

module.exports = {
    createSearchDocument,
    validateDocument
};