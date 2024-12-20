// Copyright (c) 2024 ZyptAI, tim.barrow@zyptai.com
// Proprietary and confidential to ZyptAI Github Deployment
// File: utils/loggingService.js
// Purpose: Centralized logging service that integrates with Application Insights
//          and provides consistent logging across the application.

const appInsights = require('applicationinsights');

let appInsightsClient;

/**
 * Initializes Application Insights if not already initialized
 * @private
 */
function initializeLogging() {
    if (!appInsightsClient) {
        appInsights.setup(process.env.APP_INSIGHTS_CONNECTION_STRING)
            .setAutoDependencyCorrelation(true)
            .setAutoCollectRequests(true)
            .setAutoCollectPerformance(true)
            .setAutoCollectExceptions(true)
            .setAutoCollectDependencies(true)
            .setAutoCollectConsole(true)
            .setUseDiskRetryCaching(true)
            .start();
        
        appInsightsClient = appInsights.defaultClient;
    }
}

/**
 * Logs a message to both context.log and Application Insights
 * @param {Object} context - Azure Functions context object or console
 * @param {string} message - Message to log
 * @param {Object} [obj] - Optional object with additional properties to log
 */
function logMessage(context, message, obj = null) {
    // Initialize app insights if needed
    initializeLogging();
    
    // If context is null/undefined, use console
    const logger = context || console;
    
    if (obj) {
        // Log to context/console
        if (logger.log) {
            logger.log(message, obj);
        } else {
            logger.info?.(message, obj) || logger.debug?.(message, obj) || console.log(message, obj);
        }
        
        // Log to app insights
        appInsightsClient.trackTrace({ 
            message: message, 
            properties: obj,
            severity: 1 
        });
    } else {
        // Log to context/console
        if (logger.log) {
            logger.log(message);
        } else {
            logger.info?.(message) || logger.debug?.(message) || console.log(message);
        }
        
        // Log to app insights
        appInsightsClient.trackTrace({ 
            message: message,
            severity: 1
        });
    }
}

/**
 * Logs an error to both context.log and Application Insights
 * @param {Object} context - Azure Functions context object
 * @param {Error} error - Error object to log
 * @param {Object} [extraProperties] - Optional additional properties to log with the error
 */
function logError(context, error, extraProperties = {}) {
    initializeLogging();
    
    context.log.error?.(error) ?? context.log(`Error: ${error.message}`);
    appInsightsClient.trackException({ 
        exception: error,
        properties: {
            ...extraProperties,
            stack: error.stack,
            name: error.name
        },
        severity: 3 // Error
    });
}

module.exports = {
    logMessage,
    logError
};