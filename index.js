const { app } = require('@azure/functions');
const { httpHandler, queueHandler } = require('./src/functions/SharepointIndexer');

// Register the HTTP trigger
app.http('SharepointIndexer', {
    methods: ['GET', 'POST'],
    authLevel: 'anonymous',
    handler: httpHandler
});

// Register the queue trigger
app.storageQueue('SharepointIndexerQueue', {
    queueName: 'file-processing-queue',
    connection: 'AzureWebJobsStorage',
    handler: queueHandler
});

module.exports = app;