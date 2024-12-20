// Copyright (c) 2024 ZyptAI, tim.barrow@zyptai.com
// Proprietary and confidential to ZyptAI
// File: utils/configService.js

const { AppConfigurationClient } = require("@azure/app-configuration");
const { SecretClient } = require("@azure/keyvault-secrets");
const { DefaultAzureCredential } = require("@azure/identity");

let config = null;

/**
 * Helper function to determine if a value is base64 encoded
 * @param {string} str The string to test
 * @returns {boolean}
 */
function isBase64(str) {
    if (typeof str !== 'string') return false;
    const base64Regex = /^[A-Za-z0-9+/]*={0,2}$/;
    return base64Regex.test(str) && str.length % 4 === 0;
}

/**
 * Decrypts values that might be encrypted 
 * @param {string} value The value to potentially decrypt
 * @returns {string} The decrypted value
 */
function decryptIfNeeded(value) {
    if (!value) return value;
    
    // If it has crypto_ prefix, remove it
    if (value.startsWith('crypto_')) {
        return value.substring(7); // Remove 'crypto_' prefix
    }
    
    return value;
}

/**
 * Retrieves all configuration settings
 * @returns {Promise<Object>} All configuration settings
 */
async function getRequiredConfig() {
    if (config) {
        console.log("Returning cached config:", config);
        return config;
    }

    console.log("Starting configuration retrieval...");
    
    try {
        // First get all App Config settings
        console.log("App Config Connection String:", process.env.AZURE_APP_CONFIG_CONNECTION_STRING);
        const appConfigClient = new AppConfigurationClient(process.env.AZURE_APP_CONFIG_CONNECTION_STRING);
        const settings = {};

        const settingsIterator = appConfigClient.listConfigurationSettings();
        for await (const setting of settingsIterator) {
            settings[setting.key] = setting.value;
            console.log(`Retrieved setting ${setting.key}:`, setting.value);
        }

        // Get Key Vault secrets following the sample approach exactly
        const keyVaultName = "zyptaidevkeyvault";
        const keyVaultUri = `https://${keyVaultName}.vault.azure.net`;
        console.log("Key Vault URI:", keyVaultUri);
        
        // Create credential and client exactly as shown in sample
        const credential = new DefaultAzureCredential();
        const secretClient = new SecretClient(keyVaultUri, credential);

        // Get secrets with exact names from Key Vault
        const secretNames = [
            'SECRET-AZURE-OPENAI-API-KEY',
            'SECRET-AZURE-SEARCH-KEY',
            'SECRET-BOT-PASSWORD',
            'SECRET-JIRA-API-TOKEN'
        ];

        // Retrieve secrets
        const secrets = {};
        for (const secretName of secretNames) {
            const secret = await secretClient.getSecret(secretName);
            secrets[secretName] = decryptIfNeeded(secret.value);
            console.log(`Retrieved secret ${secretName}:`, secret.value ? "present" : "missing");
        }

        // Map all settings and secrets to our config structure
        config = {
            search: {
                endpoint: settings['SEARCH_ENDPOINT'],
                apiKey: secrets['SECRET-AZURE-SEARCH-KEY'],
                indexName: settings['SEARCH_INDEX_NAME']
            },
            graphAuth: {
                tenantId: settings['BOT_TENANT_ID'],
                clientId: settings['BOT_ID'],
                clientSecret: secrets['SECRET-BOT-PASSWORD']
            },
            storage: {
                connectionString: settings['AZURE_STORAGE_CONNECTION_STRING']
            },
            appInsights: {
                connectionString: settings['APP_INSIGHTS_CONNECTION_STRING']
            },
            openai: {
                endpoint: settings['AZURE_OPENAI_ENDPOINT'],
                apiKey: secrets['SECRET-AZURE-OPENAI-API-KEY'],
                embeddingDeployment: settings['AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME']
            },
            jira: {
                baseUrl: settings['JIRA_BASE_URL'],
                username: settings['JIRA_USERNAME'],
                apiToken: secrets['SECRET-JIRA-API-TOKEN']
            }
        };

        console.log("Final Search config:", {
            endpoint: config.search.endpoint,
            indexName: config.search.indexName,
            apiKey: config.search.apiKey ? "present" : "missing"
        });

        console.log("Final Graph Auth config:", {
            tenantId: config.graphAuth.tenantId,
            clientId: config.graphAuth.clientId,
            clientSecret: config.graphAuth.clientSecret ? "present" : "missing"
        });

        console.log("Final Storage config:", {
            connectionString: config.storage.connectionString ? "present" : "missing"
        });

        console.log("Final OpenAI config:", {
            endpoint: config.openai.endpoint,
            apiKey: config.openai.apiKey ? "present" : "missing",
            embeddingDeployment: config.openai.embeddingDeployment
        });

        console.log("Final Jira config:", {
            baseUrl: config.jira.baseUrl,
            username: config.jira.username,
            apiToken: config.jira.apiToken ? "present" : "missing"
        });

        return config;
    } catch (error) {
        console.error("Failed to retrieve configuration:", error);
        throw error;
    }
}

module.exports = {
    getRequiredConfig
};