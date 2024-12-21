// Copyright (c) 2024 ZyptAI, tim.barrow@zyptai.com
// Proprietary and confidential to ZyptAI
// File: utils/configService.js
// Purpose: Centralized configuration service that integrates with App Configuration
//          and Key Vault to provide consistent configuration across the application.

const { AppConfigurationClient } = require("@azure/app-configuration");
const { SecretClient } = require("@azure/keyvault-secrets");
const { ManagedIdentityCredential, AzureCliCredential } = require("@azure/identity");

// Configuration cache
let configCache = null;

/**
 * Configuration service class to handle all configuration operations
 */
class ConfigurationService {
    constructor() {
        this.initialized = false;
        this.config = null;
    }

    /**
     * Initialize the configuration service
     * @returns {Promise<void>}
     */
    async initialize() {
        if (this.initialized) {
            return;
        }

        try {
            // Create initial AppConfigurationClient
            const appConfigClient = new AppConfigurationClient(process.env.AZURE_APP_CONFIG_CONNECTION_STRING);
            const settings = await this._getAllSettings(appConfigClient);

            // Get managed identity details
            const managedIdentityId = settings.SHAREPOINT_INDEXER_MANAGED_IDENTITY_ID;
            const keyVaultName = settings.KEY_VAULT_NAME;

            if (!managedIdentityId) {
                throw new Error("Missing SHAREPOINT_INDEXER_MANAGED_IDENTITY_ID in App Configuration");
            }

            if (!keyVaultName) {
                throw new Error("Missing KEY_VAULT_NAME in App Configuration");
            }

            // Create credential based on environment
            const credential = this._createCredential(managedIdentityId);

            // Initialize Key Vault client
            const keyVaultUri = `https://${keyVaultName}.vault.azure.net`;
            const secretClient = new SecretClient(keyVaultUri, credential);

            // Get secrets
            const secrets = await this._getSecrets(secretClient);

            // Build configuration
            this.config = this._buildConfig(settings, secrets);
            this.initialized = true;
        } catch (error) {
            console.error("Failed to initialize configuration:", error);
            throw error;
        }
    }

    /**
     * Get all settings from App Configuration
     * @param {AppConfigurationClient} client - App Configuration client
     * @returns {Promise<Object>} Settings object
     * @private
     */
    async _getAllSettings(client) {
        const settings = {};
        const settingsIterator = client.listConfigurationSettings();
        
        for await (const setting of settingsIterator) {
            settings[setting.key] = setting.value;
        }
        
        return settings;
    }

    /**
     * Create appropriate credential based on environment
     * @param {string} managedIdentityId - Managed identity client ID
     * @returns {ManagedIdentityCredential|AzureCliCredential} Credential instance
     * @private
     */
    _createCredential(managedIdentityId) {
        const isDevelopment = process.env.NODE_ENV === 'development';
        return isDevelopment ? 
            new AzureCliCredential() : 
            new ManagedIdentityCredential(managedIdentityId);
    }

    /**
     * Get secrets from Key Vault
     * @param {SecretClient} client - Key Vault secret client
     * @returns {Promise<Object>} Secrets object
     * @private
     */
    async _getSecrets(client) {
        const secretNames = [
            'SECRET-AZURE-OPENAI-API-KEY',
            'SECRET-AZURE-SEARCH-KEY',
            'SECRET-BOT-PASSWORD',
            'SECRET-JIRA-API-TOKEN'
        ];

        const secrets = {};
        for (const secretName of secretNames) {
            try {
                const secret = await client.getSecret(secretName);
                secrets[secretName] = this._decryptIfNeeded(secret.value);
            } catch (error) {
                console.error(`Error retrieving secret ${secretName}:`, error.message);
                throw error;
            }
        }

        return secrets;
    }

    /**
     * Build configuration object from settings and secrets
     * @param {Object} settings - App Configuration settings
     * @param {Object} secrets - Key Vault secrets
     * @returns {Object} Complete configuration object
     * @private
     */
    _buildConfig(settings, secrets) {
        return {
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
    }

    /**
     * Decrypt values if they are encrypted
     * @param {string} value - Value to decrypt
     * @returns {string} Decrypted value
     * @private
     */
    _decryptIfNeeded(value) {
        if (!value) return value;
        return value.startsWith('crypto_') ? value.substring(7) : value;
    }

    /**
     * Get configuration
     * @returns {Promise<Object>} Configuration object
     */
    async getConfig() {
        if (!this.initialized) {
            await this.initialize();
        }
        return this.config;
    }
}

// Singleton instance
const configurationService = new ConfigurationService();

/**
 * Get required configuration
 * @returns {Promise<Object>} Configuration object
 */
async function getRequiredConfig() {
    return await configurationService.getConfig();
}

module.exports = {
    getRequiredConfig,
    configurationService
};