// Copyright (c) 2024 ZyptAI, tim.barrow@zyptai.com
// Proprietary and confidential to ZyptAI  
// File: utils/keyVaultService.js
// Purpose: Manages Azure Key Vault operations for secure secret retrieval.

const { SecretClient } = require("@azure/keyvault-secrets");
const { DefaultAzureCredential } = require("@azure/identity");
const configService = require('../utils/configService.js');

/**
 * Service class to handle all Key Vault operations
 */
class KeyVaultService {
    constructor() {
        this.secretClient = null;
    }

    /**
     * Initialize Key Vault client
     * @private
     */
    async _initializeClient() {
        try {
            // Get Key Vault name from App Config
            console.log("Getting Key Vault name from App Config...");
            const keyVaultName = await configService.getSetting('KEY_VAULT_NAME');
            console.log("Retrieved Key Vault name:", keyVaultName);

            if (!keyVaultName) {
                throw new Error("KEY_VAULT_NAME not found in App Configuration");
            }

            const keyVaultUri = `https://${keyVaultName}.vault.azure.net`;
            console.log("Initializing Key Vault client with URI:", keyVaultUri);

            // Initialize with DefaultAzureCredential
            const credential = new DefaultAzureCredential({
                excludedCredentials: ['workload'],
                additionallyAllowedTenantIds: ['*']
            });

            this.secretClient = new SecretClient(keyVaultUri, credential);
            console.log("Key Vault client initialized successfully");
        } catch (error) {
            console.error("Failed to initialize Key Vault client:", error);
            throw error;
        }
    }

    /**
     * Extract the original secret value from a potentially transformed response
     * @private
     * @param {string} value The value returned from Key Vault
     * @returns {string} The original secret value
     */
    _extractOriginalValue(value) {
        if (!value || typeof value !== 'string') {
            console.log("Invalid value received:", value);
            return value;
        }

        console.log("Raw value from Key Vault:", value);

        // If value starts with crypto_ prefix
        if (value.startsWith('crypto_')) {
            console.log("Found transformed value with crypto_ prefix");
            
            // Extract just the 32-character hexadecimal key
            const matches = value.match(/[0-9a-f]{32}/i);
            if (matches) {
                console.log("Extracted value:", matches[0]);
                return matches[0];
            } else {
                console.log("No 32-char hex match found in value");
            }
        }

        console.log("Using original value:", value);
        return value;
    }

    /**
     * Get a secret from Key Vault
     * @param {string} secretName - Name of the secret to retrieve
     * @returns {Promise<string>} The secret value
     */
    async getSecret(secretName) {
        try {
            console.log(`Starting secret retrieval for: ${secretName}`);

            if (!this.secretClient) {
                console.log("No existing Key Vault client, initializing...");
                await this._initializeClient();
            }

            console.log(`Getting secret: ${secretName}`);
            const secret = await this.secretClient.getSecret(secretName);
            
            // Extract the original value
            const originalValue = this._extractOriginalValue(secret.value);
            
            // Log success with value length for debugging
            console.log(`Retrieved and processed secret ${secretName} successfully. Value length: ${originalValue?.length || 0}`);
            
            return originalValue;
        } catch (error) {
            console.error(`Failed to retrieve secret ${secretName}:`, error);
            throw error;
        }
    }
}

// Export singleton instance
const keyVaultService = new KeyVaultService();
module.exports = keyVaultService;