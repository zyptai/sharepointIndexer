// Copyright (c) 2024 ZyptAI, tim.barrow@zyptai.com
// Proprietary and confidential to ZyptAI
// File: utils/configService.js
// Purpose: Centralized configuration service for App Configuration access.
//          Provides consistent configuration across the application.

const { AppConfigurationClient } = require("@azure/app-configuration");

/**
 * Configuration service class to handle App Configuration operations
 */
class ConfigurationService {
    constructor() {
        this.initialized = false;
        this.settings = null;
        this.appConfigClient = null;
    }

    /**
     * Initialize App Configuration client
     * @private
     */
    async _initializeAppConfig() {
        console.log("Initializing App Configuration client");
        if (!process.env.AZURE_APP_CONFIG_CONNECTION_STRING) {
            throw new Error("AZURE_APP_CONFIG_CONNECTION_STRING not found in environment variables");
        }

        this.appConfigClient = new AppConfigurationClient(process.env.AZURE_APP_CONFIG_CONNECTION_STRING);
        console.log("App Configuration client initialized");
    }

    /**
     * Get a specific setting from App Configuration
     * @param {string} settingName - Name of the setting to retrieve
     * @returns {Promise<string>} The setting value
     */
    async getSetting(settingName) {
        try {
            if (!this.appConfigClient) {
                await this._initializeAppConfig();
            }

            const setting = await this.appConfigClient.getConfigurationSetting({ key: settingName });
            return setting.value;
        } catch (error) {
            console.error(`Failed to retrieve setting ${settingName}:`, error.message);
            throw error;
        }
    }

    /**
     * Get all settings from App Configuration
     * @returns {Promise<Object>} Object containing all settings
     */
    async getAllSettings() {
        try {
            if (!this.appConfigClient) {
                await this._initializeAppConfig();
            }

            if (this.settings) {
                return this.settings;
            }

            console.log("Fetching all settings from App Configuration");
            const settings = {};
            const settingsIterator = this.appConfigClient.listConfigurationSettings();

            for await (const setting of settingsIterator) {
                settings[setting.key] = setting.value;
            }

            this.settings = settings;
            return settings;
        } catch (error) {
            console.error("Failed to retrieve settings:", error.message);
            throw error;
        }
    }
}

// Singleton instance
const configurationService = new ConfigurationService();
module.exports = configurationService;