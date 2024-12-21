// test-keyvault.js
const { DefaultAzureCredential } = require("@azure/identity");
const { SecretClient } = require("@azure/keyvault-secrets");
const { AppConfigurationClient } = require("@azure/app-configuration");

async function getAppConfig() {
    console.log("\nRetrieving App Configuration...");
    const connectionString = "Endpoint=https://zyptai-dev-app-config.azconfig.io;Id=3b6Y;Secret=jQtx75kcpqsoAisGd4E8VHaaXFU5SODG8jAR4oQ2fg6LUtzCfALZJQQJ99AKACYeBjFWafWYAAACAZAC4Ncv";
    
    try {
        console.log("Creating AppConfigurationClient...");
        const appConfigClient = new AppConfigurationClient(connectionString);
        
        const settings = {};
        console.log("Reading settings...");
        
        const settingsIterator = appConfigClient.listConfigurationSettings();
        for await (const setting of settingsIterator) {
            settings[setting.key] = setting.value;
            console.log(`Retrieved setting ${setting.key}: ${setting.value}`);
        }
        
        return settings;
    } catch (error) {
        console.log("Error retrieving App Configuration:");
        console.error(error);
        throw error;
    }
}

async function testKeyVaultAccess() {
    console.log("\nStarting Key Vault access test...");

    try {
        // Get App Config first
        const appConfig = await getAppConfig();

        // Create DefaultAzureCredential
        console.log("\nCreating DefaultAzureCredential...");
        const credential = new DefaultAzureCredential();
        
        // Key Vault setup using values from App Config
        const keyVaultName = "zyptaidevkeyvault";
        const keyVaultUrl = `https://${keyVaultName}.vault.azure.net`;
        console.log(`Using Key Vault URL: ${keyVaultUrl}`);

        // Create Secret Client
        console.log("Creating SecretClient...");
        const secretClient = new SecretClient(keyVaultUrl, credential);

        // Try to get secrets
        const secretsToTest = [
            'SECRET-AZURE-OPENAI-API-KEY',
            'SECRET-AZURE-SEARCH-KEY',
            'SECRET-BOT-PASSWORD',
            'SECRET-JIRA-API-TOKEN'
        ];

        console.log("\nAttempting to retrieve secrets...");
        for (const secretName of secretsToTest) {
            try {
                console.log(`\nTrying to get secret: ${secretName}`);
                const secret = await secretClient.getSecret(secretName);
                console.log("Success! Secret retrieved:");
                console.log({
                    name: secret.name,
                    value: secret.value ? "present" : "missing",
                    properties: secret.properties
                });
            } catch (error) {
                console.log(`Error retrieving secret ${secretName}:`);
                console.log("Error name:", error.name);
                console.log("Error message:", error.message);
            }
        }

    } catch (error) {
        console.log("\nError occurred!");
        console.log("Error name:", error.name);
        console.log("Error message:", error.message);
        
        if (error.request) {
            console.log("\nRequest details:");
            console.log("URL:", error.request.url);
            console.log("Method:", error.request.method);
            console.log("Headers:", JSON.stringify(error.request.headers, null, 2));
        }

        if (error.details) {
            console.log("\nError details:");
            console.log(JSON.stringify(error.details, null, 2));
        }
    }
}

// Log environment before starting
console.log("Environment variables:");
console.log("AZURE_TENANT_ID:", process.env.AZURE_TENANT_ID);
console.log("AZURE_CLIENT_ID:", process.env.AZURE_CLIENT_ID);
console.log("(AZURE_CLIENT_SECRET presence:", !!process.env.AZURE_CLIENT_SECRET, ")");

// Run the test
testKeyVaultAccess().then(() => {
    console.log("\nTest complete");
}).catch(error => {
    console.log("\nTest failed with error:", error);
});