const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

// Deployment configuration
const config = {
  tempDir: '.deployment',
  includeDirs: ['src', 'host.json', 'index.js', 'package.json'],
  excludePatterns: ['.map', '.ts', '.test.js', '.spec.js']
};

function cleanup() {
  if (fs.existsSync(config.tempDir)) {
    fs.rmSync(config.tempDir, { recursive: true });
  }
}

function prepareDeploymentPackage() {
  // Create temp directory
  cleanup();
  fs.mkdirSync(config.tempDir);

  // Copy required files
  config.includeDirs.forEach(item => {
    const source = path.join(__dirname, item);
    const dest = path.join(config.tempDir, item);
    
    if (fs.existsSync(source)) {
      if (fs.lstatSync(source).isDirectory()) {
        fs.cpSync(source, dest, { recursive: true });
      } else {
        fs.copyFileSync(source, dest);
      }
    }
  });

  // Install production dependencies
  execSync('npm install --production', { 
    cwd: config.tempDir,
    stdio: 'inherit'
  });

  // Remove any excluded file patterns
  function removeExcludedFiles(dir) {
    const files = fs.readdirSync(dir);
    files.forEach(file => {
      const fullPath = path.join(dir, file);
      if (fs.lstatSync(fullPath).isDirectory()) {
        removeExcludedFiles(fullPath);
      } else {
        if (config.excludePatterns.some(pattern => file.endsWith(pattern))) {
          fs.unlinkSync(fullPath);
        }
      }
    });
  }
  
  removeExcludedFiles(config.tempDir);

  // Deploy using Azure Functions Core Tools
  execSync('func azure functionapp publish zypt-sharepointindexer --javascript', { 
    cwd: config.tempDir,
    stdio: 'inherit'
  });
}

try {
  prepareDeploymentPackage();
  cleanup();
  console.log('Deployment completed successfully');
} catch (error) {
  console.error('Deployment failed:', error);
  cleanup();
  process.exit(1);
}