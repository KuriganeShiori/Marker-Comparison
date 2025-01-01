const { contextBridge, ipcRenderer } = require('electron');
const path = require('path');
const fs = require('fs');

// Expose protected methods that allow the renderer process to use
// the ipcRenderer without exposing the entire object
contextBridge.exposeInMainWorld('electronAPI', {
    showOpenDialog: (options) => ipcRenderer.invoke('show-open-dialog', options),
    openExternal: (url) => ipcRenderer.invoke('open-external', url),
    path: {
        join: (...args) => path.join(...args),
        resolve: (...args) => path.resolve(...args),
        __dirname: __dirname
    },
    fs: {
        readFileSync: (path) => fs.readFileSync(path),
        existsSync: (path) => fs.existsSync(path),
        readdirSync: (path) => fs.readdirSync(path)
    },
    require: (modulePath) => {
        try {
            // Use absolute path resolution
            const resolvedPath = path.resolve(__dirname, modulePath);
            console.log('Loading module from:', resolvedPath);
            
            // Require the module
            const module = require(resolvedPath);
            
            // Log the module type for debugging
            console.log('Module type:', typeof module);
            console.log('Module content:', module);
            
            return module;
        } catch (error) {
            console.error(`Error requiring module ${modulePath}:`, error);
            return null;
        }
    }
}); 