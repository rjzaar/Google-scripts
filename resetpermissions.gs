// Resumable version - handles Google Apps Script time limits
// Note that this script will remove all permissions of all files and subfolders
// (including files in subfolders) of the given folder

// https://drive.google.com/drive/folders/abcdefgh
const FOLDER_ID = "abcdefgh";
const MAX_EXECUTION_TIME = 5 * 60 * 1000; // 5 minutes in milliseconds (leaving 1 min buffer)
const RESUME_DELAY = 1; // minutes to wait before resuming

// Main entry point - call this to start or resume the process
function start() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const startTime = Date.now();
  
  // Initialize if this is a fresh start
  if (!scriptProperties.getProperty('processing')) {
    Logger.log("Starting fresh permission reset process...");
    initializeProcess();
  } else {
    Logger.log("Resuming permission reset process...");
  }
  
  // Process items until we run out of time
  processItems(startTime);
  
  // Check if we're done
  const queueStr = scriptProperties.getProperty('folderQueue');
  const fileQueueStr = scriptProperties.getProperty('fileQueue');
  
  if (!queueStr && !fileQueueStr) {
    Logger.log("Process complete! All permissions have been reset.");
    cleanup();
  } else {
    Logger.log("Time limit approaching. Creating trigger to resume...");
    createResumeTrigger();
  }
}

// Initialize the processing state
function initializeProcess() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const folder = DriveApp.getFolderById(FOLDER_ID);
  
  // Store the root folder ID and mark as processing
  scriptProperties.setProperty('processing', 'true');
  scriptProperties.setProperty('rootFolderId', FOLDER_ID);
  
  // Initialize folder queue with root folder
  const folderQueue = [FOLDER_ID];
  scriptProperties.setProperty('folderQueue', JSON.stringify(folderQueue));
  scriptProperties.setProperty('fileQueue', JSON.stringify([]));
  scriptProperties.setProperty('processedFolders', JSON.stringify({}));
  scriptProperties.setProperty('processedFiles', JSON.stringify({}));
  
  Logger.log("Initialized process for folder: " + folder.getName());
}

// Process items until time limit is reached
function processItems(startTime) {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  while (Date.now() - startTime < MAX_EXECUTION_TIME) {
    // First, process any files in the queue
    if (processNextFile(startTime)) {
      continue;
    }
    
    // Then process folders
    if (!processNextFolder(startTime)) {
      // Nothing left to process
      break;
    }
  }
}

// Process the next folder in the queue
function processNextFolder(startTime) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const queueStr = scriptProperties.getProperty('folderQueue');
  
  if (!queueStr) return false;
  
  const folderQueue = JSON.parse(queueStr);
  if (folderQueue.length === 0) return false;
  
  // Check if we're running out of time before starting a new folder
  if (Date.now() - startTime >= MAX_EXECUTION_TIME) {
    return false;
  }
  
  const folderId = folderQueue.shift();
  const processedFolders = JSON.parse(scriptProperties.getProperty('processedFolders') || '{}');
  
  // Skip if already processed
  if (processedFolders[folderId]) {
    scriptProperties.setProperty('folderQueue', JSON.stringify(folderQueue));
    return true;
  }
  
  try {
    const folder = DriveApp.getFolderById(folderId);
    Logger.log("Processing folder: " + folder.getName() + " (ID: " + folderId + ")");
    
    // Get all files in this folder and add to file queue
    const fileQueue = JSON.parse(scriptProperties.getProperty('fileQueue') || '[]');
    const files = folder.getFiles();
    let fileCount = 0;
    
    while (files.hasNext()) {
      const file = files.next();
      fileQueue.push(file.getId());
      fileCount++;
    }
    
    Logger.log("Found " + fileCount + " files in " + folder.getName());
    scriptProperties.setProperty('fileQueue', JSON.stringify(fileQueue));
    
    // Get all subfolders and add to queue
    const childFolders = folder.getFolders();
    let folderCount = 0;
    
    while (childFolders.hasNext()) {
      const childFolder = childFolders.next();
      folderQueue.push(childFolder.getId());
      folderCount++;
    }
    
    Logger.log("Found " + folderCount + " subfolders in " + folder.getName());
    
    // Reset permissions on the folder itself
    resetPermissions(folder);
    
    // Mark folder as processed
    processedFolders[folderId] = true;
    scriptProperties.setProperty('processedFolders', JSON.stringify(processedFolders));
    
  } catch (e) {
    Logger.log("Error processing folder " + folderId + ": " + e.toString());
    // Mark as processed to avoid infinite loop
    processedFolders[folderId] = true;
    scriptProperties.setProperty('processedFolders', JSON.stringify(processedFolders));
  }
  
  // Save updated queue
  scriptProperties.setProperty('folderQueue', JSON.stringify(folderQueue));
  return true;
}

// Process the next file in the queue
function processNextFile(startTime) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const queueStr = scriptProperties.getProperty('fileQueue');
  
  if (!queueStr) return false;
  
  const fileQueue = JSON.parse(queueStr);
  if (fileQueue.length === 0) return false;
  
  // Check if we're running out of time
  if (Date.now() - startTime >= MAX_EXECUTION_TIME) {
    return false;
  }
  
  const fileId = fileQueue.shift();
  const processedFiles = JSON.parse(scriptProperties.getProperty('processedFiles') || '{}');
  
  // Skip if already processed
  if (processedFiles[fileId]) {
    scriptProperties.setProperty('fileQueue', JSON.stringify(fileQueue));
    return true;
  }
  
  try {
    const file = DriveApp.getFileById(fileId);
    Logger.log("Processing file: " + file.getName() + " (ID: " + fileId + ")");
    
    resetPermissions(file);
    
    // Mark file as processed
    processedFiles[fileId] = true;
    scriptProperties.setProperty('processedFiles', JSON.stringify(processedFiles));
    
  } catch (e) {
    Logger.log("Error processing file " + fileId + ": " + e.toString());
    // Mark as processed to avoid infinite loop
    processedFiles[fileId] = true;
    scriptProperties.setProperty('processedFiles', JSON.stringify(processedFiles));
  }
  
  // Save updated queue
  scriptProperties.setProperty('fileQueue', JSON.stringify(fileQueue));
  return true;
}

// Reset permissions on a file or folder
function resetPermissions(asset) {
  try {
    if (asset) {
      removeSharing(asset);
      removeEditPermissions(asset);
      removeViewPermissions(asset);
    } else {
      Logger.log("Asset not found! -error");
    }
  } catch (e) {
    Logger.log("Error resetting permissions: " + e.toString());
  }
}

// Set sharing options for the asset (file or folder)
function removeSharing(asset) {
  try {
    asset.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.NONE);
    asset.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.NONE);
  } catch (e) {
    Logger.log("Error removing sharing: " + e.toString());
  }
}

// Remove all users who have edit permissions
function removeEditPermissions(asset) {
  try {
    const editors = asset.getEditors();
    if (editors.length > 0) {
      editors.forEach(editor => {
        const email = editor.getEmail();
        if (email != "") {
          asset.removeEditor(email);
        }
      });
      Logger.log("Editors removed from " + asset.getName());
    }
  } catch (e) {
    Logger.log("Error removing editors: " + e.toString());
  }
}

// Remove all users who have view permissions
function removeViewPermissions(asset) {
  try {
    const viewers = asset.getViewers();
    if (viewers.length > 0) {
      viewers.forEach(viewer => {
        const email = viewer.getEmail();
        if (email != "") {
          asset.removeViewer(email);
        }
      });
      Logger.log("Viewers removed from " + asset.getName());
    }
  } catch (e) {
    Logger.log("Error removing viewers: " + e.toString());
  }
}

// Create a time-based trigger to resume the process
function createResumeTrigger() {
  // Delete any existing resume triggers
  deleteResumeTriggers();
  
  // Create a new trigger to run after the specified delay
  ScriptApp.newTrigger('start')
    .timeBased()
    .after(RESUME_DELAY * 60 * 1000)
    .create();
  
  Logger.log("Created trigger to resume in " + RESUME_DELAY + " minute(s)");
}

// Delete all resume triggers
function deleteResumeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'start') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// Clean up after completion
function cleanup() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('processing');
  scriptProperties.deleteProperty('folderQueue');
  scriptProperties.deleteProperty('fileQueue');
  scriptProperties.deleteProperty('processedFolders');
  scriptProperties.deleteProperty('processedFiles');
  scriptProperties.deleteProperty('rootFolderId');
  
  // Delete any remaining triggers
  deleteResumeTriggers();
  
  Logger.log("Cleanup complete!");
}

// Manual reset function - call this if you need to start over
function resetProcess() {
  cleanup();
  Logger.log("Process reset. Call start() to begin again.");
}

// Get current progress status
function getStatus() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  if (!scriptProperties.getProperty('processing')) {
    Logger.log("No process currently running.");
    return;
  }
  
  const folderQueue = JSON.parse(scriptProperties.getProperty('folderQueue') || '[]');
  const fileQueue = JSON.parse(scriptProperties.getProperty('fileQueue') || '[]');
  const processedFolders = JSON.parse(scriptProperties.getProperty('processedFolders') || '{}');
  const processedFiles = JSON.parse(scriptProperties.getProperty('processedFiles') || '{}');
  
  Logger.log("=== Current Status ===");
  Logger.log("Folders in queue: " + folderQueue.length);
  Logger.log("Files in queue: " + fileQueue.length);
  Logger.log("Folders processed: " + Object.keys(processedFolders).length);
  Logger.log("Files processed: " + Object.keys(processedFiles).length);
  Logger.log("=====================");
}
