// Author: Kevin Paulose
// Last updated: March 20, 2025

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index');
}

/**
 * Processes multiple slide links in sequence.
 * Calls exportSlidesToImages() for each link.
 */
function exportAllSlidesToImages(links) {
  var logs = [];
  var allExportedUrls = [];

  // Process each link one-by-one
  for (var i = 0; i < links.length; i++) {
    var link = links[i];
    logs.push("Starting export for Link #" + (i+1) + ": " + link);
    sendProgress(logs);

    // Call the single-link function
    var result = exportSlidesToImages(link);

    // Gather up the exported URLs
    allExportedUrls = allExportedUrls.concat(result.urls);

    logs.push("Finished export for Link #" + (i+1));
    sendProgress(logs);
  }

  logs.push("✅ All slide links processed!");
  sendProgress(logs);

  // Return everything so the front-end can show them
  return { allUrls: allExportedUrls, logs: logs };
}

/**
 * Exports all slides from a single Google Slides link to PNG images.
 * Overwrites (trashes) any existing PNGs of the same name in the folder.
 */
function exportSlidesToImages(presentationUrl) {
  try {
    var presentationId = presentationUrl.match(/[-\w]{25,}/)[0];
    if (!presentationId) throw new Error('Invalid Google Slides URL.');

    var presentation = SlidesApp.openById(presentationId);
    var slides = presentation.getSlides();
    var presentationFile = DriveApp.getFileById(presentation.getId());
    var parentFolder = presentationFile.getParents().next();

    var presentationFilename = presentationFile.getName();
    // Check if the name is being truncated due to numeric prefix
    if (/^\d+\.\d+/.test(presentationFilename)) {
      var driveMetadata = DriveApp.getFileById(presentation.getId());
      presentationFilename = driveMetadata.getName(); // Re-fetch the full name
    }

    // Check if 'PNG Slide Exports' folder exists, create if it doesn't
    var imagesFolderIterator = parentFolder.getFoldersByName('PNG Slide Exports');
    var imagesFolder = imagesFolderIterator.hasNext() 
      ? imagesFolderIterator.next() 
      : parentFolder.createFolder('PNG Slide Exports');

    // Check if subfolder exists for this presentation, otherwise create it
    var presentationFolderIterator = imagesFolder.getFoldersByName(presentationFilename);
    var presentationFolder = presentationFolderIterator.hasNext()
      ? presentationFolderIterator.next()
      : imagesFolder.createFolder(presentationFilename);

    var exportedSlides = [];
    var logs = [];

    for (var i = 0; i < slides.length; i++) {
      var slide = slides[i];
      var slideId = slide.getObjectId();
      var slideNumber = (i + 1).toString().padStart(2, '0');

      logs.push(`Processing Slide ${slideNumber}...`);
      sendProgress(logs);

      var success = false;
      var retries = 0;
      var maxRetries = 5;
      var retryDelay = 2500;

      while (!success && retries < maxRetries) {
        retries++;
        try {
          var url = `https://docs.google.com/presentation/d/${presentation.getId()}/export/png?id=${presentation.getId()}&pageid=${slideId}`;
          var options = {
            headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
            muteHttpExceptions: true
          };
          var response = UrlFetchApp.fetch(url, options);

          if (response.getResponseCode() === 200) {
            // Check if a file with the same name already exists. If so, trash it.
            var fileName = presentationFilename + '_Slide' + slideNumber + '.png';
            var existingFiles = presentationFolder.getFilesByName(fileName);
            while (existingFiles.hasNext()) {
              var existingFile = existingFiles.next();
              existingFile.setTrashed(true);
            }

            // Create the new PNG file
            var imageBlob = response.getBlob().setName(fileName);
            var file = presentationFolder.createFile(imageBlob);
            exportedSlides.push(file.getUrl());
            logs.push(`✅ Slide ${slideNumber} exported successfully.`);
            sendProgress(logs);
            success = true;
          } else if (response.getResponseCode() === 429) {
            logs.push(`⚠️ Rate limit hit for Slide ${slideNumber}. Retrying... (${retries}/${maxRetries})`);
            sendProgress(logs);
            Utilities.sleep(retryDelay);
          } else {
            throw new Error(`❌ Failed to export Slide ${slideNumber}. Response Code: ${response.getResponseCode()}`);
          }
        } catch (error) {
          logs.push(error.message);
          sendProgress(logs);
          if (retries >= maxRetries) {
            logs.push(`🚨 Max retries reached for Slide ${slideNumber}. Skipping.`);
            sendProgress(logs);
          } else {
            logs.push(`🔄 Retrying Slide ${slideNumber}...`);
            sendProgress(logs);
            Utilities.sleep(retryDelay);
          }
        }
      }

      // Pause briefly after each successful slide
      if (success) {
        Utilities.sleep(retryDelay);
      }
    }

    logs.push("✅ All slides processed!");
    sendProgress(logs);
    return { urls: exportedSlides, logs: logs };
  } catch (error) {
    return { urls: [], logs: [`❌ Error: ${error.message}`] };
  }
}

/**
 * Sends progress updates to the frontend
 */
function sendProgress(logs) {
  CacheService.getScriptCache().put("exportProgress", JSON.stringify(logs), 300);
}

/**
 * Fetches the latest progress logs
 */
function getProgress() {
  var logs = CacheService.getScriptCache().get("exportProgress");
  return logs ? JSON.parse(logs) : [];
}
