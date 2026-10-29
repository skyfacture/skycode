try {
    var doc = app.activeDocument;
    if (!doc) {
        alert("No document open. Please open a document and try again.");
        exit();
    }
    
    // Create text frame on first page
    var firstPage = doc.pages[0];
    var tocFrame = firstPage.textFrames.add({
        geometricBounds: [20, 20, 200, 190] // Adjusted for A4 landscape (mm)
    });
    tocFrame.contents = "CONTENTS\r";
    
    // Get master spreads and initialize ranges
    var masters = doc.masterSpreads.everyItem().name;
    var pageRanges = {};
    for (var i = 0; i < masters.length; i++) {
        pageRanges[masters[i]] = {start: null, end: null};
    }
    
    // Scan pages for master page ranges
    for (var i = 0; i < doc.pages.length; i++) {
        var page = doc.pages[i];
        var masterName = page.appliedMaster ? page.appliedMaster.name : "None";
        if (masterName !== "None") {
            if (pageRanges[masterName].start === null) {
                pageRanges[masterName].start = page.documentOffset + 1;
            }
            pageRanges[masterName].end = page.documentOffset + 1;
        }
    }
    
    // Generate TOC entries
    for (var master in pageRanges) {
        if (pageRanges[master].start !== null) {
            var range = pageRanges[master].start + "–" + pageRanges[master].end;
            tocFrame.contents += master.toUpperCase() + " .... " + range + "\r";
        }
    }
    
    // Apply paragraph style (fallback to default if TOC_Entry missing)
    var style = doc.paragraphStyles.itemByName("TOC_Entry") || doc.paragraphStyles[0];
    tocFrame.paragraphs.everyItem().appliedParagraphStyle = style;
    
    alert("TOC created on page 1!");
} catch (e) {
    alert("Error running script: " + e.message);
}