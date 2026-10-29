try {
    var doc = app.activeDocument;
    if (!doc) {
        alert("No document open. Please open a document and try again.");
        exit();
    }

    // Get active layout (based on selected page or fallback to page 2)
    var activePage = app.activeWindow.activePage || doc.pages[1];
    var activeLayout = activePage.parent;
    while (activeLayout.constructor.name !== "Layout" && activeLayout.constructor.name !== "Document") {
        activeLayout = activeLayout.parent;
    }
    var layoutName = activeLayout.name || "Default Layout";

    // Create text frame on page 2 (doc.pages[1])
    var firstPage = doc.pages[1];
    var tocFrame = firstPage.textFrames.add({
        geometricBounds: [20, 20, 200, 190] // A4 landscape, mm
    });
    tocFrame.contents = "CONTENTS\r";

    // Define sections with sub-sections for DISASSEMBLY and ASSEMBLY
    var sections = [
        { name: "GENERAL INFORMATION", masters: [{ master: "GI-Parent", display: "GENERAL INFORMATION" }] },
        { name: "PREPARATION", masters: [{ master: "PR-Parent", display: "PREPARATION" }] },
        {
            name: "DISASSEMBLY", isHeader: true, masters: [
                { master: "FH-DISASSY", display: "FRONT HEAD" },
                { master: "GH-DISASSY", display: "GEAR HOUSING" },
                { master: "IB-DISASSY", display: "INPUT BODY" }
            ]
        },
        {
            name: "ASSEMBLY", isHeader: true, masters: [
                { master: "FH-ASSY", display: "FRONT HEAD" },
                { master: "GH-ASSY", display: "GEAR HOUSING" },
                { master: "IB-ASSY", display: "INPUT BODY" }
            ]
        },
    ];

    // Debug: Collect master page names and assignments for active layout
    var masterNames = new Array(); // Explicitly declare as Array
    var pageAssignments = new Array(); // Explicitly declare as Array
    var layoutMasters = doc.masterSpreads.everyItem().getElements();
    for (var i = 0; i < layoutMasters.length; i++) {
        masterNames.push(layoutMasters[i].name);
    }

    // Debug: Check masterNames
    alert("Master Names: " + masterNames.join(", ") + "\nType: " + (masterNames instanceof Array ? "Array" : typeof masterNames));

    for (var i = 0; i < activeLayout.pages.length; i++) {
        var page = activeLayout.pages[i];
        var masterName = page.appliedMaster ? page.appliedMaster.name : "None";
        pageAssignments.push("Page " + (page.documentOffset + 1) + ": " + masterName);
    }

    // Validate expected master pages
    var expectedMasters = new Array(); // Explicitly declare as Array
    for (var j = 0; j < sections.length; j++) {
        if (sections[j].masters) {
            for (var k = 0; k < sections[j].masters.length; k++) {
                expectedMasters.push(sections[j].masters[k].master);
            }
        }
    }

    // Debug: Check expectedMasters
    alert("Expected Masters: " + expectedMasters.join(", ") + "\nType: " + (expectedMasters instanceof Array ? "Array" : typeof expectedMasters));

    // Find missing masters using a loop instead of indexOf
    var missingMasters = new Array(); // Explicitly declare as Array
    for (var i = 0; i < expectedMasters.length; i++) {
        var master = expectedMasters[i];
        var found = false;
        for (var m = 0; m < masterNames.length; m++) {
            if (masterNames[m] === master) {
                found = true;
                break;
            }
        }
        if (!found) {
            missingMasters.push(master);
        }
    }
    if (missingMasters.length > 0) {
        alert("Warning: The following master pages are not found in the document: " + missingMasters.join(", ") + "\nCheck master page names in the sections array.");
    }

    alert("Active Layout: " + layoutName + "\nMaster pages: " + masterNames.join(", ") + "\nAssignments:\n" + pageAssignments.join("\n"));

    // Initialize page ranges for each master
    var pageRanges = {};
    for (var i = 0; i < masterNames.length; i++) {
        pageRanges[masterNames[i]] = { start: null, end: null };
    }

    // Scan pages for master page ranges in active layout
    for (var i = 0; i < activeLayout.pages.length; i++) {
        var page = activeLayout.pages[i];
        var masterName = page.appliedMaster ? page.appliedMaster.name : "None";
        if (masterName !== "None") {
            if (pageRanges[masterName].start === null) {
                pageRanges[masterName].start = page.documentOffset + 1;
            }
            pageRanges[masterName].end = page.documentOffset + 1;
        }
    }

    // Generate TOC entries with consistent hyphens
    var entriesAdded = 0;
    for (var j = 0; j < sections.length; j++) {
        var section = sections[j];
        if (section.isHeader) {
            // Add header (e.g., DISASSEMBLY, ASSEMBLY)
            tocFrame.insertionPoints[-1].contents = section.name + "\r";
        }
        if (section.masters && section.masters instanceof Array) {
            for (var k = 0; k < section.masters.length; k++) {
                var subSection = section.masters[k];
                var master = subSection.master;
                if (pageRanges[master] && pageRanges[master].start !== null) {
                    var range = pageRanges[master].start + "-" + pageRanges[master].end; // Standard hyphen
                    var indent = section.isHeader ? "    " : "";
                    tocFrame.insertionPoints[-1].contents = indent + subSection.display + "\t" + range + "\r";
                    entriesAdded++;
                }
            }
        }
    }

    // Apply TOC_Entry style
    var style = doc.paragraphStyles.itemByName("TOC_Entry") || doc.paragraphStyles[0];
    tocFrame.paragraphs.everyItem().appliedParagraphStyle = style;

    // Ensure consistent hyphen formatting
    tocFrame.texts.everyItem().changeGrep({
        findWhat: "-",
        changeTo: "-",
        appliedCharacterStyle: doc.characterStyles[0] // Use default character style for hyphens
    });

    if (entriesAdded === 0) {
        alert("Warning: No TOC entries added in " + layoutName + ". Check master page names and assignments.");
    } else {
        alert("TOC created on page 2 of " + layoutName + " with " + entriesAdded + " entries!");
    }
} catch (e) {
    alert("Error running script: " + e.message + "\nLine: " + e.line);
}