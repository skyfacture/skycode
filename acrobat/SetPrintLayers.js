/* Acrobat JavaScript: Set layer print flags by layer name */

var nonPrintLayers = ["Notes", "Markup", "Guides"];     // layers you want NOT to print
var printLayers = ["Main", "Graphics", "Text"];      // layers you want to print

var ocgs = this.getOCGs();

if (ocgs) {
    for (var i = 0; i < ocgs.length; i++) {
        var layer = ocgs[i];
        var name = layer.name;

        // Set non‑printing layers
        if (nonPrintLayers.indexOf(name) !== -1) {
            layer.printState = false;
        }

        // Set printing layers
        if (printLayers.indexOf(name) !== -1) {
            layer.printState = true;
        }
    }
}

app.alert("Layer print states updated.");
