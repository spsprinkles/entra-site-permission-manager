// Returns the Entra icon as an SVG element
export function getEntraIcon(height?, width?, className?) {
    // Set the default values
    if (height === void 0) { height = 32; }
    if (width === void 0) { width = 32; }

    // Get the icon element
    let elDiv = document.createElement("div");
    elDiv.innerHTML = "<svg xmlns='http://www.w3.org/2000/svg' width='32' height='32' viewBox='0 0 18 18'><path d='m17.679,10.013c-.013-.015-3.337-3.818-3.334-3.806-.54-.751-1.484-1.224-2.511-1.224-.028,0-.055.003-.082.004-.856.022-1.553.463-2.085,1.031l3.701,4.182h0s-6.159,3.855-6.159,3.855c0,0-1.168.856-2.495.541.109.068,3.474,2.169,3.474,2.169.499.313,1.146.313,1.645,0,0,0,7.339-4.596,7.489-4.687.826-.501.838-1.512.357-2.066Z'/><path d='m10.113,1.467c-.525-.577-1.598-.683-2.248.043-.189.211-7.325,8.262-7.526,8.489-.568.638-.403,1.615.353,2.081,0,0,1.751,1.097,1.865,1.169.696.436,2.571,1.036,4.454-.096l1.179-.738-3.536-2.213s4.461-5.039,4.467-5.047c1.294-1.416,2.987-1.032,3.613-.735.002.001-2.607-2.937-2.621-2.951Z'/></svg>";
    let icon = elDiv.firstChild as SVGImageElement;
    if (icon) {
        // See if a class name exists
        if (className) {
            // Parse the class names
            let classNames = className.split(' ');
            for (var i = 0; i < classNames.length; i++) {
                // Add the class name
                icon.classList.add(classNames[i]);
            }
        } else {
            icon.classList.add("icon-svg");
        }

        // Set the height/width
        height ? icon.setAttribute("height", (height).toString()) : null;
        width ? icon.setAttribute("width", (width).toString()) : null;

        // Hide the icon as non-interactive content from the accessibility API
        icon.setAttribute("aria-hidden", "true");

        // Update the styling
        icon.style.pointerEvents = "none";

        // Support for IE
        icon.setAttribute("focusable", "false");
    }

    // Return the icon
    return icon;
}