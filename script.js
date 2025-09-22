// Event listener for file input
document.getElementById("fileInput").addEventListener("change", handleFileSelect, false);

function handleFileSelect(event) {
    const file = event.target.files[0];  // Get the uploaded file
    if (file && file.name.endsWith(".docx")) {  // Check if it's a .docx file
        const reader = new FileReader();
        reader.onload = function(e) {
            const arrayBuffer = e.target.result;
            // Use Mammoth.js to extract raw text from the .docx file
            mammoth.extractRawText({ arrayBuffer: arrayBuffer })
                .then(function(result) {
                    const text = result.value;
                    cleanText(text);
                })
                .catch(function(err) {
                    console.log("Error reading file: ", err);
                });
        };
        reader.readAsArrayBuffer(file);
    } else {
        alert("Please upload a valid Word document (.docx).");
    }
}

function cleanText(text) {
    // Clean non-ASCII characters from the text
    const cleanedText = text.replace(/[^\x00-\x7F]/g, '');  // Remove non-ASCII characters
    document.getElementById("output").textContent = cleanedText;  // Display the cleaned text
}
