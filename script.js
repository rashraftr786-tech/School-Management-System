// Function to verify everything is linked
console.log("Result Management System: Script loaded.");

// 1. Wrap in DOMContentLoaded to ensure HTML elements exist
document.addEventListener('DOMContentLoaded', () => {
    console.log("DOM fully loaded.");
});

// 2. The function called by the "Process File" button in your HTML
function processExcel() {
    const fileInput = document.getElementById('excelFile'); // Corrected ID
    const file = fileInput.files[0];

    if (!file) {
        alert("Please select an Excel file first.");
        return;
    }

    console.log("Processing file: ", file.name);

    const reader = new FileReader();
    reader.onload = function(evt) {
        try {
            const data = evt.target.result;
            // Using 'binary' as in your original logic
            const workbook = XLSX.read(data, { type: 'binary' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            
            // Convert Excel to JSON data
            const jsonData = XLSX.utils.sheet_to_json(firstSheet);
            
            if (jsonData.length > 0) {
                renderTable(jsonData); // Call the function to display data
                alert("Upload successful! Found " + jsonData.length + " students.");
            } else {
                alert("The Excel file seems to be empty.");
            }
        } catch (err) {
            console.error("Error reading file:", err);
            alert("Could not read Excel file. Check console for details.");
        }
    };
    reader.readAsBinaryString(file);
}

// 3. New function to actually show the data in your HTML table
function renderTable(data) {
    const tableBody = document.getElementById('tableBody');
    tableBody.innerHTML = ''; // Clear current rows

    data.forEach((student, index) => {
        const row = `
            <tr>
                <td><img src="https://via.placeholder.com/40" alt="Student"></td>
                <td>${student['Student Name'] || 'N/A'}</td>
                <td>${student["Father's Name"] || 'N/A'}</td>
                <td>${student['Roll No'] || 'N/A'}</td>
                <td>${student['Class'] || 'N/A'}</td>
                <td>${student['Total (%)'] || '0'}%</td>
                <td>${student['Grade'] || 'N/A'}</td>
                <td>${student['Rank'] || 'N/A'}</td>
                <td>
                    <button onclick="openPhotoModal(${index})" class="btn-sm">Photo</button>
                    <button onclick="generateReport(${index})" class="btn-sm">PDF</button>
                </td>
            </tr>
        `;
        tableBody.innerHTML += row;
    });
}

// Placeholder for filtering logic (called by your HTML inputs)
function filterTable() {
    console.log("Filtering table...");
    // Logic for search and class filter goes here
}
