// Assuming you've included the SheetJS library in your HTML

// Event listener for the "Generate Excel" button
document.getElementById("generateExcel").addEventListener("click", function() {
    // Create a new Excel workbook
    var workbook = XLSX.utils.book_new();
    
    // Create a worksheet
    var worksheet = XLSX.utils.json_to_sheet([
        // Sample data, replace with your actual data
        { Name: "John Doe", "2023-08-27": { "Time In": "09:00 AM", "Time Out": "05:00 PM" }},
        { Name: "Jane Smith", "2023-08-27": { "Time In": "08:30 AM", "Time Out": "04:30 PM" }}
        // Add more data as needed
    ]);
    
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, "Attendance");
    
    // Generate a blob from the workbook
    var blob = XLSX.write(workbook, { bookType: 'xlsx', type: 'blob' });

    // Create a download link
    var link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "attendance.xlsx";
    link.click();
});
