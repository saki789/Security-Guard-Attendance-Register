// Load the SheetJS library
var script = document.createElement('script');
script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js';
script.onload = function () {
    // Event listener for the "Generate Excel" button
    document.getElementById("generateExcel").addEventListener("click", function () {
        // Create a new Excel workbook
        var workbook = XLSX.utils.book_new();

        // Create a worksheet data array
        var data = [
            // Sample data, replace with your actual data
            ["Name", "2023-08-27", "", "2023-08-28", ""], // Header row
            ["", "Time In", "Time Out", "Time In", "Time Out"], // Sub-header row
            ["John Doe", "09:00 AM", "05:00 PM", "", ""], // Data row
            ["Jane Smith", "08:30 AM", "04:30 PM", "", ""] // Data row
            // Add more data rows as needed
        ];

        // Create a worksheet
        var worksheet = XLSX.utils.aoa_to_sheet(data);

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
};
document.head.appendChild(script);
