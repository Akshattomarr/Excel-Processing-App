import React, { useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

const ExcelProcessor = () => {
  const [excelData, setExcelData] = useState([]);

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = (e) => {
      const binaryStr = e.target.result;
      const workbook = XLSX.read(binaryStr, { type: "binary" });

    
      const sheetName = workbook.SheetNames[0];
      const worksheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

      setExcelData(worksheet);
    };

    reader.readAsBinaryString(file);
  };

  
  const handleDownload = () => {
    if (!excelData.length) {
      alert("Please upload an Excel sheet first.");
      return;
    }

    
    const processedData = excelData.map((row) => ({
      "Employee & Activity Code": `${row.Employee} - ${row["Activity Code"]}`,
      Time: row.Time,
      Comments: row.Comments,
    }));

    
    const sortedData = processedData.sort((a, b) => {
      const priorityCodes = ["123", "456"];
      const aCode = a["Employee & Activity Code"].split(" - ")[1];
      const bCode = b["Employee & Activity Code"].split(" - ")[1];

      if (priorityCodes.includes(aCode) && !priorityCodes.includes(bCode)) {
        return -1;
      }
      if (!priorityCodes.includes(aCode) && priorityCodes.includes(bCode)) {
        return 1;
      }
      return 0;
    });

 
    const worksheet = XLSX.utils.json_to_sheet(sortedData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Processed Data");

  
    const excelBinary = XLSX.write(workbook, { bookType: "xlsx", type: "binary" });

  
    const buffer = new ArrayBuffer(excelBinary.length);
    const view = new Uint8Array(buffer);
    for (let i = 0; i < excelBinary.length; i++) {
      view[i] = excelBinary.charCodeAt(i) & 0xff;
    }
    saveAs(new Blob([buffer], { type: "application/octet-stream" }), "Processed_Report.xlsx");
  };

  return (
    <div style={{ padding: "20px" }}>
      <h2>Excel Report Generator</h2>
      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
      <br />
      <button onClick={handleDownload} disabled={!excelData.length}>
        Download Processed Excel
      </button>
    </div>
  );
};

export default ExcelProcessor;
