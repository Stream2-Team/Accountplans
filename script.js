{\rtf1\ansi\ansicpg1252\cocoartf2818
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\paperw11900\paperh16840\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx566\tx1133\tx1700\tx2267\tx2834\tx3401\tx3968\tx4535\tx5102\tx5669\tx6236\tx6803\pardirnatural\partightenfactor0

\f0\fs24 \cf0 // Export to Excel\
function exportToExcel() \{\
  // Create a new workbook\
  const wb = XLSX.utils.book_new();\
\
  // Convert the form to a sheet (we target the entire form element)\
  const ws = XLSX.utils.table_to_sheet(document.querySelector('form'));\
\
  // Append the sheet to the workbook\
  XLSX.utils.book_append_sheet(wb, ws, 'Key Account Plan');\
\
  // Trigger file download\
  XLSX.writeFile(wb, 'Key_Account_Plan.xlsx');\
\}\
\
// Export to Word\
function exportToWord() \{\
  // Get the entire form as HTML content\
  const htmlContent = document.querySelector('form').outerHTML;\
\
  // Convert the HTML content to a Word-compatible format\
  const converted = htmlDocx.asHTML(htmlContent);\
\
  // Create a Blob (binary large object) of the Word document\
  const blob = new Blob([converted], \{ type: 'application/msword' \});\
\
  // Trigger download of the Word document\
  saveAs(blob, 'Key_Account_Plan.docx');\
\}\
}