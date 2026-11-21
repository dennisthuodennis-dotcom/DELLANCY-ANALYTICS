
/* ==========================================
   DELLANCY ANALYTICS
   Complete JavaScript application
========================================== */


/* ==========================================
   1. GLOBAL VARIABLES
========================================== */

let dataset = [];
let headers = [];
let numericColumnsList = [];
let dataChart = null;
let currentReport = null;


/* ==========================================
   2. GET HTML ELEMENTS
========================================== */

const fileInput = document.getElementById("fileInput");
const fileName = document.getElementById("fileName");
const uploadStatus = document.getElementById("uploadStatus");
const dashboard = document.getElementById("dashboard");

const totalRecords = document.getElementById("totalRecords");
const totalColumns = document.getElementById("totalColumns");
const numericColumns = document.getElementById("numericColumns");
const averageValue = document.getElementById("averageValue");

const columnSelect = document.getElementById("columnSelect");
const sumValue = document.getElementById("sumValue");
const maxValue = document.getElementById("maxValue");
const minValue = document.getElementById("minValue");
const selectedAverage = document.getElementById("selectedAverage");

const dataTable = document.getElementById("dataTable");
const previewInfo = document.getElementById("previewInfo");
const exportButton = document.getElementById("exportButton");

const automaticReportButton =
  document.getElementById("automaticReportButton");

const customReportButton =
  document.getElementById("customReportButton");

const reportMetric = document.getElementById("reportMetric");
const reportTitle = document.getElementById("reportTitle");

const reportOutput = document.getElementById("reportOutput");
const reportHeading = document.getElementById("reportHeading");
const reportTimestamp = document.getElementById("reportTimestamp");
const reportContent = document.getElementById("reportContent");

const downloadPdfButton =
  document.getElementById("downloadPdfButton");


/* ==========================================
   3. GENERAL HELPER FUNCTIONS
========================================== */

function toNumber(value) {

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  if (value === null || value === undefined || value === "") {
    return null;
  }

  const cleaned = String(value)
    .trim()
    .replace(/,/g, "")
    .replace(/[^\d.-]/g, "");

  if (cleaned === "" || cleaned === "-" || cleaned === ".") {
    return null;
  }

  const number = Number(cleaned);

  return Number.isFinite(number) ? number : null;
}


function formatNumber(value) {

  if (!Number.isFinite(value)) {
    return "—";
  }

  return new Intl.NumberFormat("en-US", {
    maximumFractionDigits: 2
  }).format(value);
}


function escapeHTML(value) {

  return String(value ?? "").replace(/[&<>"']/g, function(character) {

    const characters = {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#039;"
    };

    return characters[character];
  });
}


function getNumericValues(columnName) {

  return dataset
    .map(row => toNumber(row[columnName]))
    .filter(value => value !== null);
}


function calculateStatistics(columnName) {

  const values = getNumericValues(columnName);

  if (values.length === 0) {
    return null;
  }

  const sum = values.reduce((total, value) => total + value, 0);

  return {
    count: values.length,
    sum: sum,
    average: sum / values.length,
    minimum: Math.min(...values),
    maximum: Math.max(...values),
    missing: dataset.length - values.length
  };
}


function countMissing(columnName) {

  return dataset.filter(row => {

    const value = row[columnName];

    return value === null ||
      value === undefined ||
      String(value).trim() === "";

  }).length;
}


/* ==========================================
   4. FILE UPLOAD
========================================== */

fileInput.addEventListener("change", function(event) {

  const file = event.target.files[0];

  if (!file) return;

  fileName.textContent = file.name;
  uploadStatus.textContent = "Reading your dataset...";
  uploadStatus.className = "status-message";

  const reader = new FileReader();

  reader.onload = function(loadEvent) {

    try {

      const fileData = new Uint8Array(loadEvent.target.result);

      const workbook = XLSX.read(fileData, {
        type: "array"
      });

      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      const rows = XLSX.utils.sheet_to_json(worksheet, {
        defval: "",
        raw: false
      });

      if (rows.length === 0) {
        throw new Error("The selected file contains no data.");
      }

      const cleanedRows = rows.filter(row =>
        Object.values(row).some(value =>
          String(value).trim() !== ""
        )
      );

      if (cleanedRows.length === 0) {
        throw new Error("No usable records were found.");
      }

      dataset = cleanedRows;
      headers = Object.keys(dataset[0]);

      if (headers.length === 0) {
        throw new Error("No columns were found in the dataset.");
      }

      identifyNumericColumns();
      displayDashboard();
      populateColumnSelectors();
      displayTable();

      reportOutput.classList.add("hidden");
      currentReport = null;
      downloadPdfButton.disabled = true;

      uploadStatus.textContent = "Dataset successfully loaded.";
      uploadStatus.className = "status-message success";

      dashboard.classList.remove("hidden");

      dashboard.scrollIntoView({
        behavior: "smooth",
        block: "start"
      });

    } catch (error) {

      uploadStatus.textContent =
        "Unable to read file: " + error.message;

      uploadStatus.className = "status-message error";

      console.error(error);

    }
  };

  reader.onerror = function() {

    uploadStatus.textContent =
      "There was a problem reading this file.";

    uploadStatus.className = "status-message error";

  };

  reader.readAsArrayBuffer(file);

});


/* ==========================================
   5. IDENTIFY NUMERIC COLUMNS
========================================== */

function identifyNumericColumns() {

  numericColumnsList = [];

  headers.forEach(columnName => {

    const values = dataset
      .map(row => row[columnName])
      .filter(value =>
        value !== null &&
        value !== undefined &&
        String(value).trim() !== ""
      );

    if (values.length === 0) return;

    const numericValues = values
      .map(value => toNumber(value))
      .filter(value => value !== null);

    const numericPercentage =
      numericValues.length / values.length;

    if (numericPercentage >= 0.7) {
      numericColumnsList.push(columnName);
    }

  });

}


/* ==========================================
   6. DISPLAY DASHBOARD OVERVIEW
========================================== */

function displayDashboard() {

  totalRecords.textContent = formatNumber(dataset.length);
  totalColumns.textContent = formatNumber(headers.length);

  numericColumns.textContent =
    formatNumber(numericColumnsList.length);

  const allNumericValues = [];

  numericColumnsList.forEach(columnName => {
    allNumericValues.push(...getNumericValues(columnName));
  });

  if (allNumericValues.length > 0) {

    const total = allNumericValues.reduce(
      (sum, value) => sum + value,
      0
    );

    averageValue.textContent =
      formatNumber(total / allNumericValues.length);

  } else {

    averageValue.textContent = "—";

  }

  document.getElementById("datasetName").textContent =
    "Loaded worksheet: " + headers.length + " columns";

}


/* ==========================================
   7. POPULATE NUMERIC COLUMN SELECTORS
========================================== */

function populateColumnSelectors() {

  columnSelect.innerHTML = "";
  reportMetric.innerHTML = "";

  if (numericColumnsList.length === 0) {

    columnSelect.innerHTML =
      '<option value="">No numeric columns found</option>';

    reportMetric.innerHTML =
      '<option value="">No numeric columns found</option>';

    automaticReportButton.disabled = false;
    customReportButton.disabled = true;
    reportMetric.disabled = true;

    return;
  }

  numericColumnsList.forEach(columnName => {

    const dashboardOption = document.createElement("option");

    dashboardOption.value = columnName;
    dashboardOption.textContent = columnName;

    columnSelect.appendChild(dashboardOption);

    const reportOption = document.createElement("option");

    reportOption.value = columnName;
    reportOption.textContent = columnName;

    reportMetric.appendChild(reportOption);

  });

  columnSelect.disabled = false;
  reportMetric.disabled = false;

  automaticReportButton.disabled = false;
  customReportButton.disabled = false;

  updateMetricAnalysis();

}


/* ==========================================
   8. METRIC ANALYSIS
========================================== */

columnSelect.addEventListener("change", updateMetricAnalysis);


function updateMetricAnalysis() {

  const selectedColumn = columnSelect.value;

  if (!selectedColumn) return;

  const statistics = calculateStatistics(selectedColumn);

  if (!statistics) return;

  sumValue.textContent = formatNumber(statistics.sum);
  maxValue.textContent = formatNumber(statistics.maximum);
  minValue.textContent = formatNumber(statistics.minimum);

  selectedAverage.textContent =
    formatNumber(statistics.average);

  drawChart(selectedColumn);

}


/* ==========================================
   9. CREATE CHART
========================================== */

function drawChart(columnName) {

  const canvas = document.getElementById("dataChart");
  const context = canvas.getContext("2d");

  const validRows = dataset
    .map((row, index) => ({
      label: row[headers[0]] || "Record " + (index + 1),
      value: toNumber(row[columnName])
    }))
    .filter(item => item.value !== null)
    .slice(0, 20);

  if (dataChart) {
    dataChart.destroy();
  }

  dataChart = new Chart(context, {

    type: "bar",

    data: {
      labels: validRows.map(item =>
        String(item.label).substring(0, 25)
      ),

      datasets: [{
        label: columnName,
        data: validRows.map(item => item.value),
        backgroundColor: "rgba(112, 84, 199, 0.75)",
        borderColor: "#7054c7",
        borderWidth: 1,
        borderRadius: 5
      }]
    },

    options: {
      responsive: true,
      maintainAspectRatio: false,

      plugins: {
        legend: {
          display: true
        },

        title: {
          display: true,
          text: columnName + " — First 20 valid records"
        }
      },

      scales: {
        y: {
          beginAtZero: false
        }
      }
    }

  });

}


/* ==========================================
   10. DISPLAY DATA TABLE
========================================== */

function displayTable() {

  const tableHead = dataTable.querySelector("thead");
  const tableBody = dataTable.querySelector("tbody");

  tableHead.innerHTML = "";
  tableBody.innerHTML = "";

  const headerRow = document.createElement("tr");

  headers.forEach(columnName => {

    const heading = document.createElement("th");
    heading.textContent = columnName;

    headerRow.appendChild(heading);

  });

  tableHead.appendChild(headerRow);

  const previewRows = dataset.slice(0, 100);

  previewRows.forEach(row => {

    const tableRow = document.createElement("tr");

    headers.forEach(columnName => {

      const cell = document.createElement("td");
      cell.textContent = row[columnName] ?? "";

      tableRow.appendChild(cell);

    });

    tableBody.appendChild(tableRow);

  });

  previewInfo.textContent =
    "Showing " + previewRows.length +
    " of " + dataset.length + " records.";

}


/* ==========================================
   11. EXPORT CLEANED CSV
========================================== */

exportButton.addEventListener("click", function() {

  if (dataset.length === 0) return;

  const worksheet = XLSX.utils.json_to_sheet(dataset);
  const csv = XLSX.utils.sheet_to_csv(worksheet);

  const blob = new Blob([csv], {
    type: "text/csv;charset=utf-8;"
  });

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = "dellancy_cleaned_data.csv";

  document.body.appendChild(link);
  link.click();
  link.remove();

  URL.revokeObjectURL(url);

});


/* ==========================================
   12. AUTOMATIC REPORT
========================================== */

automaticReportButton.addEventListener("click", function() {

  if (dataset.length === 0) {
    alert("Please upload a dataset first.");
    return;
  }

  const report = {
    type: "automatic",
    title: "Automatic Dataset Report",
    created: new Date(),
    datasetName: fileName.textContent,
    records: dataset.length,
    columns: headers.length,
    numericColumns: numericColumnsList.length,
    statistics: [],
    missingCells: 0,
    chartColumn: numericColumnsList[0] || null
  };

  numericColumnsList.forEach(columnName => {

    const stats = calculateStatistics(columnName);

    if (stats) {
      report.statistics.push({
        name: columnName,
        ...stats
      });
    }

  });

  headers.forEach(columnName => {
    report.missingCells += countMissing(columnName);
  });

  currentReport = report;

  displayAutomaticReport(report);

  reportHeading.textContent = report.title;

  showReportOutput();

});


function displayAutomaticReport(report) {

  let tableRows = "";

  report.statistics.forEach(stat => {

    tableRows += `
      <tr>
        <td>${escapeHTML(stat.name)}</td>
        <td>${formatNumber(stat.count)}</td>
        <td>${formatNumber(stat.average)}</td>
        <td>${formatNumber(stat.minimum)}</td>
        <td>${formatNumber(stat.maximum)}</td>
      </tr>
    `;

  });

  let findings = "";

  if (report.statistics.length > 0) {

    const highestAverage = [...report.statistics].sort(
      (a, b) => b.average - a.average
    )[0];

    findings = `
      <div class="report-content-block">
        <h4>Statistical observation</h4>
        <p>
          Among the numeric columns analysed,
          "${escapeHTML(highestAverage.name)}" has the highest
          calculated average value of
          ${formatNumber(highestAverage.average)}.
          This compares raw averages; different metrics may use
          different units and should not automatically be treated
          as directly comparable.
        </p>
      </div>
    `;

  } else {

    findings = `
      <div class="report-content-block">
        <h4>Statistical observation</h4>
        <p>No numeric columns were identified for statistical analysis.</p>
      </div>
    `;

  }

  reportContent.innerHTML = `

    <div class="report-summary-grid">

      <div class="report-stat">
        <span>Total records</span>
        <strong>${formatNumber(report.records)}</strong>
      </div>

      <div class="report-stat">
        <span>Total columns</span>
        <strong>${formatNumber(report.columns)}</strong>
      </div>

      <div class="report-stat">
        <span>Numeric columns</span>
        <strong>${formatNumber(report.numericColumns)}</strong>
      </div>

      <div class="report-stat">
        <span>Missing cells</span>
        <strong>${formatNumber(report.missingCells)}</strong>
      </div>

      <div class="report-stat">
        <span>Rows analysed</span>
        <strong>${formatNumber(report.records)}</strong>
      </div>

      <div class="report-stat">
        <span>Numeric metrics analysed</span>
        <strong>${formatNumber(report.statistics.length)}</strong>
      </div>

    </div>

    <div class="report-content-block">
      <h4>Dataset description</h4>
      <p>
        This report analyses "${escapeHTML(report.datasetName)}".
        The dataset contains ${formatNumber(report.records)} records
        and ${formatNumber(report.columns)} columns.
        ${formatNumber(report.numericColumns)} columns were identified
        as numeric.
      </p>
    </div>

    <div class="report-content-block">
      <h4>Numeric column statistics</h4>

      <div class="table-wrapper">
        <table class="report-table">
          <thead>
            <tr>
              <th>Column</th>
              <th>Valid values</th>
              <th>Average</th>
              <th>Minimum</th>
              <th>Maximum</th>
            </tr>
          </thead>

          <tbody>
            ${tableRows || `
              <tr>
                <td colspan="5">No numeric statistics available.</td>
              </tr>
            `}
          </tbody>
        </table>
      </div>
    </div>

    ${findings}

    <div class="report-content-block">
      <h4>Data quality</h4>
      <p>
        The dataset contains ${formatNumber(report.missingCells)}
        empty cells across its columns. Missing values should be
        reviewed before drawing business conclusions.
      </p>
    </div>

    <p class="muted-text">
      This report contains calculated statistical summaries.
      It does not independently verify the accuracy of the source data.
    </p>

  `;

}


/* ==========================================
   13. CUSTOM REPORT
========================================== */

customReportButton.addEventListener("click", function() {

  if (dataset.length === 0) {
    alert("Please upload a dataset first.");
    return;
  }

  const selectedMetric = reportMetric.value;

  if (!selectedMetric) {
    alert("Please select a numeric metric.");
    return;
  }

  const statistics = calculateStatistics(selectedMetric);

  if (!statistics) {
    alert("No valid numeric values were found in this column.");
    return;
  }

  const title = reportTitle.value.trim() ||
    selectedMetric + " Report";

  currentReport = {
    type: "custom",
    title: title,
    created: new Date(),
    datasetName: fileName.textContent,
    records: dataset.length,
    metric: selectedMetric,
    statistics: statistics,
    chartColumn: selectedMetric
  };

  displayCustomReport(currentReport);

  reportHeading.textContent = title;

  showReportOutput();

});


function displayCustomReport(report) {

  const stats = report.statistics;

  const missingPercentage = report.records > 0
    ? (stats.missing / report.records) * 100
    : 0;

  reportContent.innerHTML = `

    <div class="report-summary-grid">

      <div class="report-stat">
        <span>Selected metric</span>
        <strong>${escapeHTML(report.metric)}</strong>
      </div>

      <div class="report-stat">
        <span>Valid values</span>
        <strong>${formatNumber(stats.count)}</strong>
      </div>

      <div class="report-stat">
        <span>Missing values</span>
        <strong>${formatNumber(stats.missing)}</strong>
      </div>

      <div class="report-stat">
        <span>Average</span>
        <strong>${formatNumber(stats.average)}</strong>
      </div>

      <div class="report-stat">
        <span>Maximum</span>
        <strong>${formatNumber(stats.maximum)}</strong>
      </div>

      <div class="report-stat">
        <span>Minimum</span>
        <strong>${formatNumber(stats.minimum)}</strong>
      </div>

    </div>

    <div class="report-content-block">
      <h4>Report overview</h4>
      <p>
        This report focuses on the "${escapeHTML(report.metric)}"
        column in the uploaded dataset.
        The dataset contains ${formatNumber(report.records)} records.
        ${formatNumber(stats.count)} valid numeric values were
        included in the calculations.
      </p>
    </div>

    <div class="report-content-block">
      <h4>Performance summary</h4>
      <p>
        The total of the valid values is ${formatNumber(stats.sum)}.
        The average is ${formatNumber(stats.average)}.
        The highest recorded value is ${formatNumber(stats.maximum)},
        while the lowest is ${formatNumber(stats.minimum)}.
      </p>
    </div>

    <div class="report-content-block">
      <h4>Data completeness</h4>
      <p>
        ${formatNumber(stats.missing)} records have no usable
        numeric value for this metric.
        This represents ${formatNumber(missingPercentage)}%
        of the total records.
      </p>
    </div>

    <div class="report-content-block">
      <h4>Interpretation</h4>
      <p>
        The average, minimum and maximum provide a basic view
        of the distribution of this metric. Further analysis
        should consider the business meaning of the column,
        its units and the period covered by the data.
      </p>
    </div>

  `;

}


/* ==========================================
   14. SHOW REPORT
========================================== */

function showReportOutput() {

  reportTimestamp.textContent =
    "Generated on " + currentReport.created.toLocaleString();

  reportOutput.classList.remove("hidden");
  downloadPdfButton.disabled = false;

  if (currentReport.chartColumn) {

    columnSelect.value = currentReport.chartColumn;
    updateMetricAnalysis();

  }

  reportOutput.scrollIntoView({
    behavior: "smooth",
    block: "start"
  });

}


/* ==========================================
   15. DOWNLOAD PDF REPORT
========================================== */

downloadPdfButton.addEventListener("click", function() {

  if (!currentReport) {
    alert("Generate a report before downloading.");
    return;
  }

  if (!window.jspdf || !window.jspdf.jsPDF) {
    alert("The PDF library could not load. Check your internet connection and refresh.");
    return;
  }

  const { jsPDF } = window.jspdf;

  const pdf = new jsPDF({
    orientation: "portrait",
    unit: "mm",
    format: "a4"
  });

  const navy = [20, 39, 70];
  const purple = [112, 84, 199];

  // PDF heading

  pdf.setFillColor(...navy);
  pdf.rect(0, 0, 210, 42, "F");

  pdf.setTextColor(255, 255, 255);
  pdf.setFontSize(22);
  pdf.setFont("helvetica", "bold");
  pdf.text("DELLANCY", 15, 18);

  pdf.setFontSize(10);
  pdf.setFont("helvetica", "normal");
  pdf.text("BUSINESS ANALYTICS AND INTELLIGENCE", 15, 27);

  // Report title

  pdf.setTextColor(...navy);
  pdf.setFontSize(17);
  pdf.setFont("helvetica", "bold");
  pdf.text(currentReport.title, 15, 55);

  // Report information

  pdf.setTextColor(90, 100, 120);
  pdf.setFontSize(9);
  pdf.setFont("helvetica", "normal");

  pdf.text("Dataset: " + currentReport.datasetName, 15, 64);

  pdf.text(
    "Generated: " + currentReport.created.toLocaleString(),
    15,
    70
  );

  let startY = 82;

  // Automatic report

  if (currentReport.type === "automatic") {

    pdf.setTextColor(...purple);
    pdf.setFontSize(12);
    pdf.setFont("helvetica", "bold");
    pdf.text("DATASET OVERVIEW", 15, startY);

    pdf.autoTable({
      startY: startY + 5,
      head: [[
        "Total Records",
        "Total Columns",
        "Numeric Columns",
        "Missing Cells"
      ]],
      body: [[
        String(currentReport.records),
        String(currentReport.columns),
        String(currentReport.numericColumns),
        String(currentReport.missingCells)
      ]],
      theme: "grid",
      styles: {
        fontSize: 9,
        cellPadding: 3
      },
      headStyles: {
        fillColor: purple
      }
    });

    const automaticRows = currentReport.statistics.map(stat => [
      stat.name,
      String(stat.count),
      formatNumber(stat.average),
      formatNumber(stat.minimum),
      formatNumber(stat.maximum)
    ]);

    pdf.autoTable({
      startY: pdf.lastAutoTable.finalY + 12,
      head: [[
        "Column",
        "Valid Values",
        "Average",
        "Minimum",
        "Maximum"
      ]],
      body: automaticRows,
      theme: "striped",
      styles: {
        fontSize: 8,
        cellPadding: 3
      },
      headStyles: {
        fillColor: navy
      }
    });

    startY = pdf.lastAutoTable.finalY + 12;

    pdf.setFontSize(10);
    pdf.setTextColor(...navy);

    pdf.text(
      "Data quality: " +
      currentReport.missingCells +
      " empty cells were identified.",
      15,
      startY
    );

  }

  // Custom report

  if (currentReport.type === "custom") {

    const stats = currentReport.statistics;

    pdf.setTextColor(...purple);
    pdf.setFontSize(12);
    pdf.setFont("helvetica", "bold");
    pdf.text("SELECTED METRIC", 15, startY);

    pdf.autoTable({
      startY: startY + 5,
      head: [[
        "Metric",
        "Valid Values",
        "Missing Values"
      ]],
      body: [[
        currentReport.metric,
        String(stats.count),
        String(stats.missing)
      ]],
      theme: "grid",
      styles: {
        fontSize: 9,
        cellPadding: 4
      },
      headStyles: {
        fillColor: purple
      }
    });

    pdf.autoTable({
      startY: pdf.lastAutoTable.finalY + 12,
      head: [[
        "Sum",
        "Average",
        "Minimum",
        "Maximum"
      ]],
      body: [[
        formatNumber(stats.sum),
        formatNumber(stats.average),
        formatNumber(stats.minimum),
        formatNumber(stats.maximum)
      ]],
      theme: "striped",
      styles: {
        fontSize: 9,
        cellPadding: 4
      },
      headStyles: {
        fillColor: navy
      }
    });

    startY = pdf.lastAutoTable.finalY + 12;

    pdf.setFontSize(10);
    pdf.setTextColor(...navy);

    pdf.text(
      "Data completeness: " +
      formatNumber(stats.missing) +
      " missing values out of " +
      formatNumber(currentReport.records) +
      " records.",
      15,
      startY
    );

  }

  // Add chart image

  if (currentReport.chartColumn && dataChart) {

    let imageY = startY + 12;

    if (imageY > 190) {
      pdf.addPage();
      imageY = 20;
    }

    pdf.setTextColor(...purple);
    pdf.setFontSize(12);
    pdf.setFont("helvetica", "bold");
    pdf.text("DATA VISUALISATION", 15, imageY);

    const chartImage = dataChart.toBase64Image();

    pdf.addImage(
      chartImage,
      "PNG",
      15,
      imageY + 5,
      180,
      85
    );

  }

  // Footer on every page

  const pageCount = pdf.getNumberOfPages();

  for (let page = 1; page <= pageCount; page++) {

    pdf.setPage(page);

    pdf.setFontSize(8);
    pdf.setTextColor(120, 130, 145);

    pdf.text(
      "Dellancy Analytics | Business Analytics and Intelligence",
      15,
      287
    );

    pdf.text(
      "Page " + page + " of " + pageCount,
      175,
      287
    );

  }

  const safeTitle = currentReport.title
    .replace(/[^a-z0-9-_]/gi, "_")
    .substring(0, 60);

  pdf.save("Dellancy_" + safeTitle + ".pdf");

});


/* ==========================================
   16. INDUSTRY EXPLORER
========================================== */

const industryInformation = {

  "real-estate": {
    title: "Real Estate Analytics",

    description:
      "Understand property performance through pricing, occupancy, " +
      "sales activity, rental performance and market information. " +
      "Dellancy can help property businesses organise their data " +
      "and identify important performance patterns.",

    services: [
      {
        title: "Property Performance",
        description:
          "Compare properties, prices, unit types and building performance."
      },
      {
        title: "Occupancy Analysis",
        description:
          "Measure occupied units against total available units."
      },
      {
        title: "Sales and Pricing",
        description:
          "Examine asking prices, completed sales and price per square foot."
      },
      {
        title: "Market Monitoring",
        description:
          "Track changes in property performance over time."
      }
    ],

    metrics: [
      "Occupancy Rate",
      "Average Sale Price",
      "Price per Sq Ft",
      "Vacancy Rate",
      "Rental Revenue",
      "Sales Volume"
    ]
  },

  hospitality: {
    title: "Hospitality Analytics",

    description:
      "Help hotels and hospitality businesses understand room " +
      "occupancy, revenue, guest activity and operational performance.",

    services: [
      {
        title: "Room Performance",
        description:
          "Monitor occupied rooms, available rooms and occupancy."
      },
      {
        title: "Revenue Analysis",
        description:
          "Measure room revenue and average daily rates."
      },
      {
        title: "Guest Activity",
        description:
          "Explore guest numbers, bookings and customer trends."
      },
      {
        title: "Operations",
        description:
          "Track service activity, costs and operational performance."
      }
    ],

    metrics: [
      "Occupancy Rate",
      "Room Revenue",
      "Average Daily Rate",
      "Revenue per Available Room",
      "Bookings",
      "Operating Costs"
    ]
  },

  retail: {
    title: "Retail Analytics",

    description:
      "Give retailers a clearer view of sales, products, inventory " +
      "and store performance. Organised business data can help " +
      "identify trends and monitor commercial activity.",

    services: [
      {
        title: "Sales Analysis",
        description:
          "Track revenue, transaction volumes and sales performance."
      },
      {
        title: "Product Performance",
        description:
          "Identify products generating the most sales."
      },
      {
        title: "Inventory Monitoring",
        description:
          "Organise stock levels and monitor product movement."
      },
      {
        title: "Store Performance",
        description:
          "Compare performance across different store locations."
      }
    ],

    metrics: [
      "Total Revenue",
      "Units Sold",
      "Average Transaction Value",
      "Stock Levels",
      "Gross Profit",
      "Sales Growth"
    ]
  },

  healthcare: {
    title: "Healthcare Analytics",

    description:
      "Support healthcare organisations in understanding service " +
      "activity, resource utilisation, operational performance " +
      "and service demand. Analysis should use appropriately " +
      "protected and authorised data.",

    services: [
      {
        title: "Service Activity",
        description:
          "Monitor appointments, visits and service volumes."
      },
      {
        title: "Resource Utilisation",
        description:
          "Understand how available resources are being used."
      },
      {
        title: "Operational Performance",
        description:
          "Track service delivery and operational indicators."
      },
      {
        title: "Capacity Monitoring",
        description:
          "Compare available capacity with service demand."
      }
    ],

    metrics: [
      "Patient Visits",
      "Appointments",
      "Bed Occupancy",
      "Average Waiting Time",
      "Resource Utilisation",
      "Service Volume"
    ]
  }

};


// Industry panel elements

const industryDetails = document.getElementById("industryDetails");
const industryTitle = document.getElementById("industryTitle");

const industryDescription =
  document.getElementById("industryDescription");

const industryServices =
  document.getElementById("industryServices");

const industryMetrics =
  document.getElementById("industryMetrics");

const closeIndustryButton =
  document.getElementById("closeIndustryButton");

const industryWorkspaceButton =
  document.getElementById("industryWorkspaceButton");


// Attach click events to industry cards

document.querySelectorAll("[data-industry]").forEach(card => {

  card.addEventListener("click", function() {

    const industryKey = this.dataset.industry;
    const information = industryInformation[industryKey];

    if (!information) return;

    industryTitle.textContent = information.title;

    industryDescription.textContent =
      information.description;

    industryServices.innerHTML =
      information.services.map(service => {

        return `
          <div class="detail-item">
            <strong>${escapeHTML(service.title)}</strong>
            <p>${escapeHTML(service.description)}</p>
          </div>
        `;

      }).join("");

    industryMetrics.innerHTML =
      information.metrics.map(metric => {

        return `
          <span class="metric-tag">
            ${escapeHTML(metric)}
          </span>
        `;

      }).join("");

    industryDetails.classList.remove("hidden");

    industryDetails.scrollIntoView({
      behavior: "smooth",
      block: "start"
    });

  });

});


// Close industry panel

closeIndustryButton.addEventListener("click", function() {

  industryDetails.classList.add("hidden");

});


// Go to workspace

industryWorkspaceButton.addEventListener("click", function() {

  document.getElementById("workspace").scrollIntoView({
    behavior: "smooth"
  });

});


/* ==========================================
   17. BUSINESS TEMPLATES
========================================== */

const businessTemplates = {

  property: {

    title: "Property Performance Template",

    description:
      "Record property information, asking prices, occupancy " +
      "and sales performance.",

    filename: "Dellancy_Property_Performance.csv",

    instructions:
      "Enter one property or property record per row. " +
      "Use consistent currency and area units throughout the dataset.",

    headers: [
      "Property_ID",
      "Property_Name",
      "Location",
      "Property_Type",
      "Total_Units",
      "Occupied_Units",
      "Asking_Price_AED",
      "Sale_Price_AED",
      "Area_SqFt",
      "Date"
    ],

    rows: [
      [
        "P001", "Building A", "Sharjah", "Residential",
        "100", "92", "1200000", "1150000", "1200", "2026-01-01"
      ],
      [
        "P002", "Building B", "Dubai", "Residential",
        "80", "75", "1500000", "1450000", "1400", "2026-01-02"
      ]
    ]

  },

  sales: {

    title: "Sales Tracker Template",

    description:
      "Record sales transactions, products, quantities, revenue " +
      "and sales representatives.",

    filename: "Dellancy_Sales_Tracker.csv",

    instructions:
      "Enter one transaction per row. Keep revenue and cost " +
      "figures in the same currency.",

    headers: [
      "Transaction_ID",
      "Date",
      "Product",
      "Category",
      "Quantity",
      "Unit_Price_AED",
      "Total_Revenue_AED",
      "Cost_AED",
      "Sales_Representative",
      "Location"
    ],

    rows: [
      [
        "S001", "2026-01-01", "Product A", "Electronics",
        "2", "500", "1000", "650", "John", "Sharjah"
      ],
      [
        "S002", "2026-01-02", "Product B", "Furniture",
        "1", "1200", "1200", "800", "Mary", "Dubai"
      ]
    ]

  },

  operations: {

    title: "Operations Tracker Template",

    description:
      "Record daily operational activities, assigned employees, " +
      "task status, costs and completion dates.",

    filename: "Dellancy_Operations_Tracker.csv",

    instructions:
      "Enter one operational task per row. Use consistent status " +
      "labels such as Pending, In Progress and Completed.",

    headers: [
      "Task_ID",
      "Date",
      "Department",
      "Task_Description",
      "Assigned_To",
      "Priority",
      "Status",
      "Start_Date",
      "Completion_Date",
      "Cost_AED"
    ],

    rows: [
      [
        "T001", "2026-01-01", "Maintenance",
        "Air conditioning inspection", "Employee A",
        "High", "Completed", "2026-01-01",
        "2026-01-02", "350"
      ],
      [
        "T002", "2026-01-02", "Cleaning",
        "Building inspection", "Employee B",
        "Medium", "In Progress", "2026-01-02",
        "", "150"
      ]
    ]

  }

};


// Template panel elements

const templateDetails =
  document.getElementById("templateDetails");

const templateTitle =
  document.getElementById("templateTitle");

const templateDescription =
  document.getElementById("templateDescription");

const templatePreviewTable =
  document.getElementById("templatePreviewTable");

const templateInstructions =
  document.getElementById("templateInstructions");

const downloadTemplateButton =
  document.getElementById("downloadTemplateButton");

const uploadTemplateButton =
  document.getElementById("uploadTemplateButton");

const closeTemplateButton =
  document.getElementById("closeTemplateButton");

const templateStatus =
  document.getElementById("templateStatus");

let selectedTemplate = null;


// Open a template preview

document.querySelectorAll("[data-template]").forEach(card => {

  card.addEventListener("click", function() {

    const templateKey = this.dataset.template;
    const template = businessTemplates[templateKey];

    if (!template) return;

    selectedTemplate = template;

    templateTitle.textContent = template.title;
    templateDescription.textContent = template.description;
    templateInstructions.textContent = template.instructions;

    templateStatus.textContent = "";
    templateStatus.className = "status-message";

    const tableHead =
      templatePreviewTable.querySelector("thead");

    const tableBody =
      templatePreviewTable.querySelector("tbody");

    tableHead.innerHTML = "";
    tableBody.innerHTML = "";

    // Build column headings

    const headingRow = document.createElement("tr");

    template.headers.forEach(header => {

      const heading = document.createElement("th");
      heading.textContent = header;

      headingRow.appendChild(heading);

    });

    tableHead.appendChild(headingRow);

    // Build example rows

    template.rows.forEach(row => {

      const tableRow = document.createElement("tr");

      row.forEach(value => {

        const cell = document.createElement("td");
        cell.textContent = value;

        tableRow.appendChild(cell);

      });

      tableBody.appendChild(tableRow);

    });

    templateDetails.classList.remove("hidden");

    templateDetails.scrollIntoView({
      behavior: "smooth",
      block: "start"
    });

  });

});


// Close template preview

closeTemplateButton.addEventListener("click", function() {

  templateDetails.classList.add("hidden");
  selectedTemplate = null;

});


// Convert a value into a CSV cell

function escapeCSV(value) {

  const text = String(value ?? "");

  return '"' + text.replace(/"/g, '""') + '"';

}


// Download CSV template

downloadTemplateButton.addEventListener("click", function() {

  if (!selectedTemplate) {

    templateStatus.textContent =
      "Please select a template first.";

    return;

  }

  const template = selectedTemplate;

  const allRows = [
    template.headers,
    ...template.rows
  ];

  const csvContent = allRows.map(row => {

    return row.map(value => escapeCSV(value)).join(",");

  }).join("\r\n");

  // UTF-8 BOM helps spreadsheet applications read the file correctly

  const csvWithBOM = "\uFEFF" + csvContent;

  const blob = new Blob([csvWithBOM], {
    type: "text/csv;charset=utf-8;"
  });

  const downloadURL = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = downloadURL;
  link.download = template.filename;

  document.body.appendChild(link);
  link.click();
  link.remove();

  setTimeout(() => {
    URL.revokeObjectURL(downloadURL);
  }, 1000);

  templateStatus.textContent =
    "Your template has been prepared for download.";

  templateStatus.className = "status-message success";

});


// Upload a completed template

uploadTemplateButton.addEventListener("click", function() {

  if (!selectedTemplate) {

    templateStatus.textContent =
      "Please select a template first.";

    return;

  }

  templateStatus.textContent =
    "Choose your completed CSV or Excel template below.";

  templateStatus.className = "status-message";

  // Move to the existing workspace upload area

  document.getElementById("workspace").scrollIntoView({
    behavior: "smooth"
  });

  fileInput.click();

});


/* ==========================================
   18. NAVIGATION AND FOOTER
========================================== */

// Automatically update the copyright year

document.getElementById("currentYear").textContent =
  new Date().getFullYear();


// Highlight the navigation section being viewed

const navigationLinks =
  document.querySelectorAll(".nav-links a");

const pageSections =
  document.querySelectorAll("main section[id]");

const observer = new IntersectionObserver(function(entries) {

  entries.forEach(entry => {

    if (entry.isIntersecting) {

      navigationLinks.forEach(link => {

        link.classList.remove("active");

        if (
          link.getAttribute("href") === "#" + entry.target.id
        ) {
          link.classList.add("active");
        }

      });

    }

  });

}, {
  threshold: 0.25
});


pageSections.forEach(section => {
  observer.observe(section);
});


/* ==========================================
   END OF DELLANCY APPLICATION
========================================== */
