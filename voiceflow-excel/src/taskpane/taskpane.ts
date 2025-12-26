Office.onReady(() => {
  const btn = document.getElementById("startVoiceBtn");
  if (btn) btn.onclick = startListening;
});

function startListening() {
  const SpeechRecognition =
    (window as any).webkitSpeechRecognition ||
    (window as any).SpeechRecognition;

  if (!SpeechRecognition) {
    alert("Speech recognition not supported");
    return;
  }

  const recognition = new SpeechRecognition();
  recognition.lang = "en-US";

  recognition.onresult = async (event: any) => {
    const text = event.results[0][0].transcript;
    (document.getElementById("outputText") as HTMLElement).innerText = text;
    await sendToBackend(text);
  };

  recognition.start();
}

async function sendToBackend(text: string) {
  const res = await fetch("http://localhost:8000/command", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text }),
  });

  const plan = await res.json();
  console.log("📥 PLAN:", plan);

  await executeExcelPlan(plan);
}

async function executeExcelPlan(plan: any) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    switch (plan.action) {

  /* ---------------- WRITE ---------------- */
  case "write":
    if (!plan.cell) throw new Error("Missing cell");
    sheet.getRange(plan.cell).values = [[plan.value ?? ""]];
    break;

  /* ---------------- DELETE ---------------- */
  case "delete_cell":
    if (!plan.cell) throw new Error("Missing cell");
    sheet.getRange(plan.cell).clear(Excel.ClearApplyTo.contents);
    break;

  /* ---------------- INSERT ROW ---------------- */
  case "insert_row":
    if (!plan.row) throw new Error("Missing row");
    sheet
      .getRange(`${plan.row}:${plan.row}`)
      .insert(Excel.InsertShiftDirection.down);
    break;

  /* ---------------- INSERT COLUMN ---------------- */
  case "insert_column":
    if (!plan.column) throw new Error("Missing column");
    sheet
      .getRange(`${plan.column}:${plan.column}`)
      .insert(Excel.InsertShiftDirection.right);
    break;

  /* ---------------- SUM ---------------- */
  case "sum":
  if (!plan.range) throw new Error("Missing range");

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const columnLetter = plan.range.split(":")[0];

    const usedRange = sheet.getUsedRange();
    usedRange.load("rowCount");
    await context.sync();

    const rowCount = usedRange.rowCount;

    const columnRange = sheet.getRange(
      `${columnLetter}1:${columnLetter}${rowCount}`
    );
    columnRange.load("values");
    await context.sync();

    const values = columnRange.values;
    let lastRow = 0;

    for (let i = 0; i < values.length; i++) {
      if (values[i][0] !== "" && values[i][0] !== null) {
        lastRow = i + 1;
      }
    }

    const sumTarget = `${columnLetter}${lastRow + 1}`;

    // 🚨 CRITICAL FIX: exclude the sum cell
    const sumFormula = `=SUM(${columnLetter}1:${columnLetter}${lastRow})`;

    sheet.getRange(sumTarget).formulas = [[sumFormula]];
  });

  break;
  /*-----------------std----------------------- */
  case "stddev":
  if (!plan.range) throw new Error("Missing range");

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const column = plan.range.split(":")[0];

    const usedRange = sheet.getUsedRange();
    usedRange.load("rowCount");
    await context.sync();

    const rowCount = usedRange.rowCount;

    const colRange = sheet.getRange(`${column}1:${column}${rowCount}`);
    colRange.load("values");
    await context.sync();

    let lastRow = 0;
    for (let i = 0; i < colRange.values.length; i++) {
      if (colRange.values[i][0] !== "" && colRange.values[i][0] !== null) {
        lastRow = i + 1;
      }
    }

    const target = `${column}${lastRow + 1}`;
    const formula = `=STDEV.S(${column}1:${column}${lastRow})`;

    sheet.getRange(target).formulas = [[formula]];
  });

  break;
  /*---------------average------------------- */
  case "average":
  if (!plan.range) throw new Error("Missing range");

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const column = plan.range.split(":")[0];

    const usedRange = sheet.getUsedRange();
    usedRange.load("rowCount");
    await context.sync();

    const rowCount = usedRange.rowCount;

    const colRange = sheet.getRange(`${column}1:${column}${rowCount}`);
    colRange.load("values");
    await context.sync();

    let lastRow = 0;
    for (let i = 0; i < colRange.values.length; i++) {
      if (colRange.values[i][0] !== "" && colRange.values[i][0] !== null) {
        lastRow = i + 1;
      }
    }

    const target = `${column}${lastRow + 1}`;
    const formula = `=AVERAGE(${column}1:${column}${lastRow})`;

    sheet.getRange(target).formulas = [[formula]];
  });

  break;

  /* ---------------- BOLD ---------------- */
  case "bold":
    if (!plan.range) throw new Error("Missing range");
    sheet.getRange(plan.range).format.font.bold = true;
    break;

  /* ---------------- SORT ---------------- */
  case "sort":
    if (!plan.range) throw new Error("Missing range");

    sheet.getRange(plan.range).sort.apply(
      [
        {
          key: 0,
          ascending: plan.order !== "desc",
        },
      ],
      false // hasHeaders = false (safe default)
    );
    break;

  /* ---------------- FILTER ---------------- */
  case "filter":
    if (!plan.range || !plan.condition)
      throw new Error("Missing range or condition");

    sheet.autoFilter.apply(
      sheet.getRange(plan.range),
      0,
      {
        filterOn: Excel.FilterOn.custom,
        criterion1: plan.condition,
      }
    );
    break;

  /* ---------------- CREATE CHART ---------------- */
  case "create_chart":
    if (!plan.x_column || !plan.y_column)
      throw new Error("Missing x_column or y_column");

    const chartRange = sheet.getRange(
      `${plan.x_column}:${plan.y_column}`
    );

    const chart = sheet.charts.add(
      Excel.ChartType.columnClustered,
      chartRange,
      Excel.ChartSeriesBy.columns
    );

    chart.title.text = "Generated Chart";
    chart.legend.visible = true;
    break;

  
  default:
    console.error("❌ Unknown action:", plan);
}


    await context.sync();
  });
}
