/* global Excel */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = setupEventHandler;
    document.getElementById("app-body").style.display = "flex";
  }
});

// Hàm setup event handler
async function setupEventHandler() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Register onChange event
      sheet.onChanged.add(handleCellChange);

      await context.sync();

      console.log("Event handler registered successfully!");
      document.getElementById("status").innerHTML =
        "✅ Event handler active! Change A1 to see magic.";
    });
  } catch (error) {
    console.error(error);
    document.getElementById("status").innerHTML = "❌ Error: " + error.message;
  }
}

// Hàm xử lý khi có thay đổi
async function handleCellChange(event: Excel.WorksheetChangedEventArgs) {
  await Excel.run(async (context) => {
    console.log("Cell changed:", event.address);

    // Chỉ xử lý khi thay đổi ô A1
    if (event.address === "A1") {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const a1 = sheet.getRange("A1");
      a1.load("values");

      await context.sync();

      const value = a1.values[0][0] as string;
      let operator = "";

      switch (value) {
        case "A":
          operator = "+";
          break;
        case "B":
          operator = "-";
          break;
        case "C":
          operator = "*";
          break;
        default:
          return; // Không làm gì nếu giá trị không hợp lệ
      }

      // Thay đổi công thức
      sheet.getRange("C3").formulas = [["=A3" + operator + "A4"]];
      sheet.getRange("C4").formulas = [["=A4" + operator + "A5"]];

      await context.sync();

      console.log(`✅ Formulas updated with operator: ${operator}`);
      document.getElementById("status").innerHTML = `✅ Updated with operator: ${operator}`;
    }
  });
}
