const {
  getAllClients,
  saveClient,
  calculateTaxes,
  getAllMunicipalities,
  saveMunicipality,
} = require("./db");
const { ipcRenderer } = require("electron"); // Добавете това най-горе в renderer.js

// Поставете това в самото начало на renderer.js, веднага под първия ред с require:
window.addEventListener("DOMContentLoaded", () => {
  if (typeof updateClientDropdown === "function") updateClientDropdown();
  if (typeof updateMunicipalityDropdown === "function")
    updateMunicipalityDropdown();
});

const fs = require("fs");
const iconv = require("iconv-lite");
const ExcelJS = require("exceljs");
const {
  Document,
  Packer,
  Paragraph,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  TextRun,
} = require("docx");

let currentCalculation = null;
// ОБНОВЕНА ФУНКЦИЯ ЗА ПАДАЩОТО МЕНЮ С ОБЩИНИ
function updateMunicipalityDropdown() {
  const select = document.getElementById("municipality");
  if (!select) return; // Защита от грешки

  select.innerHTML = ""; // Изчистваме старите общини

  // Вземаме динамичните общини от db.js
  const munis = getAllMunicipalities();

  // Пълним менюто с новите общини
  Object.keys(munis).forEach((key) => {
    const opt = document.createElement("option");
    opt.value = key;
    opt.innerText = key;
    select.appendChild(opt);
  });
}

// 1. ИНИЦИАЛИЗИРАНЕ НА ПАДАЩОТО МЕНЮ
function updateClientDropdown() {
  const select = document.getElementById("client-select");
  select.innerHTML = '<option value="new">-- Създай нов клиент --</option>';
  const clients = getAllClients();
  clients.forEach((c) => {
    const opt = document.createElement("option");
    opt.value = c.egn;
    opt.innerText = `${c.name} (${c.egn})`;
    select.appendChild(opt);
  });
}
updateClientDropdown();

// 2. СЪБИТИЕ ПРИ СМЯНА НА КЛИЕНТ ОТ МЕНЮТО (РАБОТИ СУПЕР БЪРЗО)
document.getElementById("client-select").addEventListener("change", (e) => {
  // Изчистваме старите резултати
  document.getElementById("result-table").innerHTML =
    "Тук ще се генерира справката...";
  document.getElementById("export-docx-btn").style.display = "none";
  currentCalculation = null;

  if (e.target.value === "new") {
    document.getElementById("name").value = "";
    document.getElementById("egn").value = "";
    document.getElementById("bulstat").value = "";
    document.getElementById("municipality").value = "Sozopol";
    document.getElementById("rooms").value = "3";
    document.getElementById("isPensioner").checked = false;
    return;
  }

  const clients = getAllClients();
  const client = clients.find((c) => c.egn === e.target.value);

  if (client) {
    document.getElementById("name").value = client.name || "";
    document.getElementById("egn").value = client.egn || "";
    document.getElementById("bulstat").value = client.bulstat || "";
    document.getElementById("municipality").value =
      client.municipality || "Sozopol";
    document.getElementById("rooms").value = client.rooms || "3";
    document.getElementById("isPensioner").checked =
      client.isPensioner || false;
  }
});

// 3. БУТОН: ЗАПАЗИ ПРОФИЛА НА КЛИЕНТА
document.getElementById("save-btn").addEventListener("click", () => {
  const client = {
    name: document.getElementById("name").value,
    egn: document.getElementById("egn").value,
    bulstat: document.getElementById("bulstat").value,
    municipality: document.getElementById("municipality").value,
    rooms: parseInt(document.getElementById("rooms").value) || 0,
    isPensioner: document.getElementById("isPensioner").checked,
  };
  if (!client.name || !client.egn) return alert("Грешка: Въведете име и ЕГН!");
  saveClient(client);
  alert("Профилът на клиента е записан успешно!");
  updateClientDropdown();
  document.getElementById("client-select").value = client.egn; // Маркира го автоматично в менюто
});

// 4. БУТОН: ИЗЧИСЛИ ДАНЪЦИ И ТАКСИ
document.getElementById("calc-btn").addEventListener("click", () => {
  const client = {
    municipality: document.getElementById("municipality").value,
    rooms: parseInt(document.getElementById("rooms").value) || 0,
    isPensioner: document.getElementById("isPensioner").checked,
  };

  const monthData = {
    nights: parseInt(document.getElementById("nights").value) || 0,
    incomePerNight: parseFloat(document.getElementById("price").value) || 0,
    avgGuests: parseInt(document.getElementById("guests").value) || 0,
    platformCommission:
      parseFloat(document.getElementById("platformCommission").value) || 0,
    monthsWorked: parseInt(document.getElementById("monthsWorked").value) || 1,
  };

  currentCalculation = calculateTaxes(client, monthData);

  document.getElementById("result-table").innerHTML = `
        <table>
            <tr><th>Перо</th><th>Сума (EUR)</th></tr>
            <tr style="color: #27ae60; font-weight: bold; background: #e8f8f5;">
                <td>Брутен приход</td>
                <td>+ ${currentCalculation.bruto.toFixed(2)}</td>
            </tr>
            <tr style="color: #c0392b;">
                <td>Комисиона платформи (${monthData.platformCommission}%)</td>
                <td>- ${currentCalculation.commission.toFixed(2)}</td>
            </tr>
            <tr style="color: #c0392b;">
                <td>Авансови осигуровки (за ${monthData.monthsWorked} мес.)</td>
                <td>- ${currentCalculation.advanceOsh.toFixed(2)}</td>
            </tr>
            <tr style="color: #c0392b;">
                <td>ДДС по чл. 97а</td>
                <td>- ${currentCalculation.dds97a.toFixed(2)}</td>
            </tr>
            <tr style="color: #c0392b;">
                <td>Туристически данък</td>
                <td>- ${currentCalculation.touristTax.toFixed(2)}</td>
            </tr>
            <tr style="color: #c0392b;">
                <td>Патентен данък (за ${monthData.monthsWorked} мес.)</td>
                <td>- ${currentCalculation.patent.toFixed(2)}</td>
            </tr>
            <tr style="color: #c0392b; font-weight: bold;">
                <td>Годишно довнасяне осигуровки</td>
                <td>- ${currentCalculation.finalOsh.toFixed(2)}</td>
            </tr>
            <tr style='background:${currentCalculation.netProfit.toFixed(2) > 0 ? "#27ae60" : "#c0392b"} ;color: white; font-weight: bold;'>
                <td>ЧИСТА ПЕЧАЛБА ЗА СОБСТВЕНИКА</td>
                <td> ${currentCalculation.netProfit.toFixed(2) > 0 ? +currentCalculation.netProfit.toFixed(2) : currentCalculation.netProfit.toFixed(2)}</td>
            </tr>
        </table>
    `;

  document.getElementById("export-docx-btn").style.display = "block";
});

// 5. БУТОН: ЕКСПОРТИРАЙ РАЗЧЕТА В WORD (.DOCX)
document.getElementById("export-docx-btn").addEventListener("click", () => {
  if (!currentCalculation) return alert("Няма налични данни!");

  const name = document.getElementById("name").value || "Клиент";
  const municipality = document.getElementById("municipality").value;
  const months = document.getElementById("monthsWorked").value;
  const commissionInput = document.getElementById("platformCommission").value;

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            text: `ФИНАНСОВ РАЗЧЕТ ЗА КРАТКОСРОЧЕН НАЕМ`,
            heading: "Heading1",
            alignment: AlignmentType.CENTER,
          }),
          new Paragraph({ text: "" }),
          new Paragraph({ text: `Собственик: ${name}`, font: "Arial" }),
          new Paragraph({
            text: `Дестинация (Община): ${municipality} | Активен сезон: ${months} месеца`,
            font: "Arial",
          }),
          new Paragraph({ text: "" }),

          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              createRow("Перо", "Сума (EUR)", true, "333333"),
              createRow(
                "Брутен приход от резервации",
                `+ ${currentCalculation.bruto.toFixed(2)}`,
                true,
                "27AE60",
              ),
              createRow(
                `Комисиона на платформите (${commissionInput}%)`,
                `- ${currentCalculation.commission.toFixed(2)}`,
                false,
                "C0392B",
              ),
              createRow(
                `Авансови осигуровки като СОЛ (за ${months} мес.)`,
                `- ${currentCalculation.advanceOsh.toFixed(2)}`,
                false,
                "C0392B",
              ),
              createRow(
                "Месечен ДДС по чл. 97а (20% върху комисиона)",
                `- ${currentCalculation.dds97a.toFixed(2)}`,
                false,
                "C0392B",
              ),
              createRow(
                "Местен туристически данък за нощувки",
                `- ${currentCalculation.touristTax.toFixed(2)}`,
                false,
                "C0392B",
              ),
              createRow(
                `Патентен данък към общината (за ${months} мес.)`,
                `- ${currentCalculation.patent.toFixed(2)}`,
                false,
                "C0392B",
              ),
              createRow(
                "Годишно довнасяне на осигуровки в НАП (до 30 април)",
                `- ${currentCalculation.finalOsh.toFixed(2)}`,
                true,
                "C0392B",
              ),
              createRow(
                "ЧИСТА ПЕЧАЛБА ЗА СОБСТВЕНИКА (НЕТО)",
                `${currentCalculation.netProfit.toFixed(2) > 0 ? +currentCalculation.netProfit.toFixed(2) : currentCalculation.netProfit.toFixed(2)}`,
                true,
                `${currentCalculation.netProfit.toFixed(2) > 0 ? "27AE60" : "C0392B"}`,
              ),
            ],
          }),
        ],
      },
    ],
  });

  Packer.toBuffer(doc).then((buffer) => {
    const cleanName = name.replace(/\s+/g, "_");
    const fileName = `Razchet_${cleanName}_${municipality}.docx`;
    try {
      fs.writeFileSync(fileName, buffer);
      alert(`Документът бе генериран успешно: ${fileName}`);
    } catch (error) {
      if (error.code === "EBUSY")
        alert(`⚠️ Файлът е отворен в Word! Затворете го и опитайте пак.`);
    }
  });
});

// 6. БУТОН: ГЕНЕРАТОР НА ОФИЦИАЛЕН ПРОТОКОЛ ПО ЧЛ. 117 [nra.bg]
document.getElementById("gen-117").addEventListener("click", () => {
  if (!currentCalculation) return alert("Първо изчислете данъците!");

  const name = document.getElementById("name").value || "Избран_Клиент";
  const egn = document.getElementById("egn").value || "";
  const bulstat = document.getElementById("bulstat").value || egn;
  const commission = currentCalculation.commission.toFixed(2);
  const ddsAmount = currentCalculation.dds97a.toFixed(2);

  const today = new Date();
  const formattedDate = today.toLocaleDateString("bg-BG");
  const protocolNumber =
    "00000000" + Math.floor(Math.random() * 100).toString();

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            text: `П Р О Т О К О Л`,
            heading: "Heading1",
            alignment: AlignmentType.CENTER,
          }),
          new Paragraph({
            text: `№ ${protocolNumber.slice(-10)} / Дата: ${formattedDate} г.`,
            alignment: AlignmentType.CENTER,
          }),
          new Paragraph({
            text: `[на основание чл. 117, ал. 2 от ЗДДС]`,
            alignment: AlignmentType.CENTER,
          }),
          new Paragraph({ text: "" }),
          new Paragraph({ text: `1. ИЗДАТЕЛ: ${name}` }),
          new Paragraph({ text: `ЕИК по БУЛСТАТ / ЕГН: ${bulstat}` }),
          new Paragraph({ text: "" }),
          new Paragraph({
            text: `2. ДОСТАВЧИК: Airbnb Ireland UC / Booking.com B.V.`,
          }),
          new Paragraph({ text: "" }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              createRow(
                "Данъчна основа (Комисиона)",
                `${commission} EUR`,
                true,
                "333333",
              ),
              createRow("Ставка на данъка", "20 %", false, "333333"),
              createRow(
                "Начислен ДДС (за внасяне)",
                `${ddsAmount} EUR`,
                true,
                "C0392B",
              ),
            ],
          }),
        ],
      },
    ],
  });

  Packer.toBuffer(doc).then((buffer) => {
    const fileName = `Protokol_117_${egn}_06.docx`;
    try {
      fs.writeFileSync(fileName, buffer);
      alert(`Протоколът по чл. 117 бе генериран успешно!`);
    } catch (error) {
      if (error.code === "EBUSY") alert(`⚠️ Затворете стария протокол в Word!`);
    }
  });
});

// ПОМОЩНА ФУНКЦИЯ ЗА WORD РЕДОВЕ
function createRow(col1, col2, isBold = false, hexColor = "333333") {
  const cell1Text = new TextRun({
    text: col1,
    font: "Arial",
    bold: isBold,
    color: hexColor,
  });
  const cell2Text = new TextRun({
    text: col2,
    font: "Arial",
    bold: isBold,
    color: hexColor,
  });
  return new TableRow({
    children: [
      new TableCell({
        width: { size: 70, type: WidthType.PERCENTAGE },
        children: [new Paragraph({ children: [cell1Text] })],
      }),
      new TableCell({
        width: { size: 30, type: WidthType.PERCENTAGE },
        children: [new Paragraph({ children: [cell2Text] })],
      }),
    ],
  });
}

// 7. ГЕНЕРАТОР: EXCEL ТУРИСТИЧЕСКИ ДАНЪК [sozopol.bg]
document.getElementById("gen-tourism").addEventListener("click", async () => {
  if (!currentCalculation) return alert("Първо изчислете данъците!");
  const name = document.getElementById("name").value;
  const municipality = document.getElementById("municipality").value;
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Справка Туризъм");
  sheet.columns = [
    { header: "Параметър", width: 35 },
    { header: "Стойност", width: 30 },
  ];
  sheet.addRow(["Декларатор:", name]);
  sheet.addRow(["Община:", municipality]);
  sheet.addRow([
    "Дължим туристически данък (EUR):",
    currentCalculation.touristTax.toFixed(2),
  ]);
  const fileName = `Spravka_Turizam_${name.replace(/\s+/g, "_")}.xlsx`;
  await workbook.xlsx.writeFile(fileName);
  alert(`Екселският файл ${fileName} е готов!`);
});

// 8. ОФИЦИАЛЕН ГЕНЕРАТОР: НАП ДЕКЛАРАЦИЯ ОБРАЗЕЦ 1 (.TXT WIN-1251) [nra.bg]
document.getElementById("gen-d1").addEventListener("click", () => {
  if (!currentCalculation)
    return alert("Грешка: Първо изчислете данъците за периода!");

  const name = document.getElementById("name").value || "Клиент";
  const egn = document.getElementById("egn").value;
  const isPensioner = document.getElementById("isPensioner").checked;

  if (!egn || egn.length !== 10) {
    return alert("Грешка: Моля въведете валидно 10-цифрено ЕГН или ЛНЧ!");
  }

  // Вземаме датите динамично за текущия отчетен период
  const today = new Date();
  const currentMonth = String(today.getMonth() + 1).padStart(2, "0"); // напр. "06"
  const currentYear = String(today.getFullYear()); // "2026"

  // Определяне на официалния Код за вид осигурен по НАП стандартите [nra.bg]
  // Код 12 = Самоосигуряващо се лице (СОЛ)
  // Код 13 = Самоосигуряващо се лице - пенсионер
  const vidOsiguren = isPensioner ? "13" : "12";

  // Минималният осигурителен доход (550.66 EUR) превърнат в стотинки, без десетична запетая [nra.bg]
  const modFormatted = "055066";

  // Сглобяване на официалния НАП ред (строго фиксирани 73 позиции) [nra.bg]
  // ЕГН(10) + Месец(2) + Година(4) + Вид осигурен(2) + Мин.Доход(6) + Декл.Доход(6) + Нули за останалите пера(33)
  // НАМЕРЕТЕ ТОЗИ РЕД В БУТОНА ГЕНЕРИРАЙ Д1:
  // const officialNapRow = `${egn}${currentMonth}${currentYear}${vidOsiguren}${modFormatted}${modFormatted}000000000000000000000000000000001\r\n`;

  // 🛠️ ЗАМЕНЕТЕ ГО С ТОЗИ (Изрязани са 2 нули, за да закове точно 73 позиции):
  const officialNapRow = `${egn}${currentMonth}${currentYear}${vidOsiguren}${modFormatted}${modFormatted}000000000000000000000000000001\r\n`;

  // Конвертиране от UTF-8 към Windows-1251 (Задължително изискване на НАП) [nra.bg]
  const buffer = iconv.encode(officialNapRow, "windows-1251");

  const fileName = `D1_${egn}_месец_${currentMonth}.txt`;

  try {
    fs.writeFileSync(fileName, buffer);
    alert(
      `🎉 УСПЕХ: Файлът за Декларация Образец 1 бе генериран успешно!\nНаименование: ${fileName}\n\nФайлът е готов за директно подаване в Електронния портал на НАП.`,
    );
  } catch (error) {
    alert(`Грешка при запис на файла: ${error.message}`);
  }
});

// БУТОН: ДОБАВИ / ОБНОВИ ОБЩИНА (Постави го най-отдолу в renderer.js)
// 1. ОБНОВЕНА ФУНКЦИЯ ЗА ПЪЛНЕНЕ НА МЕНЮТАТА (Заменете старата updateMunicipalityDropdown)
// 1. КОРИГИРАНА ФУНКЦИЯ ЗА КЛИЕНТИТЕ (С ОПРЕСНЯВАНЕ НА ГРАФИКАТА)
function updateClientDropdown() {
  const select = document.getElementById("client-select");
  if (!select) return;
  select.innerHTML = '<option value="new">-- Създай нов клиент --</option>';
  getAllClients().forEach((c) => {
    const opt = document.createElement("option");
    opt.value = c.egn;
    opt.innerText = `${c.name} (${c.egn})`;
    select.appendChild(opt);
  });

  // 🚀 ПУСКАМЕ ХАРДУЕРНОТО ОПРЕСНЯВАНЕ
  ipcRenderer.send("force-ui-refresh");
}
// 2. СЪБИТИЕ ПРИ ИЗБОР НА ОБЩИНА ЗА РЕДАКЦИЯ (В ТАБ 3)
document.getElementById("muni-select").addEventListener("change", (e) => {
  const selectedMuni = e.target.value;

  const nameInput = document.getElementById("muniName");
  const touristInput = document.getElementById("muniTouristTax");
  const patentInput = document.getElementById("muniPatentTax");
  const saveBtn = document.getElementById("save-muni-btn");

  if (selectedMuni === "new") {
    // Ако е нова община, отключваме полето за име и изчистваме
    nameInput.value = "";
    nameInput.disabled = false;
    touristInput.value = "";
    patentInput.value = "";
    saveBtn.innerText = "➕ Запиши Общината в Системата";
    return;
  }

  // Ако изберем съществуваща, четем данните ѝ от JSON файла
  const munis = getAllMunicipalities();
  const muniData = munis[selectedMuni];

  if (muniData) {
    nameInput.value = selectedMuni;
    nameInput.disabled = true; // Заключваме името, за да не се промени ключа в базата данни
    touristInput.value = muniData.touristTaxPerPerson;
    patentInput.value = muniData.patentPerRoomAnnual;
    saveBtn.innerText = "💾 Обнови данните на общината";
  }
});

// 1. ОБНОВЕНА ФУНКЦИЯ ЗА ПЪЛНЕНЕ НА МЕНЮТАТА (Заменете старата updateMunicipalityDropdown)
// 2. КОРИГИРАНА ФУНКЦИЯ ЗА ОБЩИНИТЕ (С ОПРЕСНЯВАНЕ НА ГРАФИКАТА)
function updateMunicipalityDropdown() {
  const clientMuniSelect = document.getElementById("municipality");
  const settingsMuniSelect = document.getElementById("muni-select");
  if (!clientMuniSelect || !settingsMuniSelect) return;

  clientMuniSelect.innerHTML = "";
  settingsMuniSelect.innerHTML =
    '<option value="new">-- Създай нова община --</option>';

  const munis = getAllMunicipalities();
  Object.keys(munis).forEach((key) => {
    const opt1 = document.createElement("option");
    opt1.value = key;
    opt1.innerText = key;
    clientMuniSelect.appendChild(opt1);

    const opt2 = document.createElement("option");
    opt2.value = key;
    opt2.innerText = key;
    settingsMuniSelect.appendChild(opt2);
  });

  // 🚀 ПУСКАМЕ ХАРДУЕРНОТО ОПРЕСНЯВАНЕ
  ipcRenderer.send("force-ui-refresh");
}

// 2. СЪБИТИЕ ПРИ ИЗБОР НА ОБЩИНА ЗА РЕДАКЦИЯ (В ТАБ 3)
// document.getElementById("muni-select").addEventListener("change", (e) => {
//   const selectedMuni = e.target.value;

//   const nameInput = document.getElementById("muniName");
//   const touristInput = document.getElementById("muniTouristTax");
//   const patentInput = document.getElementById("muniPatentTax");
//   const saveBtn = document.getElementById("save-muni-btn");

//   if (selectedMuni === "new") {
//     // Ако е нова община, отключваме полето за име и изчистваме
//     nameInput.value = "";
//     nameInput.disabled = false;
//     touristInput.value = "";
//     patentInput.value = "";
//     saveBtn.innerText = "➕ Запиши Общината в Системата";
//     return;
//   }

//   // Ако изберем съществуваща, четем данните ѝ от JSON файла
//   const munis = getAllMunicipalities();
//   const muniData = munis[selectedMuni];

//   if (muniData) {
//     nameInput.value = selectedMuni;
//     nameInput.disabled = true; // Заключваме името, за да не се промени ключа в базата данни
//     touristInput.value = muniData.touristTaxPerPerson;
//     patentInput.value = muniData.patentPerRoomAnnual;
//     saveBtn.innerText = "💾 Обнови данните на общината";
//   }
// });
// ПОСТАВЕТЕ ТОВА НАЙ-ДОЛУ В RENDERER.JS
document.getElementById("save-muni-btn").addEventListener("click", () => {
  const nameInput = document.getElementById("muniName");

  // Временно отключваме полето за име, за да прочетем текста му (ако е било заключено при редакция)
  const wasDisabled = nameInput.disabled;
  nameInput.disabled = false;

  const muniName = nameInput.value.trim();
  const touristTax = parseFloat(
    document.getElementById("muniTouristTax").value,
  );
  const patentTax = parseFloat(document.getElementById("muniPatentTax").value);

  // Връщаме състоянието
  nameInput.disabled = wasDisabled;

  if (!muniName || isNaN(touristTax) || isNaN(patentTax)) {
    return alert("Грешка: Моля попълнете име на община, такса и патент!");
  }

  // Записваме в db.js
  saveMunicipality(muniName, touristTax, patentTax);
  alert(`Данните за община "${muniName}" бяха успешно записани/обновени!`);

  // Нулираме формите
  nameInput.value = "";
  nameInput.disabled = false; // Задължително го отключваме
  document.getElementById("muniTouristTax").value = "";
  document.getElementById("muniPatentTax").value = "";
  document.getElementById("save-muni-btn").innerText =
    "➕ Запиши Общината в Системата";

  // Обновяваме падащите менюта в програмата
  if (typeof updateMunicipalityDropdown === "function") {
    updateMunicipalityDropdown();
  }

  document.getElementById("muni-select").value = "new";
});
