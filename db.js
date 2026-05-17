const fs = require("fs");
const path = require("path");
const FILE_PATH = path.join(__dirname, "clients.json");

// 1. МАТРИЦА С МЕСТНИ ДАНЪЦИ И ТАКСИ (Централизирана конфигурация)
// const MUNICIPALITY_RATES = {
//   Sozopol: {
//     touristTaxPerPerson: 0.41, // 0.80 лв. за места клас "В" [sozopol.bg]
//     patentPerRoomAnnual: 35.79, // 70 лв. на година за стая [sozopol.bg]
//   },
//   Nesebar: {
//     touristTaxPerPerson: 0.33, // 0.65 лв. съгласно Наредба №14 на Община Несебър
//     patentPerRoomAnnual: 51.13, // 100 лв. на година за стая
//   },
//   // Тук лесно можете да добавите Варна, Бургас, Поморие и т.н.
// };

function getAllClients() {
  if (!fs.existsSync(FILE_PATH)) {
    fs.writeFileSync(FILE_PATH, JSON.stringify([]));
    return [];
  }
  return JSON.parse(fs.readFileSync(FILE_PATH, "utf-8"));
}

function saveClient(client) {
  const clients = getAllClients();
  const index = clients.findIndex((c) => c.egn === client.egn);

  if (index !== -1) {
    // Запазваме новите данни, като съхраняваме старата хронология (history)
    clients[index] = { ...clients[index], ...client };
  } else {
    client.history = [];
    clients.push(client);
  }
  fs.writeFileSync(FILE_PATH, JSON.stringify(clients, null, 2), "utf-8");
}

// 2. ДИНАМИЧЕН ИЗЧИСЛИТЕЛЕН МОДУЛ
// Променена функция - вече приема комисионата като динамичен параметър
function calculateTaxes(client, currentMonthData) {
  const bruto = currentMonthData.incomePerNight * currentMonthData.nights;

  // Динамична комисиона на платформата
  const commissionPercent = currentMonthData.platformCommission / 100;
  const commission = bruto * commissionPercent;
  const dds97a = commission * 0.2; // 20% ДДС върху комисионата

  // Данъчни ставки на Общината
  const munis = getAllMunicipalities();
  const selectedRates = munis[client.municipality];
  if (!selectedRates) {
    throw new Error(
      `Ненамерени данъчни ставки за община: ${client.municipality}`,
    );
  }

  const totalPeopleNights =
    currentMonthData.nights * currentMonthData.avgGuests;
  const touristTax = totalPeopleNights * selectedRates.touristTaxPerPerson;
  const patent =
    ((selectedRates.patentPerRoomAnnual * client.rooms) / 12) *
    currentMonthData.monthsWorked;

  // ОСИГУРИТЕЛНА ЛОГИКА (БЕЗУПРЕЧНА МАТЕМАТИКА) [nra.bg]
  const oshPercent = client.isPensioner ? 0.08 : 0.278;
  const minMoD = 550.66; // 1077 лв.
  const maxMoD = 2111.64; // 4130 лв.

  // 1. Изчисляваме точно какво е платено авансово за 3-те месеца общо
  const advanceOsh = minMoD * oshPercent * currentMonthData.monthsWorked;

  // 2. Намираме реалния средномесечен доход от дейността
  const realMonthlyIncome = bruto / currentMonthData.monthsWorked;

  // 3. НАП ограничава дохода до максималния осигурителен таван
  const boundedMonthlyIncome = Math.min(realMonthlyIncome, maxMoD);

  // 4. Изчисляваме колко общо трябва да се внесе за 3-те месеца на база реалния доход
  const totalDulzhaOsh = boundedMonthlyIncome * oshPercent * 3;

  // 5. Годишното довнасяне е разликата между общо дължимото и вече платеното авансово
  let finalOsh = 0;
  if (boundedMonthlyIncome > minMoD) {
    finalOsh = totalDulzhaOsh - advanceOsh;
  }

  // Чиста печалба в джоба след абсолютно всички пера
  const netProfit =
    bruto - (commission + advanceOsh + dds97a + touristTax + patent + finalOsh);

  return {
    bruto,
    commission,
    advanceOsh,
    dds97a,
    touristTax,
    patent,
    finalOsh,
    netProfit,
  };
}

// 3. ФУНКЦИЯ ЗА ЗАПИС НА МЕСЕЧНИЯ РЕГИСТЪР В ИСТОРИЯТА НА КЛИЕНТА
function archiveMonthToHistory(egn, monthKey, monthData, calculatedResults) {
  const clients = getAllClients();
  const client = clients.find((c) => c.egn === egn);

  if (!client) return false;

  const record = {
    month: monthKey, // напр. "06-2026"
    ...monthData,
    ...calculatedResults,
  };

  const existingIndex = client.history.findIndex((h) => h.month === monthKey);
  if (existingIndex !== -1) {
    client.history[existingIndex] = record; // Обновяване, ако вече има такъв месец
  } else {
    client.history.push(record); // Добавяне на нов месец
  }

  fs.writeFileSync(FILE_PATH, JSON.stringify(clients, null, 2), "utf-8");
  return true;
}

// 1. ФУНКЦИЯ ЗА ПРОЧИТАНЕ НА ВСИЧКИ ОБЩИНИ ОТ ФАЙЛА MUNICIPALITIES.JSON
function getAllMunicipalities() {
  const muniPath = path.join(__dirname, "municipalities.json");

  // Ако файлът още не съществува, го създаваме с базовите ставки за Несебър и Созопол [1.2.3, sozopol.bg]
  if (!fs.existsSync(muniPath)) {
    const defaultMunis = {
      Sozopol: { touristTaxPerPerson: 0.41, patentPerRoomAnnual: 35.79 },
      Nesebar: { touristTaxPerPerson: 0.33, patentPerRoomAnnual: 51.13 },
    };
    fs.writeFileSync(muniPath, JSON.stringify(defaultMunis, null, 2), "utf-8");
    return defaultMunis;
  }
  return JSON.parse(fs.readFileSync(muniPath, "utf-8"));
}

// 2. ФУНКЦИЯ ЗА ЗАПИС ИЛИ ОБНОВЯВАНЕ НА ОБЩИНА ВЪВ ФАЙЛА
function saveMunicipality(name, touristTax, patentTax) {
  const muniPath = path.join(__dirname, "municipalities.json");
  const munis = getAllMunicipalities();

  // Добавяме новата община в масива
  munis[name] = {
    touristTaxPerPerson: touristTax,
    patentPerRoomAnnual: patentTax,
  };

  fs.writeFileSync(muniPath, JSON.stringify(munis, null, 2), "utf-8");
}

module.exports = {
  getAllClients,
  saveClient,
  calculateTaxes,
  archiveMonthToHistory,
  getAllMunicipalities,
  saveMunicipality,
};
