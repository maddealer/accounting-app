// Отворете main.js и заменете съдържанието му с това:
const { app, BrowserWindow, ipcMain } = require("electron");
const path = require("path");

let win;

function createWindow() {
  win = new BrowserWindow({
    width: 1000,
    height: 800,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    },
  });

  win.loadFile("index.html");
}

app.whenReady().then(createWindow);

// 🛠️ ХАРДУЕРНИЯТ ФИКС: Приема сигнал от renderer.js и насилствено преначертава прозореца
ipcMain.on("force-ui-refresh", () => {
  if (win) {
    // Симулираме мигновено преоразмеряване с 1 пиксел и връщане обратно,
    // което принуждава Windows и Chromium да опреснят всички менюта за 1 милисекунда!
    const size = win.getSize();
    win.setSize(size[0] + 1, size[1]);
    win.setSize(size[0], size[1]);
  }
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
