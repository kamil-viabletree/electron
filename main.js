const { app, BrowserWindow, Menu, ipcMain } = require("electron");
const path = require("path");

app.setMaxListeners(15);

let mainWindow;
const isMac = process.platform === "darwin";
const isDev = true;

function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 700,
    height: 530,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    },
  });

  // Open the dev tools window on load
  // if (isDev) mainWindow.webContents.openDevTools();

  mainWindow.loadFile(path.join(__dirname, "renderer/index.html"));
}

const menu = [
  //   {
  //     label: "File",
  //     submenu: [
  //       {
  //         label: "Quit",
  //         click: () => app.quit(),
  //         accelerator: "CmdOrCtrl+W",
  //       },
  //     ],
  //   },
];

app.whenReady().then(() => {
  createMainWindow();

  // Implement menu items
  const mainMenu = Menu.buildFromTemplate(menu);
  Menu.setApplicationMenu(mainMenu);

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createMainWindow();
  });
});

app.on("window-all-closed", () => {
  // If not its a mac
  if (!isMac) app.quit();
});
