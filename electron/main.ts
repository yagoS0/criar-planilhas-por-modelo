import {
  app, BrowserWindow, ipcMain, dialog,
} from 'electron';
import execute from '../src/controller/execute';

if (require('electron-squirrel-startup')) app.quit();

let mainWindow: BrowserWindow | null;

declare const MAIN_WINDOW_WEBPACK_ENTRY: string;
declare const MAIN_WINDOW_PRELOAD_WEBPACK_ENTRY: string;

// const assetsPath =
//   process.env.NODE_ENV === 'production'
//     ? process.resourcesPath
//     : app.getAppPath()

function createWindow() {
  // if (require('electron-squirrel-startup')) return app.quit()

  mainWindow = new BrowserWindow({
    // icon: path.join(assetsPath, 'assets', 'icon.png'),
    width: 1280,
    height: 1200,
    backgroundColor: '#191622',
    webPreferences: {
      preload: MAIN_WINDOW_PRELOAD_WEBPACK_ENTRY,
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true,
      nodeIntegrationInWorker: true,
      nodeIntegrationInSubFrames: true,

    },
  });

  mainWindow.loadURL(MAIN_WINDOW_WEBPACK_ENTRY);
  mainWindow.webContents.session.on(
    'will-download',
    (event, item, webContents) => {
      item.setSaveDialogOptions({ properties: ['createDirectory'] });
    },
  );

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

async function registerListeners() {
  /**
   * This comes from bridge integration, check bridge.ts
   */
  ipcMain.on('message', (_, message) => {
    console.log(message);
  });
}

ipcMain.on('send-openDialog', async (event) => {
  try {
    const dir = await dialog.showOpenDialog({
      properties: ['openDirectory', 'multiSelections'],
    });
    await execute(dir.filePaths[0]);
  } catch (error) {
    console.log(`Erro ocorrido: ${error}`);
  }
});

app.commandLine.appendSwitch('--no-sandbox');

app
  .on('ready', createWindow)
  .whenReady()
  .then(registerListeners)
  .catch((e) => console.error(e));

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});
