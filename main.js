const electron = require('electron');
// Module to control application life.
const app = electron.app;
// Module to create native browser window.
const BrowserWindow = electron.BrowserWindow;
const path = require('path');

// Keep a global reference of the window object, if you don't, the window will
// be closed automatically when the JavaScript object is garbage collected.
let mainWindow;

function createWindow () {
  // Create the browser window.
  mainWindow = new BrowserWindow({
    webPreferences: {
      nodeIntegration: true,
      preload: path.join(__dirname, 'preload.js')
    }
  });

  // and load the index.html of the app.
  mainWindow.loadURL('https://office.live.com/start/Word.aspx?omkt=en-GB&auth=1&nf=1');

  // Open the DevTools.
  mainWindow.webContents.openDevTools();

  mainWindow.webContents.on('dom-ready', function() {
    mainWindow.webContents.insertCSS("#h_bar { display: none; }");
    mainWindow.webContents.insertCSS("#b_dialogpanel { top: 0; }");
    mainWindow.webContents.insertCSS("#AppHeaderPanel { display: none; }");
  });

  mainWindow.webContents.on('new-window', function(event, url) {
    console.log(url);
    event.preventDefault();
    if (url === 'about:blank') {
      return;
    }
    mainWindow.loadURL(url);
  });

  // Emitted when the window is closed.
  mainWindow.on('closed', function () {
    // Dereference the window object, usually you would store windows
    // in an array if your app supports multi windows, this is the time
    // when you should delete the corresponding element.
    mainWindow = null;
  })
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.on('ready', createWindow);

// Quit when all windows are closed.
app.on('window-all-closed', function () {
  // On OS X it is common for applications and their menu bar
  // to stay active until the user quits explicitly with Cmd + Q
  if (process.platform !== 'darwin') {
    app.quit();
  }
})

app.on('activate', function () {
  // On OS X it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  if (mainWindow === null) {
    createWindow();
  }
});

// In this file you can include the rest of your app's specific main process
// code. You can also put them in separate files and require them here.
