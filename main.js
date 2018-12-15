
   

//-------------------------------------------
// Modules to control application life and create native browser window
//-------------------------------------------
const {app,BrowserWindow,ipcMain} = require('electron')

//-------------------------------------------
// Keep a global reference of the window object, if you don't, the window will
// be closed automatically when the JavaScript object is garbage collected.
//-------------------------------------------
let mainWindow

function createWindow () {
	
	process.env['ELECTRON_DISABLE_SECURITY_WARNINGS'] = 'true';
	
  //-------------------------------------------	
  // Create the browser window.
  //-------------------------------------------
  mainWindow = new BrowserWindow({width: 1000, height: 600});
  
  //remove frame = frame:false
  
var interProcessComunicaton = require('electron').ipcMain;

interProcessComunicaton.on('myEvent', function(event, argument){
    //event.sender is of type webContents, more on this later
    //argument is 'myArgument'
});

	//-------------------------------------------
	// and load the index.html of the app.
	//-------------------------------------------  
	mainWindow.loadFile('index.html');	
	console.log('*** load html');
	
	//-------------------------------------------
	//open chrome dev tools
	//-------------------------------------------
	
	//	mainWindow.webContents.openDevTools();

   	//-------------------------------------------
	// Call the function to read the excel file
	//-------------------------------------------
	 read_mklink_excel();		
	 console.log('*** read excel');
	
	// Emitted when the window is closed.

mainWindow.on('closed', function () {
    // Dereference the window object, usually you would store windows
    // in an array if your app supports multi windows, this is the time
    // when you should delete the corresponding element.
    mainWindow = null
})

}

//--------------------------------------
// Read excel and create\remove folders
//--------------------------------------
function read_mklink_excel(){	 

	 var XLSX = require('xlsx');
	var workbook = XLSX.readFile('./mklink.xlsx');
	var sheet_name_list = workbook.SheetNames;	
	var xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
		

	// var child = require('child_process').execFile;
	// var executablePath = "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE";
	// var parameters = ["mklink.xlsx"];

	// child(executablePath, parameters, function(err, data) {
	// 	console.log(err)
	// 	console.log(data.toString());
	// });

	//-------------------------------------------		
	//Get hostname
	//-------------------------------------------
	var os = require("os");
	var hostname = os.hostname();
	console.log("===========================");
	console.log("HOST = "+hostname);
	console.log("===========================");
	
	var link_folder = "";
	
	//-------------------------------------------
	//Loop all rows in sheet
	//-------------------------------------------
	for(i in xlData)
	{		
		//-------------------------------------------
		//Get current variables
		//-------------------------------------------
		var location = xlData[i].location;
		var destination = xlData[i].destination;
		var source = xlData[i].source;
		var mklink_status = xlData[i].status;
		link_folder = xlData[i].link_folder;		
		

		console.log("=== "+i+ " "+location+ " "+ destination);
		
			//-------------------------------------------
			//Check hostname
			//-------------------------------------------
			if(location === hostname)
			{				
				console.log(i+" CURRENT LOCATION "+location+" "+hostname);
				
				var mklink="";
				
				//-------------------------------------------
				//Create new folder
				//-------------------------------------------
				if( mklink_status === "Y" )
				{				
					//-------------------------------------------
					//call function to create folder
					//-------------------------------------------
					create_folder(destination,link_folder,source);
					
				}	
				
				//-------------------------------------------
				//Remove folder
				//-------------------------------------------
				else if( mklink_status === "N" )
				{					
			
					//-------------------------------------------
					//Check folder exists and if exists call the function to remove it					
					//-------------------------------------------
					check_folder_exists(destination,link_folder);
				} //remove folder if exists
								
			}//check hostname

			console.log("");
			
			delete rmdir_check_cmd;
			delete rmdir_cmd;
			delete cmd;
			
	}//loop excel rows
	
	//-------------------------------------------		
	//Json data to renderer.js
	//-------------------------------------------
	
	console.log('*** CALL RENDERER');
	
	
	mainWindow.webContents.on('did-finish-load',WindowsReady);

function WindowsReady() {
    console.log('**** Ready');
	mainWindow.webContents.send('someMessage', xlData,hostname);	
	console.log('*** AFTER CALL RENDERER');
}

	
}//end of read excel function

//-------------------------------------------
// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
//-------------------------------------------
app.on('ready', createWindow)




//-------------------------------------------
// Quit when all windows are closed.
//-------------------------------------------
app.on('window-all-closed', function () {
  // On OS X it is common for applications and their menu bar
  // to stay active until the user quits explicitly with Cmd + Q
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', function () {
  // On OS X it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  
  if (mainWindow === null) {
    createWindow()
  }
})

//--------------------------------------------
// Quit app
//--------------------------------------------
ipcMain.on('quit-app', function() {
	// tray.window.close(); // Standart Event of the BrowserWindow object.
	app.quit(); // Standart event of the app - that will close our app.
})

//--------------------------------------------
// Reload excel
//--------------------------------------------
ipcMain.on('reload_excel', function() {		
	read_mklink_excel();
	mainWindow.reload();
})

//-----------------------------------------------
// Create folder
//-----------------------------------------------
function create_folder(destination,link_folder,source)
{
					//-------------------------------------------
					// create the mklink command string
					//-------------------------------------------
					mklink = 'mklink /D "'+destination+'\\'+link_folder+'" "'+source+'"';
					
					console.log('CREATE FOLDER CMD = '+mklink);			
					var cmd = require('node-cmd');

					//-------------------------------------------
					//Execute the mklind command
					//-------------------------------------------
					cmd.get(
					mklink,function(err,data,stderr){
						console.log('MESSAGE:',data)						
						}
					);
}

//-----------------------------------------------
// Check folder exists
//-----------------------------------------------
function check_folder_exists(dest,folder)
{	
					//-------------------------------------------
					// Create the dir command to check if the folder exists
					//-------------------------------------------
					var check_folder_cmd = 'dir "'+dest+'" | find "'+folder+'" | find "SYMLINKD"';
					
					console.log("REMOVE FOLDER CMD = "+check_folder_cmd);

					var rmdir_check_cmd = require('node-cmd');
					
					//-------------------------------------------
					//Execute the dir command					
					//-------------------------------------------
					rmdir_check_cmd.get(
					
							check_folder_cmd,function(err,data,stderr){								
								
								console.log('REMOVE:',data)


								var folder_exists = ""; 

								//-------------------------------------------								
								//Update the folder_exists variable
								//-------------------------------------------
								if(data.toString().length >0)
								{							
									folder_exists = "Yes";
									
									//-------------------------------------------
									//Remove folder		
									//-------------------------------------------
									console.log("REMOVING FOLDER "+dest+'\\'+folder);
									
									//-------------------------------------------
									//Call the function to remove the folder
									//-------------------------------------------
									var remove_cmd = 'rmdir "'+dest+'\\'+folder+'"';
									remove_folder(remove_cmd);
									
								}else
								{						
									folder_exists = "No"		

									//-------------------------------------------		
									//The folder is missing so just log this the the command screen.
									//-------------------------------------------
									console.log("FOLDER "+dest+'\\'+folder+'"'+" does not exist");
									
								}
								
							}//check_folder_cmd
					); //Check folder exists command					
}					

//-----------------------------------------------
// Remove folder
//-----------------------------------------------
function remove_folder(remove_cmd_str)
{
		console.log("REMOVE FOLDER "+remove_cmd_str);
	
		//-------------------------------------------
		//Execute the command to remove the folder
		//-------------------------------------------
		var rmdir_cmd = require('node-cmd');										
		rmdir_cmd.get(
		remove_cmd_str,function(err,data2,stderr){
		console.log('REMOVE FOLDER:',data2)								
	}
 );							
}