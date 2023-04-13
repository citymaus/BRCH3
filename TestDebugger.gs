// TestDebugger.gs

//------------------------------------------------------------------------------------
// SKS:
// Occasionally debugger will fail to run if an egregious error exists in your script.
// Run the debugger on this function for a debugger sanity check. 
// If debugger runs, your script is good. If it does not, try a Run:
// and watch for exception toasts up top.
//
// ...Google Apps debugger is weird sometimes
//------------------------------------------------------------------------------------
function testDebugger() {  
  Logger.log("debugger working");
  let test = "";
}

