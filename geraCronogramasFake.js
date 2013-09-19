var fileSystem = new ActiveXObject("Scripting.FileSystemObject");
var objShell = WScript.CreateObject("WScript.Shell");
var projectApp = new ActiveXObject("MSProject.Application");

var path = objShell.CurrentDirectory;
var folder = fileSystem.GetFolder(path);

// Constantes do project
pjFixedUnits = 0
pjFixedDuration = 1
pjFixedWork = 2

var fc = new Enumerator(folder.Files)
try {
    for(;!fc.atEnd(); fc.moveNext() ) {
	var mppFilePath = fc.item().Path
	if(!mppFilePath.match(/.mpp$/)) continue
	if(mppFilePath.match(/Fake/)) continue

	projectApp.FileOpen(mppFilePath);
	var thisProject = projectApp.ActiveProject

	WScript.Echo("================")
	WScript.Echo("Iterating over project :" + thisProject.Name)
	WScript.Echo("================")

	for(var i=1; i<=thisProject.Tasks.Count ; i++) {
	    var task = thisProject.Tasks.Item(i)
	    WScript.Echo("iterating over task : " + thisProject.Tasks.Item(i).Name);
	    try {	    
		task.Type = pjFixedDuration;
		var percWork = task.PercentWorkComplete;
		task.Work = 1;
		task.PercentWorkComplete = percWork;
	    }
	    catch(e){}
	}

	for(var i=1; i<=thisProject.Resources.Count ; i++) {
	    var resource = thisProject.Resources.Item(i);
	    WScript.Echo("iterating over resource : " + resource.Name);
	    resource.StandardRate = 0;
	}

	thisProject.SaveAs(mppFilePath.replace(/.mpp$/, "Fake.mpp"))
    }
}
catch(e){}

finally {
    projectApp.Quit();
}
