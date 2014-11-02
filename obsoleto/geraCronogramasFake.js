// ##########
// CronoFaker chemtech para a braskem
//
// Este script remove as informacoes confidenciais da chemtech dos
// cronogramas .mpp que devem ser enviados ao planejamento bsk
// ##########

var fileSystem = new ActiveXObject("Scripting.FileSystemObject");
var objShell = WScript.CreateObject("WScript.Shell");
var projectApp = new ActiveXObject("MSProject.Application");

var path = objShell.CurrentDirectory;
WScript.Echo(path)
var folder = fileSystem.GetFolder(path);

// Constantes do project
pjFixedUnits    = 0
pjFixedDuration = 1
pjFixedWork     = 2

var fc = new Enumerator(folder.Files)
try {
    for(;!fc.atEnd(); fc.moveNext() ) {
        var mppFilePath = fc.item().Path;
        if(!mppFilePath.match(/.mpp$/)) continue;
        if(mppFilePath.match(/\/Fake.*.mpp$/)) continue;
        var timesFactor = 1.2 + Math.floor(Math.random()*101)/100
        
        projectApp.FileOpen(mppFilePath);
        var thisProject = projectApp.ActiveProject

        WScript.Echo("================")
        WScript.Echo("Iterating over project :" + thisProject.Name)
        WScript.Echo("================")
    
        for(var i=1; i<=thisProject.Resources.Count ; i++) {
            var resource = thisProject.Resources.Item(i);
            WScript.Echo("iterating over resource : " + resource.Name);
            try{ // Se resource.type <> Work, o codigo abaixo quebra. A classe resource viola o good-citizen principle, viola liskov e tudo mais.
                resource.StandardRate = 0;
                resource.OvertimeRate = 0;
            }
            catch(e) {WScript.Echo(e);}
        }
        
        for(var i=1; i<=thisProject.Tasks.Count ; i++) {
            var task = thisProject.Tasks.Item(i)
            WScript.Echo("iterating over task : " + thisProject.Tasks.Item(i).Name);
            try {
                var startDate = task.Start;
                var finishDate = task.Finish;
                task.ConstraintType = 0; // quando uma tarefa tem a finish ou start date setada na mao, cria-se uma constraint que quebra a macro.
                task.Type = pjFixedDuration;
                
                task.Work = task.Work * timesFactor;
                task.ActualWork = task.ActualWork * timesFactor;
                
                task.BaselineCost = 0;
                task.BaselineWork = 0;
                task.Baseline1Cost = 0;
                task.Baseline1Work = 0;
                
                task.Start = startDate;
                task.Finish = finishDate;
            }
            catch(e){WScript.Echo(e);}
        }
    
        thisProject.SaveAs(mppFilePath.replace(thisProject.Name, "FakeProporc-" + thisProject.Name));
    }
}
catch(e){WScript.Echo(e);}

finally {
    projectApp.Quit();
}
