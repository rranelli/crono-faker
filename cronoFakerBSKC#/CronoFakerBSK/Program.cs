using System;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Interop.MSProject;
using log4net;
using Application = Microsoft.Office.Interop.MSProject.Application;
using Exception = System.Exception;

namespace CronoFakerBSK
{
  class Program
  {
    public static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

    [STAThread]
    static void Main()
    {
      var execPath = StriperService.GetFilesPath();
      System.IO.Directory.CreateDirectory(execPath + "\\Fakes");

      var fileNames = System.IO.Directory.GetFiles(execPath)
	.Where(x => x.Contains(".mpp"));

      var projectApp = new Application();
      try
	{
	  fileNames.ToList()
	    .ForEach(projPath =>
		{
		  // recuperando o projeto do arquivo e retirando informações
		  projectApp.FileOpen(projPath);
		  var project = projectApp.ActiveProject;
		  var projName = project.Name;

		  var newName = StriperService.GetProjectName(project);
		  var fakePath = execPath + "\\Fakes\\" + newName;
		  System.IO.File.Delete(fakePath);

		  // Conversão
		  Project convertedProject = StriperService.MakeMaskedProjectFile(projectApp, newName, StriperService.GetScalingFactor(project), StriperService.MakeTaskInfoDic(project));
		  convertedProject.SaveAs(fakePath);

		  // Validação
		  var convertedInfoDic = StriperService.MakeTaskInfoDic(convertedProject);
		  projectApp.FileCloseAll(PjSaveType.pjDoNotSave);

		  projectApp.FileOpen(projPath);
		  var origInfoDic = StriperService.MakeTaskInfoDic(projectApp.ActiveProject);
		  projectApp.FileCloseAll(PjSaveType.pjDoNotSave);

		  new Validator(Log).ValidateAndLogEachTask(projName, origInfoDic, convertedInfoDic);
		});
	}
      catch (Exception ex) { Log.Error("Exceção não tratada", ex); }

      finally { projectApp.Quit(PjSaveType.pjDoNotSave); }
    }
  }
}
