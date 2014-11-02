using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.MSProject;
using Application = Microsoft.Office.Interop.MSProject.Application;

namespace CronoFakerBSK
{
  public static class PropositalBreaker
  {
    private const string TotalHoursComercial = "Horas (PM070)";
    private const string ProjectsClientNumber = "Número do Projeto (Cliente)";

    public static Project MakeMaskedProjectFile(Application projectApp, string newName, double scalingFactor, Dictionary<string, TaskInfo> taskInfoDic)
    {
      Program.Log.Info(String.Format("Iniciando converção do projeto :  {0} ", newName));
      Console.WriteLine("Convertendo o projeto :  {0} ", newName);

      projectApp.FileNew();
      var thisProject = projectApp.ActiveProject;
      thisProject.ProjectSummaryTask.Name = newName;

      taskInfoDic.ToList().ForEach(kvp =>
	  {
	    var taskInfo = kvp.Value;
	    var thisTask = thisProject.Tasks.Add(taskInfo.Name, thisProject.Tasks.Count + 1);

	    thisTask.Type = PjTaskFixedType.pjFixedWork;

	    Enumerable.Range(1, 10).ToList().ForEach(x => { try { thisTask.OutlineOutdent(); } catch { } });
	    Enumerable.Range(1, taskInfo.OutlineLevel - 1).ToList().ForEach(x => thisTask.OutlineIndent());

	    Debug.Assert(thisTask.OutlineLevel == taskInfo.OutlineLevel);
	    thisTask.Text29 = taskInfo.Id; //terrible workaround !

	    if (thisTask.OutlineChildren.Count == 0) return;

	    thisTask.Start = taskInfo.Start;
	    thisTask.Finish = taskInfo.Finish;

	    thisTask.Work = taskInfo.Work*scalingFactor;
	    thisTask.ActualWork = taskInfo.ActualWork*scalingFactor;

	    if (thisTask.Work == 0 && thisTask.ActualWork == 0)
	      thisTask.PercentWorkComplete = taskInfo.PercentWorkComplete;
	  });

      Program.Log.Info(String.Format("Converção completa do projeto  :  {0} ", newName));
      return thisProject;
    }

    public static string GetFilesPath()
    {
      var fbDialog = new FolderBrowserDialog
      {
	Description = "Selecione a pasta contendo os cronogramas para conversão"
      };
      fbDialog.ShowDialog();
      var execPath = fbDialog.SelectedPath;
      return execPath;
    }

    public static Dictionary<string,TaskInfo> MakeTaskInfoDic(_IProjectDoc actvProject)
    {
      var taskInfoDic = actvProject.Tasks.Cast<Task>().Select(
							      task =>
							      new TaskInfo(
									   task.Text29 != string.Empty // not proud of this.
									   ? task.Text29
									   : task.Guid,
									   task.Name,
									   task.Work,
									   task.ActualWork,
									   task.PercentWorkComplete,
									   task.Start,
									   task.Deadline.Equals("NA") ? task.Finish : task.Deadline,
									   task.OutlineLevel,
									   task.OutlineChildren.Count != 0)
							      ).ToDictionary(x => x.Id);

      return taskInfoDic;
    }

    public static dynamic GetScalingFactor(Project actvProject)
    {
      var scalingFactor = GetComericalHours(actvProject) / (actvProject.ProjectSummaryTask.Work / 60);
      return scalingFactor;
    }

    public static string GetProjectName(Project project)
    {
      try
	{
	  var fldCte = project.Application.FieldNameToFieldConstant(ProjectsClientNumber);
	  var fldVal = project.ProjectSummaryTask.GetField(fldCte);

	  return fldVal != String.Empty
	    ? fldVal
	    : project.Name;
	}
      catch { return project.Name; }
    }

    public static double GetComericalHours(Project project)
    {
      try
	{
	  var fieldVal =
	    project.ProjectSummaryTask.GetField(project.Application.FieldNameToFieldConstant(TotalHoursComercial));

	  return fieldVal != String.Empty
	    ? Convert.ToDouble(fieldVal)
	    : (new Random().NextDouble() + 1) * project.ProjectSummaryTask.Work / 60;
	}
      catch
	{
	  return (new Random().NextDouble() + 1) * project.ProjectSummaryTask.Work / 60;
	}
    }
  }
}
