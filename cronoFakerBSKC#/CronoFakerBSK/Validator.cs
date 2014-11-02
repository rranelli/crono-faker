using System;
using System.Collections.Generic;
using System.Linq;
using log4net;
using Microsoft.Office.Interop.MSProject;

namespace CronoFakerBSK
{
  public class Validator
  {
    private readonly ILog _log;

    public Validator(ILog logger)
    {
      _log = logger;
    }

    public Validator ValidateAndLogEachTask(string projName, Dictionary<string, TaskInfo> taskInfoDic,
					    IDictionary<string, TaskInfo> validationInfo)
    {
      taskInfoDic.ToList()
	.ForEach(kvp =>
	    {
	      var thisValidInfo = validationInfo[kvp.Key];
	      var thisTaskInfo = kvp.Value;

	      // validating finish date
	      var spanFinish = thisTaskInfo.Finish - thisValidInfo.Finish;
	      if (spanFinish.Hours > 12)
		{
		  _log.Warn(String.Format(
					  @"Consistency Error - Finish Date
                                             | Projeto => {0}
                                             | Tarefa => {1}
                                             | Finish Inicial = {2}
                                             | Finish Após Conversão {3}", projName, thisTaskInfo.Name,
					  thisTaskInfo.Finish, thisValidInfo.Finish).Trim());
		}

	      // validating start date
	      var spanStart = thisTaskInfo.Start - thisValidInfo.Start;
	      if (spanStart.Hours > 12)
		{
		  _log.Warn(String.Format(
					  @"Consistency Error - Start Date
                                             | Projeto => {0}
                                             | Tarefa => {1}
                                             | Start Inicial = {2}
                                             | Start Após Conversão {3}", projName, thisTaskInfo.Name,
					  thisTaskInfo.Start, thisValidInfo.Start).Trim());
		}

	      // validating percent work complete
	      if (Math.Abs(thisTaskInfo.PercentWorkComplete - thisValidInfo.PercentWorkComplete) > 1.1)
		{
		  _log.Warn(String.Format(
					  @"Consistency Error - PercentWorkComplete
                                          | Projeto => {0}
                                          | Tarefa => {1}
                                          | %Work Inicial = {2}
                                          | %Work Pós Conversão = {3}", projName, thisTaskInfo.Name,
					  thisTaskInfo.PercentWorkComplete, thisValidInfo.PercentWorkComplete).Trim());
		}
	    });
      return this;
    }

    public Validator ValidateAndLogProject(_IProjectDoc actvProject, dynamic originalFinishDate, dynamic originalStartDate,
					   dynamic originalPercWorkComplete)
    {
      var startSpan = (DateTime) originalStartDate - (DateTime) actvProject.Start;
      if(startSpan.Hours > 12)
	_log.Warn(String.Format("Consistency Error - Finish Date | Projeto => {0}", actvProject.Name));

      var finishSpan = (DateTime)originalFinishDate - (DateTime)actvProject.Finish;
      if (finishSpan.Hours > 12)
	_log.Warn(String.Format("Consistency Error - Finish Date | Projeto => {0}", actvProject.Name));
      if (Math.Abs(originalPercWorkComplete-actvProject.PercentWorkComplete) > 1.1)
	_log.Warn(String.Format("Consistency Error - % Work Complete | Projeto => {0}", actvProject.Name));
      _log.Info("----------");

      return this;
    }
  }
}
