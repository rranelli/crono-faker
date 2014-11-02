using System;
using System.Collections.Generic;

namespace CronoFakerBSK
{
  public class TaskInfo
  {
    public string Id { get; private set; }
    public string Name { get; private set; }
    public double Work { get; private set; }
    public double ActualWork { get; private set; }
    public double PercentWorkComplete { get; private set; }
    public DateTime Start { get; private set; }
    public DateTime Finish { get; private set; }
    public int OutlineLevel { get; private set; }
    public bool HasChildren { get; private set; }

    public TaskInfo(string id, string name, double work, double actualWork, double percentWorkComplete, DateTime start, DateTime finish, int outlineLevel, bool hasChildren)
    {
      HasChildren = hasChildren;
      Id = id;
      Name = name;
      Work = work;
      ActualWork = actualWork;
      PercentWorkComplete = percentWorkComplete;
      Start = start;
      Finish = finish;
      OutlineLevel = outlineLevel;
    }

    public override bool Equals(object value)
    {
      var type = value as TaskInfo;
      return (type != null)
	&& Equals(type.Id, Id)
	&& EqualityComparer<string>.Default.Equals(type.Name, Name)
	&& EqualityComparer<double>.Default.Equals(type.Work, Work)
	&& EqualityComparer<double>.Default.Equals(type.ActualWork, ActualWork)
	&& EqualityComparer<double>.Default.Equals(type.PercentWorkComplete, PercentWorkComplete)
	&& EqualityComparer<DateTime>.Default.Equals(type.Start, Start)
	&& EqualityComparer<DateTime>.Default.Equals(type.Finish, Finish);
    }

    public override int GetHashCode()
    {
      var num = 0x7a2f0b42;
      num = (-1521134295*num) + EqualityComparer<string>.Default.GetHashCode(Id);
      num = (-1521134295*num) + EqualityComparer<string>.Default.GetHashCode(Name);
      num = (-1521134295*num) + EqualityComparer<double>.Default.GetHashCode(Work);
      num = (-1521134295*num) + EqualityComparer<double>.Default.GetHashCode(ActualWork);
      num = (-1521134295*num) + EqualityComparer<double>.Default.GetHashCode(PercentWorkComplete);
      num = (-1521134295*num) + EqualityComparer<DateTime>.Default.GetHashCode(Start);
      return (-1521134295 * num) + EqualityComparer<DateTime>.Default.GetHashCode(Finish);
    }
  }
}
