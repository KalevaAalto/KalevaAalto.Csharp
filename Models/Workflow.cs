using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace KalevaAalto.Models
{
    public class Workflow
    {
        private Stopwatch _globalStopwatch = new Stopwatch();
        private Stopwatch _stopwatch = new Stopwatch();
        private string _subName;
        private Action<string>? _log;
        private string _workingContent = string.Empty;
        private List<Task> _tasks = new List<Task>();
        public string WorkingContent
        {
            set
            {
                if (!string.IsNullOrEmpty(_workingContent) && _log is not null)
                {
                    _log($"进程：{_subName}：{_workingContent}成功！！！" + _stopwatch.ClockString());
                }
                _stopwatch.Restart();
                _workingContent = value;
            }
        }
        public Workflow(Action<string>? log, string subName)
        {
            _log = log;
            _globalStopwatch.Restart();
            _subName = subName;
            Log(@"-----------------------------------------");
        }
        public Workflow(string subName, Action<string>? log)
        {
            _log = log;
            _globalStopwatch.Restart();
            _subName = subName;
            Log(@"-----------------------------------------");
        }
        public Workflow(string subName)
        {
            _globalStopwatch.Restart();
            _subName = subName;
            Log(@"-----------------------------------------");
        }
        public void Log(string str)
        {
            if (_log is not null) _log(str);
        }

        public void AddTask(Task task, string? workingContent = null)
        {

            _tasks.Add(Task.Run(async () =>
            {
                Stopwatch stopwatch = Stopwatch.StartNew();
                await task;
                if (!string.IsNullOrEmpty(workingContent))Log($"进程：{_subName}：{_workingContent}成功！！！" + stopwatch.ClockString());
            }));
        }

        public async Task WhenAllTask()
        {
            await Task.WhenAll(_tasks);
            _tasks.Clear();
        }

        public void End()
        {
            if (!string.IsNullOrEmpty(_workingContent))Log($"进程：{_subName}：{_workingContent}成功！！！" + _stopwatch.ClockString());
            Log(@"==============================================");
            Log($"进程：{_subName}：结束！！！！" + _globalStopwatch.ClockString());
            
        }
        public void Stop(string message = EmptyString)
        {
            if (string.IsNullOrEmpty(message))Log($"进程：{_subName}：{_workingContent}成功！！！" + _stopwatch.ClockString());
            else Log($"进程：{_subName}：{_workingContent}成功，{message}！！！" + _stopwatch.ClockString());
            _workingContent = string.Empty;
        }

        public void Error(Exception error)
        {
            Log(@"==============================================");
            Regex regex = new Regex(@"\s+");
            string errorMessage = regex.Replace(error.Source + "：" + error.Message, " ");
            Log($"进程：{_subName}：{(string.IsNullOrEmpty(_workingContent) ? _subName : _workingContent)}：异常：{errorMessage}");
        }

        public void Error(string errorMessage)
        {
            Log(@"==============================================");
            Log($"进程：{_subName}：{(string.IsNullOrEmpty(_workingContent) ? _subName : _workingContent)}：异常：{errorMessage}");
        }

    }
}
