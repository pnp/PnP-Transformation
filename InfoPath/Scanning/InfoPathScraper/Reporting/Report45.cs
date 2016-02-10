using System;
using System.Collections.Generic;
using System.Linq;
using System.Collections.Concurrent;
using System.Threading.Tasks;

namespace InfoPathScraper.Reporting
{
	// partial implementation that is 4.5+ compatible
	partial class Report
	{
		private ConcurrentQueue<string> _processed = new ConcurrentQueue<string>();
		private List<Task> _runningTasks = new List<Task>();

		private int InProgress { get { return _runningTasks.Count; } }
		private int Complete { get { return _processed.Count; } }

		/// <summary>
		/// This method runs on the main thread only (as currently used) so it doesn't have to be thread-safe. 
		/// While we have work left to do, keep pulling paths out of our queue and spinning up Tasks to handle them.
		/// I cap the number of Tasks at once to 10, this is kind of arbitrary but I randomly decided it'd be better
		/// than perhaps spinning up 20,000 Tasks at once. Once we have a huge dataset we can experiment with this.
		/// Being limited to just this thread means I don't need to use Concurrent* collections.
		/// </summary>
		public void Run()
		{
			// while we have elements or outstanding work, fill up our slots.
			while (Remaining > 0 || InProgress > 0)
			{
				PurgeDoneTasks();
				while (InProgress < 10 && Remaining > 0)
				{
					Task t = ProcessNext(_unprocessed.Dequeue());
					_runningTasks.Add(t);
				}
				// in case all our Tasks are just sitting there, done, wake up every 5 seconds.
				// I don't think this is strictly necessary, but that's ok.
				Task.WaitAny(_runningTasks.ToArray(), 5000);
			}
		}

		/// <summary>
		/// We need to remove all completed Tasks, otherwise they'll clog up our queue
		/// </summary>
		private void PurgeDoneTasks()
		{
			for (int i = _runningTasks.Count - 1; i >= 0; i--)
			{
				if (_runningTasks[i].IsCompleted)
				{
					_runningTasks.RemoveAt(i);
				}
			}
		}

		/// <summary>
		/// Create a Task to load and run based on the path
		/// </summary>
		/// <param name="path"></param>
		/// <returns></returns>
		private Task ProcessNext(string path)
		{
			Task t = Task.Factory.StartNew(() =>
			{
				TaskBody(path);
			});
			return t;
		}

	}
}
