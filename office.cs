
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;


public class Startup
{
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);
    
    
    [DllImport("oleacc.dll", SetLastError = true)]
    internal static extern IntPtr GetProcessHandleFromHwnd(IntPtr hwnd);


    [DllImport("kernel32.dll")]
    internal static extern int GetProcessId(IntPtr handle);

    const UInt32 WM_CLOSE = 0x0010;

    Word.Application msword;
    Excel.Application msexcel;
    PowerPoint.Application mspowerpoint;
    TaskFactory scheduler;
    public async Task<object> Invoke(object input)
    {
        scheduler = new TaskFactory(new LimitedConcurrencyLevelTaskScheduler(1));
        
        return new { 
          word = (Func<object,Task<dynamic>>)word,
          excel = (Func<object,Task<dynamic>>)excel,
          powerPoint = (Func<object,Task<dynamic>>)powerPoint,
          close = (Func<object,Task<dynamic>>)close,
        };
    }
    
    void Delay(int stage) {
        //Thread.Sleep(100);
    }
    
    private void CreateWord() {
        lock(this) {
            if (msword == null) {
                msword = new Word.Application();
                msword.Visible = true;
                msword.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                msword.ChangeFileOpenDirectory(Directory.GetCurrentDirectory());
                msword.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            }
        }
    }
    
    private void CreateExcel() {
        lock (this) {
           if (msexcel == null) {
                msexcel = new Excel.Application();
                msexcel.Visible = true;
                msexcel.DisplayAlerts = false;
                msexcel.DefaultFilePath = Directory.GetCurrentDirectory();
                msexcel.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                msexcel.AskToUpdateLinks = false;
            }
        }
    }
    
    private void CreatePowerPoint() {
        lock(this) {
           if (mspowerpoint == null) {
                mspowerpoint = new PowerPoint.Application();
                mspowerpoint.Visible = MsoTriState.msoTrue;
                mspowerpoint.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                mspowerpoint.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                
                //No way to do this in powerpoint, it seems
                //mspowerpoint.?Path? = Directory.GetCurrentDirectory();
            }
        }
    }
    
    public async Task<dynamic> word(dynamic opts) {
    
        var file = Path.GetFullPath(opts.input as string);
        var pdfFile = Path.GetFullPath(opts.output as string);
            
        return await scheduler.StartNew(() => {
            Thread.Sleep(100);
            CreateWord();
            
            var doc = msword.Documents.OpenNoRepairDialog(file, false, true, false, "", "", true, "", "", Type.Missing, Type.Missing, true, true, Type.Missing, true, Type.Missing);
            Delay(0);
            try {
                Delay(1);
                doc.ExportAsFixedFormat(pdfFile, Word.WdExportFormat.wdExportFormatPDF);
                Delay(2);
            }
            finally {
                (doc as Word._Document).Close(false);
                Marshal.ReleaseComObject(doc);
            }
            
            //closeInternal();
            return pdfFile;
        });
    }
    
    public async Task<dynamic> excel(dynamic opts) {
    
        var file = Path.GetFullPath(opts.input as string);
        var pdfFile = Path.GetFullPath(opts.output as string);
              
        return await scheduler.StartNew(() => {
            Thread.Sleep(100);
            CreateExcel();
            
            var book = msexcel.Workbooks.Open(file, 2, true, Type.Missing, "", "", false, Type.Missing, Type.Missing, false, false, Type.Missing, false, true, Excel.XlCorruptLoad.xlNormalLoad);
            Delay(0);
            
            try {
                Delay(1);
                book.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfFile, Excel.XlFixedFormatQuality.xlQualityStandard, true, false);
                Delay(2);
            } finally {
                (book as Excel._Workbook).Close(false);
                Marshal.ReleaseComObject(book);
            }
            
            //closeInternal();
            return pdfFile;
        });
    }
    
    public async Task<dynamic> powerPoint(dynamic opts) {
    
        var file = Path.GetFullPath(opts.input as string);
        var pdfFile = Path.GetFullPath(opts.output as string);
          
        return await scheduler.StartNew(() => {
            Thread.Sleep(100);
            CreatePowerPoint();
            
            var presso = mspowerpoint.Presentations.Open(file, MsoTriState.msoTrue, WithWindow: MsoTriState.msoFalse);
            Delay(0);
            try {
                Delay(1);
              //  presso.ExportAsFixedFormat(pdfFile, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF, PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint);
              presso.SaveAs(pdfFile, PowerPoint.PpSaveAsFileType.ppSaveAsPDF);
                Delay(2);
            } finally {
                presso.Close();
                Marshal.ReleaseComObject(presso);
            }
            
            //closeInternal();
            return pdfFile;
        });
    }
    
    private void closeInternal() {
        Delay(4);
        Console.WriteLine("Time to quit now ..\n");
        
        if (msword != null) {
          try {
              (msword as Word._Application).Quit();
          }
          catch (Exception) { }
          Marshal.ReleaseComObject(msword);
          msword = null;
        }
        
        
        if (msexcel != null) {
            try {
                msexcel.Quit();
            }
            catch (Exception) { }
            
            Marshal.ReleaseComObject(msexcel);
            msexcel = null;
        }

        if (mspowerpoint != null) {
            try {
                mspowerpoint.Quit();
            }
            catch (Exception) { }
          
            Marshal.ReleaseComObject(mspowerpoint);
            mspowerpoint = null;
        }
        
        Thread.Sleep(1000);
    }
    //Close open office apps on cleanup
    public async Task<object> close(dynamic opts) {
        //Powerpoint & co sometimes need some more time before they can quit grafeully
        return await scheduler.StartNew(() => {
            closeInternal();
            return true;
        });
    }
    
    ~Startup() {
      close(null).Wait();
    }
}

// <summary>
/// Provides a task scheduler that ensures a maximum concurrency level while
/// running on top of the ThreadPool.
/// </summary>
public class LimitedConcurrencyLevelTaskScheduler : TaskScheduler
{
    /// <summary>Whether the current thread is processing work items.</summary>
    [ThreadStatic]
    private static bool _currentThreadIsProcessingItems;
    /// <summary>The list of tasks to be executed.</summary>
    private readonly LinkedList<Task> _tasks = new LinkedList<Task>(); // protected by lock(_tasks)
    /// <summary>The maximum concurrency level allowed by this scheduler.</summary>
    private readonly int _maxDegreeOfParallelism;
    /// <summary>Whether the scheduler is currently processing work items.</summary>
    private int _delegatesQueuedOrRunning = 0; // protected by lock(_tasks)

    /// <summary>
    /// Initializes an instance of the LimitedConcurrencyLevelTaskScheduler class with the
    /// specified degree of parallelism.
    /// </summary>
    /// <param name="maxDegreeOfParallelism">The maximum degree of parallelism provided by this scheduler.</param>
    public LimitedConcurrencyLevelTaskScheduler(int maxDegreeOfParallelism)
    {
        if (maxDegreeOfParallelism < 1) throw new ArgumentOutOfRangeException("maxDegreeOfParallelism");
        _maxDegreeOfParallelism = maxDegreeOfParallelism;
    }

    /// <summary>Queues a task to the scheduler.</summary>
    /// <param name="task">The task to be queued.</param>
    protected sealed override void QueueTask(Task task)
    {
        // Add the task to the list of tasks to be processed.  If there aren't enough
        // delegates currently queued or running to process tasks, schedule another.
        lock (_tasks)
        {
            _tasks.AddLast(task);
            if (_delegatesQueuedOrRunning < _maxDegreeOfParallelism)
            {
                ++_delegatesQueuedOrRunning;
                NotifyThreadPoolOfPendingWork();
            }
        }
    }

    /// <summary>
    /// Informs the ThreadPool that there's work to be executed for this scheduler.
    /// </summary>
    private void NotifyThreadPoolOfPendingWork()
    {
        ThreadPool.UnsafeQueueUserWorkItem(_ =>
        {
            // Note that the current thread is now processing work items.
            // This is necessary to enable inlining of tasks into this thread.
            _currentThreadIsProcessingItems = true;
            try
            {
                // Process all available items in the queue.
                while (true)
                {
                    Task item;
                    lock (_tasks)
                    {
                        // When there are no more items to be processed,
                        // note that we're done processing, and get out.
                        if (_tasks.Count == 0)
                        {
                            --_delegatesQueuedOrRunning;
                            break;
                        }

                        // Get the next item from the queue
                        item = _tasks.First.Value;
                        _tasks.RemoveFirst();
                    }

                    // Execute the task we pulled out of the queue
                    base.TryExecuteTask(item);
                }
            }
            // We're done processing items on the current thread
            finally { _currentThreadIsProcessingItems = false; }
        }, null);
    }

    /// <summary>Attempts to execute the specified task on the current thread.</summary>
    /// <param name="task">The task to be executed.</param>
    /// <param name="taskWasPreviouslyQueued"></param>
    /// <returns>Whether the task could be executed on the current thread.</returns>
    protected sealed override bool TryExecuteTaskInline(Task task, bool taskWasPreviouslyQueued)
    {
        // If this thread isn't already processing a task, we don't support inlining
        if (!_currentThreadIsProcessingItems) return false;

        // If the task was previously queued, remove it from the queue
        if (taskWasPreviouslyQueued) TryDequeue(task);

        // Try to run the task.
        return base.TryExecuteTask(task);
    }

    /// <summary>Attempts to remove a previously scheduled task from the scheduler.</summary>
    /// <param name="task">The task to be removed.</param>
    /// <returns>Whether the task could be found and removed.</returns>
    protected sealed override bool TryDequeue(Task task)
    {
        lock (_tasks) return _tasks.Remove(task);
    }

    /// <summary>Gets the maximum concurrency level supported by this scheduler.</summary>
    public sealed override int MaximumConcurrencyLevel { get { return _maxDegreeOfParallelism; } }

    /// <summary>Gets an enumerable of the tasks currently scheduled on this scheduler.</summary>
    /// <returns>An enumerable of the tasks currently scheduled.</returns>
    protected sealed override IEnumerable<Task> GetScheduledTasks()
    {
        bool lockTaken = false;
        try
        {
            Monitor.TryEnter(_tasks, ref lockTaken);
            if (lockTaken) return _tasks.ToArray();
            else throw new NotSupportedException();
        }
        finally
        {
            if (lockTaken) Monitor.Exit(_tasks);
        }
    }
}