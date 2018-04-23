using System;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;

namespace Selenium_Automated_Testing.Utilities
{
    public class PeriodicCapture
    {
        // File path prefix
        private readonly Common _clsCommon = new Common();
        private readonly string _tempDir;
        private readonly string _saveDir;
        // Screenshot period in milli-second.
        private const int RecordPeriod = 2000;
        // The instance of screen capture class
        private ScreenCapture _screenCapture;
        // Frame count
        private int _frameCount;
        // Inter-thread communication stuff
        private volatile EventType _eventType;
        private volatile string _eventDescription;
        private AutoResetEvent _reqEvent;
        private AutoResetEvent _ackEvent;
        private Thread _workerThread;
        private enum EventType { EventNone, EventStop, EventShot };
        // Worker thread internal use
        private EventType _receivedEventType;
        private string _receivedDescription;
        #region Main thread code
        // Constructor.
        public PeriodicCapture()
        {
            _tempDir = _clsCommon.StrRootPath + @"CAPTEMP\";
            _saveDir = _clsCommon.StrRootPath + @"CAPSAVE\";
        }
        // Start periodic recording.
        public void Start(string description)
        {
            Clear();
            _screenCapture  = new ScreenCapture();
            _frameCount     = 0;
            _eventType      = EventType.EventNone;
            _workerThread   = new Thread(WorkerThreadMain);
            _reqEvent       = new AutoResetEvent(false);
            _ackEvent       = new AutoResetEvent(false);
            _eventDescription = String.Format("Start {0} -",description);
            // Start worker thread.
            _workerThread.Start();
            // ReSharper disable once EmptyEmbeddedStatement
            while (!_workerThread.IsAlive);
        }
        // Take a screenshot right now.
        public void Shot(string description)
        {
            SendSyncEvent(EventType.EventShot, description);
        }
        // Stop recording.
        public void Stop()
        {
            SendSyncEvent(EventType.EventStop, null);
            _workerThread.Join();
            _reqEvent.Close();
            _ackEvent.Close();
            _eventType       = EventType.EventNone;
            _eventDescription = null;
            _workerThread    = null;
            _reqEvent        = null;
            _ackEvent        = null;
            _screenCapture   = null;
        }
        // Save screenshots as a zip file.
        public void Save(string suiteName)
        {
            // TODO: Compression is not implemented. Just moving files.
            foreach (string file in Directory.GetFiles(_saveDir, "*.*"))
            {
                File.Delete(file);
            }
            int count = 0;
            foreach (string src in Directory.GetFiles(_tempDir, "*.*"))
            {
                string dst = String.Format("{0}{1}_{2:D8}.PNG", _saveDir, suiteName, count);
                File.Move(src, dst);
                count++;
            }
        }
        // Clear temporary picture directory.
        public void Clear()
        {
            foreach (string file in Directory.GetFiles(_tempDir, "*.*"))
            {
                File.Delete(file);
            }
        }
        // Send an event to the worker thread.
        private void SendSyncEvent(EventType eventType, string description)
        {
            _eventType = eventType;
            _eventDescription = description;
            _reqEvent.Set();
            _ackEvent.WaitOne();
        }
        #endregion

        #region Worker thread code

        private void WorkerThreadMain()
        {
            while (WaitForEvent(RecordPeriod))
            {
                TakeScreenshot();
            }
        }
        private bool WaitForEvent(int wait)
        {
            // Get event with timer
            if (!_reqEvent.WaitOne(wait))
                return true; // timed out

            // Save event info
            _receivedEventType = _eventType;
            _receivedDescription = _eventDescription;
 
            // Send acknowledge to main thread.
            _ackEvent.Set();

            // Returns true if we should continue loop.
            return _receivedEventType != EventType.EventStop;
        }

        private void TakeScreenshot()
        {
            if (_receivedEventType == EventType.EventShot)
            {
                string keywordShots = String.Format("{0}{1:D8}_{2}.PNG", _tempDir, _frameCount, _receivedDescription);
                _screenCapture.CaptureScreenToFile(keywordShots, ImageFormat.Png);
            }
            string file = String.Format("{0}{1:D8}.PNG", _tempDir, _frameCount);
            _screenCapture.CaptureScreenToFile(file, ImageFormat.Png);
            _frameCount++;
        }
    }
    #endregion
}
