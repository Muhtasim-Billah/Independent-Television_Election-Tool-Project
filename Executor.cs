using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.ComponentModel;
using System.IO;


namespace ElectionResult
{
    class Executor
    {
        private FormElectionResult mainForm;
        private int count;
        public Executor(FormElectionResult mainForm)
        {
            this.mainForm = mainForm;
            this.count = 0;
        }
        public void executorMethod()
        {
            while(true)
            {
                count++;
                this.mainForm.threadSafeCommand(0);
                if (count == 600)
                {
                    count = 0;
                    this.mainForm.threadSafeCommand(1);
                }
              //  Thread.Sleep(new TimeSpan(0, 0, Utility.getIntervalInMinute()));
                Thread.Sleep(new TimeSpan(0, 0, 10));
            }
        
        }
    }
}
