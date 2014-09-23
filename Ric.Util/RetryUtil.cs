using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HtmlAgilityPack;
using System.Threading;

namespace Ric.Util
{
    public class RetryUtil
    {
        public delegate void NoArgumentHandler();
        /// <summary>
        /// retry mechanism without argument
        /// </summary>
        /// <param name="retryTimes">try times</param>
        /// <param name="interval">time span</param>
        /// <param name="throwIfFail">throw exception</param>
        /// <param name="function">function name</param>
        public static void Retry(int retryTimes, TimeSpan interval, bool throwIfFail, NoArgumentHandler function)
        {
            if (function == null)
                return;

            for (int i = 0; i < retryTimes; ++i)
            {
                try
                {
                    function();
                    break;
                }
                catch (Exception)
                {
                    if (i == retryTimes - 1)
                    {
                        if (throwIfFail)
                            throw;
                        else
                            break;
                    }
                    else
                    {
                        if (interval != null)
                            Thread.Sleep(interval);
                    }
                }
            }
        }
    }
}
