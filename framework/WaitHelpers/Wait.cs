using System;
using System.Threading;
using System.Threading.Tasks;
using Framework.Exceptions;
using Framework.TimeoutManagement;
using NLog;
using Polly;
using Polly.Retry;

namespace Framework.WaitHelpers
{
  /// <summary>
  /// Wait class
  /// </summary>
  public static class Wait
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();


        public static CancellationTokenSource RepeatTask(int pauseSeconds, Action myCancellationAction)
        {
            Log.Info("RepeatTask started at : {0}", DateTime.Now);
            var tokenSource = new CancellationTokenSource();
            var cancellationToken = tokenSource.Token;

            Task.Run(() =>
            {
                while (true)
                {
                    Log.Info("Task Delay started at : {0}", DateTime.Now);
                    try
                    {
                        Task.Delay(pauseSeconds*1000, cancellationToken).Wait();
                        Log.Info("Task Delay completed at : {0}", DateTime.Now);
                    }
                    catch (Exception e)
                    {
                        Log.Info("Task Delay terminated at : {0}", DateTime.Now);
                        Log.Info(e.Message);
                    }

                    if (cancellationToken.IsCancellationRequested)
                    {
                        Log.Info($"Cancelling at {DateTime.Now}");
                        break;
                    }

                    Log.Info("Executing my action");
                    myCancellationAction();
                    Log.Info("Executing my action completed");
                }
            }, cancellationToken);

            return tokenSource;
        }

        /// <summary>
        /// Wait for expression to become true, sleepBetween tries in ms
        /// </summary>
        /// <param name="func"></param>
        /// <param name="timeout"></param>
        /// <param name="msPause"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public static bool Until(Func<bool> func, int timeout, int msPause = 300, string message = "")
        {
            try
            {
                UntilOrThrow(func, timeout, msPause,message);
                return true;
            }
            catch
            {
                return false;
            }
        }


        public static T Until<T>(Func<T> func, int timeout, int msPause = 300, string message = "") where T : class
        {
            try
            {
                return UntilOrThrow(func, timeout, msPause, message);
            }
            catch
            {
                return null;
            }
        }


        private static RetryPolicy GetRetryPolicy(int ms, string message)
        {
            var retries = 0;
            var waitAndRetryPolicy = Policy
                .Handle<Exception>()
                .WaitAndRetryForever(
                    (attempt, context) => TimeSpan.FromMilliseconds(ms),
                    (exception, calculatedWaitDuration, context) =>
                    {
                        retries++;
                        Log.Info($"Wait - Retry = {retries} for '{message}'");
                        context["Err"] = exception;
                    });
            return waitAndRetryPolicy;
        }

        /// <summary>
        /// Wait for expression to become true, sleepBetween tries in ms
        /// </summary>
        /// <param name="func"></param>
        /// <param name="timeout"></param>
        /// <param name="msPause"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public static void UntilOrThrow(Func<bool> func, int timeout=20, int msPause = 500, string message = null)
            {
                Log.Debug($"Message : '{message}'");
                Log.Debug($"Timeout : {timeout} secs, overrides any underlying timeouts.");
                var timeoutPolicy = FaultHandling.TimeoutPolicy(timeout);

                var waitAndRetryPolicy = GetRetryPolicy(msPause,message);
                var wrap = Policy.Wrap(timeoutPolicy, waitAndRetryPolicy);
                var result = wrap.ExecuteAndCapture(() =>
                {
                    if (!func()) throw new WaitUntilException($"Expression failed... : '{message}'");
                });

                if (result.Outcome == OutcomeType.Successful)
                {
                    Log.Debug($"Message : '{message}' completed successfully...");
                    return;
                }

                Log.Warn($"Message : '{message}' failed, timeout...");
                Log.Warn($"{result.FinalException.Message}");
                var exception = new WaitUntilException(message,result.FinalException);
                throw exception;
            }

        /// <summary>
        /// Wait for expression to become true, sleepBetween tries in ms
        /// </summary>
        /// <param name="func"></param>
        /// <param name="timeout"></param>
        /// <param name="msPause"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public static T UntilOrThrow<T>(Func<T> func, int timeout = 20, int msPause = 500, string message = null)
        {
            Log.Debug($"Message : '{message}'");
            Log.Debug($"Timeout : {timeout} secs, overrides any underlying timeouts.");
            var timeoutPolicy = FaultHandling.TimeoutPolicy(timeout);

            var waitAndRetryPolicy = GetRetryPolicy(msPause, message);
            var wrap = Policy.Wrap(timeoutPolicy, waitAndRetryPolicy);
            var res = default(T);
            var result = wrap.ExecuteAndCapture(() =>
            {
                res = func();
                if (res == null) throw new WaitUntilException($"Expression failed... : '{message}'");
            });

            if (result.Outcome == OutcomeType.Successful)
            {
                Log.Debug($"Message : '{message}' completed successfully...");
                return res;
            }

            Log.Warn($"Message : '{message}' failed, timeout...");
            Log.Warn($"{result.FinalException.Message}");
            var exception = new WaitUntilException(message, result.FinalException);
            throw exception;
        }
    }
}
