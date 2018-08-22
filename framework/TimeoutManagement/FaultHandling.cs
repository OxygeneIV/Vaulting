using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NLog;
using Polly;
using Polly.Timeout;

namespace Framework.TimeoutManagement
{
    public class FaultHandling
    {

        protected static Logger Log= LogManager.GetCurrentClassLogger();

        public static TimeoutPolicy TimeoutPolicyAsync(int seconds = 20)
        {
            Log.Debug($"Creating Timeout policy Async, timeout = {seconds} seconds");

            return Policy.TimeoutAsync(seconds, TimeoutStrategy.Pessimistic);

        }

        public static TimeoutPolicy TimeoutPolicy(int seconds = 20)
        {
            Log.Debug($"Creating Timeout policy, timeout = {seconds} seconds");


            var timeoutPolicy = Policy.Timeout(TimeSpan.FromSeconds(seconds), TimeoutStrategy.Pessimistic,
                    (context, timespan, task) =>
                    {

                        // Todo Try a cleanup using ConinueWith method

                        
                        task.ContinueWith(t => { // ContinueWith important!: the abandoned task may very well still be executing, when the caller times out on waiting for it! 
                            Log.Warn("Timeout occured...");

                            if (t.IsFaulted)
                            {

                                Log.Info($"TIMEOUT: {context.PolicyKey} at {context.ExecutionKey}: execution timed out after {timespan.TotalSeconds} seconds, eventually terminated with: {t.Exception}.");
                            }
                            else if (t.IsCanceled)
                            {

                                // (If the executed delegates do not honour cancellation, this IsCanceled branch may never be hit.  It can be good practice however to include, in case a Policy configured with TimeoutStrategy.Pessimistic is used to execute a delegate honouring cancellation.)  
                                Log.Info($"TIMEOUT: {context.PolicyKey} at {context.ExecutionKey}: execution timed out after {timespan.TotalSeconds} seconds, task cancelled.");
                            }
                            else
                            {
                                // extra logic (if desired) for tasks which complete, despite the caller having 'walked away' earlier due to timeout.
                                // Log.Info("TIMEOUT: Place for extra logic");
                            }

                            // Additionally, clean up any resources ...


                        });

                        
                        // Todo commented this to try and see what the error looks like
                    });
            return timeoutPolicy;
        }


        public static TimeoutPolicy TimeoutPolicyOrg(int seconds = 20)
        {



            //Log.Debug($"Creating Timeout policy, timeout = {seconds} seconds");
            var timeoutPolicy = Policy.Timeout(TimeSpan.FromSeconds(seconds), TimeoutStrategy.Pessimistic,
                    (context, timespan, task) =>
                    {

                    // Todo Try a cleanup using ConinueWith method

                       
                    task.ContinueWith(t => { // ContinueWith important!: the abandoned task may very well still be executing, when the caller times out on waiting for it! 

                        //TimeoutLogger.Info("Timeout occured, terminating task...");
                        //task.Dispose();
                        //TimeoutLogger.Info("Task Disposed...");


                        if (t.IsFaulted)
                            {
                            //Log.Error($"{context.PolicyKey} at {context.ExecutionKey}: execution timed out after {timespan.TotalSeconds} seconds, eventually terminated with: {t.Exception}.");
                        }
                            else if (t.IsCanceled)
                            {
                            // (If the executed delegates do not honour cancellation, this IsCanceled branch may never be hit.  It can be good practice however to include, in case a Policy configured with TimeoutStrategy.Pessimistic is used to execute a delegate honouring cancellation.)  
                            //logger.Error($"{context.PolicyKey} at {context.ExecutionKey}: execution timed out after {timespan.TotalSeconds} seconds, task cancelled.");
                        }
                            else
                            {

                            // extra logic (if desired) for tasks which complete, despite the caller having 'walked away' earlier due to timeout.
                        }

                        // Additionally, clean up any resources ...
                        

                    });


                    // Todo commented this to try and see what the error looks like
                    //Exception err = context.ContainsKey("Err") ? (Exception)context["Err"] : null;

                    //if (err == null) return;
                    //var mess = $"Timeout ({seconds} sec)";
                    //var newErr = new TimeoutRejectedException(mess,err);
                    //throw newErr;
                });

            return timeoutPolicy;
        }
    }
}
