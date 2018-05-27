using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;

namespace Wox.Plugin.Services
{
    public class Plugin : IPlugin
    {
        public void Init(PluginInitContext context)
        {
        }

        public List<Result> Query(Query query)
        {
            var search = query.Search;
            var fuzzyMatcher = FuzzyMatcher.Create(search);

            var services = ServiceController
                .GetServices()
                .Where(CanStartOrStopService)
                .Select(s => new {Service = s, Score = CalculateScore(s, fuzzyMatcher)})
                .Where(s => s.Score > 0 || string.IsNullOrEmpty(search))
                .ToList();

            var results = services.Select(s => new Result
            {
                Title = s.Service.ServiceName,
                SubTitle = s.Service.DisplayName,
                Score = s.Score,
                IcoPath = s.Service.Status == ServiceControllerStatus.Running ? @"Images\green.png" : @"Images\red.png",
                Action = context =>
                {
                    var action = s.Service.Status == ServiceControllerStatus.Running ? "stop" : "start";

                    var info = new ProcessStartInfo
                    {
                        FileName = "net.exe",
                        UseShellExecute = true,
                        Verb = "runas", // Provides Run as Administrator
                        Arguments = $"{action} {s.Service.ServiceName}"
                    };

                    try
                    {
                        Process.Start(info);
                    }
                    catch (Win32Exception ex)
                    {
                        const int cancelledErrorCode = 1223;
                        if (ex.NativeErrorCode == cancelledErrorCode)
                            return false;
                        throw;
                    }

                    return true;
                }
            }).ToList();

            return results;
        }

        private static bool CanStartOrStopService(ServiceController service)
        {
            var canStart = service.Status == ServiceControllerStatus.Stopped;
            var canStop = service.Status == ServiceControllerStatus.Running && service.CanStop;
            return canStart || canStop;
        }

        private static int CalculateScore(ServiceController controller, FuzzyMatcher matcher)
        {
            return Math.Max(matcher.Evaluate(controller.ServiceName).Score, matcher.Evaluate(controller.DisplayName).Score);
        }
    }
}
