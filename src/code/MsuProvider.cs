// Copyright (c) Thomas Nieto - All Rights Reserved
// You may use, distribute and modify this code under the
// terms of the MIT license.

using Microsoft.Deployment.Compression.Cab;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Threading;

namespace AnyPackage.Provider.Msu
{
    [PackageProvider("Msu")]
    public sealed class MsuProvider : PackageProvider, IFindPackage, IGetPackage, IInstallPackage
    {
        private readonly static Guid s_id = new Guid("314633fe-c7e9-4eeb-824b-382a8a4e92b8");

        public MsuProvider() : base(s_id) { }

        public void FindPackage(PackageRequest request)
        {
            var file = new CabInfo(request.Name).GetFiles()
                                                .Where(x => Path.GetExtension(x.Name) == ".txt")
                                                .FirstOrDefault();

            if (file is null)
            {
                return;
            }

            string line;
            Dictionary<string, string> metadata = new Dictionary<string, string>();
            using var reader = file.OpenText();

            while ((line = reader.ReadLine()) is not null)
            {
                var values = line.Split('=');
                var key = values[0].Replace(" ", "");

                // Remove quotes around value
                var value = values[1].Substring(1, values[1].Length - 2);

                metadata.Add(key, value);
            }

            if (metadata.ContainsKey("KBArticleNumber"))
            {
                var kb = string.Format("KB{0}", metadata["KBArticleNumber"]);
                request.WritePackage(kb,
                                     new PackageVersion("0"),
                                     metadata["PackageType"],
                                     request.NewSourceInfo(request.Name, request.Name));
            }
        }

        public void GetPackage(PackageRequest request)
        {
            var quickFix = new ManagementObjectSearcher(@"root\cimv2", "select * from Win32_QuickFixEngineering");

            foreach (var hotFix in quickFix.Get())
            {
                if (request.IsMatch((string)hotFix["HotFixID"]))
                {
                    request.WritePackage((string)hotFix["HotFixID"],
                                         new PackageVersion("0"),
                                         (string)hotFix["Description"]);
                }
            }
        }

        public void InstallPackage(PackageRequest request)
        {
            var args = $"{request.Name} /quiet /norestart";
            var psi = new ProcessStartInfo("wusa.exe", args);
            psi.CreateNoWindow = true;
            psi.RedirectStandardOutput = true;
            psi.RedirectStandardError = true;

            using var process = new Process();
            process.StartInfo = psi;
            process.Start();

            while (!process.HasExited)
            {
                if (request.Stopping)
                {
                    process.Kill();
                    throw new OperationCanceledException();
                }

                Thread.Sleep(1000);
            }

            if (process.ExitCode == 3010)
            {
                request.WriteWarning("Reboot required to complete install.");
            }
            else if (process.ExitCode != 0)
            {
                throw new PackageProviderException($"Package '{request.Name}' failed to install.");
            }

            FindPackage(request);
        }
    }
}
