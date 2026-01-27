// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

#:package WeihanLi.Common

using WeihanLi.Common.Helpers;
using WeihanLi.Extensions;

var solutionPath = "./WeihanLi.Npoi.slnx";
string[] srcProjects = [ 
    "./src/WeihanLi.Npoi/WeihanLi.Npoi.csproj"
];
string[] testProjects = [ 
    "./test/WeihanLi.Npoi.Test/WeihanLi.Npoi.Test.csproj"
];
string runFileSamplesDir = "./samples/run-file-samples";

await DotNetPackageBuildProcess
    .Create(options => 
    {
        options.SolutionPath = solutionPath;
        options.SrcProjects = srcProjects;
        options.TestProjects = testProjects;
        options.AdditionalConfigure = c =>
        {
            c.WithTask("build", (b) =>
            {
                b.WithExecution(() =>
                {
                    Console.WriteLine($"Building {solutionPath}...");
                    CommandExecutor.ExecuteCommandAndOutput($"dotnet build {solutionPath}").EnsureSuccessExitCode();
                    foreach (var file in Directory.GetFiles(runFileSamplesDir, "*.cs", SearchOption.AllDirectories))
                    {
                        Console.WriteLine($"Building {file}...");
                        CommandExecutor.ExecuteCommandAndOutput($"dotnet build {file}").EnsureSuccessExitCode();
                    }
                });
            });
        };
    })
    .ExecuteAsync(args);
