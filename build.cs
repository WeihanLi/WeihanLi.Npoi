// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

#:package WeihanLi.Common

using WeihanLi.Common.Helpers;

var solutionPath = "./WeihanLi.Npoi.slnx";
string[] srcProjects = [ 
    "./src/WeihanLi.Npoi/WeihanLi.Npoi.csproj"
];
string[] testProjects = [ 
    "./test/WeihanLi.Npoi.Test/WeihanLi.Npoi.Test.csproj"
];

await DotNetPackageBuildProcess
    .Create(options => 
    {
        options.SolutionPath = solutionPath;
        options.SrcProjects = srcProjects;
        options.TestProjects = testProjects;
        // options.AdditionalConfigure = c =>
        // {
        //     c.WithTask("test", (b) =>
        //     {
        //         b.WithExecution(async () =>
        //         {
        //             Console.WriteLine("Running custom test task...");
        //             await Task.CompletedTask;
        //         });
        //     });
        // };
    })
    .ExecuteAsync(args);
