// Copyright (c) Weihan Li. All rights reserved.
// Licensed under the Apache license.

#:package WeihanLi.Core

using WeihanLi.Common.Helpers;

var solutionPath = "./WeihanLi.Npoi.slnx";
string[] srcProjects = [ 
    "./src/WeihanLi.Npoi/WeihanLi.Npoi.csproj"
];
string[] testProjects = [ 
    "./test/WeihanLi.Npoi.Test/WeihanLi.Npoi.Test.csproj"
];
string[] runFileSamplesFolders = [
    "./samples/run-file-samples"
];

await DotNetPackageBuildProcess
    .Create(options => 
    {
        options.SolutionPath = solutionPath;
        options.SrcProjects = srcProjects;
        options.TestProjects = testProjects;
        options.RunFileSampleFolders = runFileSamplesFolders;
    })
    .ExecuteAsync(args);
