var target = CommandLineParser.Val("target", "Default", args);
var apiKey = CommandLineParser.Val("apiKey", args);
var stable = CommandLineParser.BooleanVal("stable", args);
var noPush = CommandLineParser.BooleanVal("noPush", args);
var branchName = EnvHelper.Val("BUILD_SOURCEBRANCHNAME", "local");

var solutionPath = "./WeihanLi.Npoi.slnx";
string[] srcProjects = [ 
    "./src/WeihanLi.Npoi/WeihanLi.Npoi.csproj"
];
string[] testProjects = [ "./test/WeihanLi.Npoi.Test/WeihanLi.Npoi.Test.csproj" ];

await new BuildProcessBuilder()
    .WithSetup(() =>
    {
        // cleanup artifacts
        if (Directory.Exists("./artifacts/packages"))
            Directory.Delete("./artifacts/packages", true);

        // args
        Console.WriteLine("Arguments");
        Console.WriteLine($"    {args.StringJoin(" ")}");
    })
    .WithTaskExecuting(task => Console.WriteLine($@"===== Task {task.Name} {task.Description} executing ======"))
    .WithTaskExecuted(task => Console.WriteLine($@"===== Task {task.Name} {task.Description} executed ======"))
    .WithTask("hello", b => b.WithExecution(() => Console.WriteLine("Hello dotnet-exec build")))
    .WithTask("build", b =>
    {
        b.WithDescription("dotnet build")
            .WithExecution(() => ExecuteCommandAsync($"dotnet build {solutionPath}"))
            ;
    })
    .WithTask("test", b =>
    {
        b.WithDescription("dotnet test")
            .WithDependency("build")
            .WithExecution(async () =>
            {
                foreach (var project in testProjects)
                {
                    await ExecuteCommandAsync($"dotnet test --collect:\"XPlat Code Coverage;Format=cobertura,opencover;ExcludeByAttribute=ExcludeFromCodeCoverage,Obsolete,GeneratedCode,CompilerGeneratedAttribute\" {project}");
                }
            })
            ;
    })
    .WithTask("pack", b => b.WithDescription("dotnet pack")
        .WithDependency("test")
        .WithExecution(async () =>
        {
            if (stable || branchName == "master")
            {
                foreach (var project in srcProjects)
                {
                    await ExecuteCommandAsync($"dotnet pack {project} -o ./artifacts/packages");
                }
            }
            else
            {
                var suffix = $"preview-{DateTime.UtcNow:yyyyMMdd-HHmmss}";
                foreach (var project in srcProjects)
                {
                    await ExecuteCommandAsync($"dotnet pack {project} -o ./artifacts/packages --version-suffix {suffix}");
                }
            }

            if (noPush)
            {
                Console.WriteLine("Skip push there's noPush specified");
                return;
            }
            
            if (string.IsNullOrEmpty(apiKey))
            {
                // try to get apiKey from environment variable
                apiKey = Environment.GetEnvironmentVariable("NuGet__ApiKey");
                
                if (string.IsNullOrEmpty(apiKey))
                {
                    Console.WriteLine("Skip push since there's no apiKey found");
                    return;
                }
            }

            if (branchName != "master" && branchName != "preview")
            {
                Console.WriteLine($"Skip push since branch name {branchName} not support push packages");
                return;
            }

            // push nuget packages
            var source = Environment.GetEnvironmentVariable("NuGet__SourceUrl");
            var sourceConfig = string.IsNullOrEmpty(source) ? "" : $"-s {source}";
            foreach (var file in Directory.GetFiles("./artifacts/packages/", "*.nupkg"))
            {
                await ExecuteCommandAsync($"dotnet nuget push {file} -k {apiKey} --skip-duplicate {sourceConfig}");
            }
        }))
    .WithTask("Default", b => b.WithDependency("hello").WithDependency("pack"))
    .Build()
    .ExecuteAsync(target, ApplicationHelper.ExitToken);

async Task ExecuteCommandAsync(string commandText)
{
    var commandTextWithReplacements = commandText;
    Console.WriteLine($"Executing command: \n    {commandTextWithReplacements}");
    Console.WriteLine();
    var result = await CommandExecutor.ExecuteCommandAndOutputAsync(commandText);
    result.EnsureSuccessExitCode();
    Console.WriteLine();
}
