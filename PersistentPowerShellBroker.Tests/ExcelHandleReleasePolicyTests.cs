using System.Text.Json;
using System.Management.Automation.Runspaces;
using PersistentPowerShellBroker.Native;

namespace PersistentPowerShellBroker.Tests;

public sealed class ExcelHandleReleasePolicyTests
{
    [Fact]
    public async Task ReleaseHandle_DoesNotQuit_WhenAppIsNotBrokerOwned()
    {
        using var runspace = RunspaceFactory.CreateRunspace();
        runspace.Open();

        var app = new FakeExcelApplication(workbookCount: 1);
        var workbook = new FakeWorkbook("C:/Temp/Test.xlsx");
        var handleName = "excelHandle_test_not_owned";
        var bundle = ExcelCommandSupport.BuildBundle(
            app,
            workbook,
            requestedTarget: workbook.FullName,
            workbookFullName: workbook.FullName,
            isReadOnly: false,
            attachedExisting: true,
            openedWorkbook: false,
            createdApplicationByBroker: false);
        ExcelCommandSupport.SetGlobalVariable(runspace, handleName, bundle);
        ExcelHandleRegistry.Register(new ExcelHandleMetadata(
            VariableName: handleName,
            RequestedTarget: workbook.FullName,
            WorkbookFullName: workbook.FullName,
            AttachedExisting: true,
            OpenedWorkbook: false,
            IsReadOnly: false,
            CreatedApplicationByBroker: false,
            CreatedUtc: DateTime.UtcNow));

        var cmd = new BrokerExcelReleaseHandleCommand();
        var context = new BrokerContext
        {
            PipeName = "test",
            ProcessId = Environment.ProcessId,
            StartedAtUtc = DateTimeOffset.UtcNow,
            RequestStop = static () => { }
        };

        var args = JsonDocument.Parse("""
        {
          "psVariableName":"excelHandle_test_not_owned",
          "closeWorkbook":false,
          "quitExcel":true,
          "onlyIfNoOtherWorkbooks":true,
          "displayAlerts":false
        }
        """).RootElement;

        var result = await cmd.ExecuteAsync(args, context, runspace, CancellationToken.None);

        Assert.True(result.Success, $"error={result.Error} stdout={result.Stdout}");
        Assert.False(app.QuitCalled);

        using var payload = JsonDocument.Parse(result.Stdout);
        var root = payload.RootElement;
        Assert.True(root.GetProperty("quitSkipped").GetBoolean());
        Assert.Equal("NotBrokerOwnedApplication", root.GetProperty("quitSkipReason").GetString());
    }

    private sealed class FakeExcelApplication
    {
        public FakeExcelApplication(int workbookCount)
        {
            Workbooks = new FakeWorkbooks(workbookCount);
        }

        public bool Visible { get; set; }
        public bool UserControl { get; set; }
        public bool DisplayAlerts { get; set; } = true;
        public FakeWorkbooks Workbooks { get; }
        public bool QuitCalled { get; private set; }

        public void Quit()
        {
            QuitCalled = true;
        }
    }

    private sealed class FakeWorkbooks
    {
        public FakeWorkbooks(int count)
        {
            Count = count;
        }

        public int Count { get; }
    }

    private sealed class FakeWorkbook
    {
        public FakeWorkbook(string fullName)
        {
            FullName = fullName;
            Windows = new FakeWindows();
        }

        public string FullName { get; }
        public FakeWindows Windows { get; }

        public void Close(object? saveChanges)
        {
        }

        public void Activate()
        {
        }
    }

    private sealed class FakeWindows
    {
        public int Count => 0;
    }
}
