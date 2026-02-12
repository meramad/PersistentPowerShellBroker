namespace PersistentPowerShellBroker.Native;

internal sealed class ExcelApplicationSession
{
    public ExcelApplicationSession(object application, bool createdByBroker)
    {
        Application = application;
        CreatedByBroker = createdByBroker;
    }

    public object Application { get; }
    public bool CreatedByBroker { get; }

    public void EnsureVisible(bool forceVisible, object? workbook = null)
    {
        if (!forceVisible)
        {
            return;
        }

        dynamic app = Application;
        app.Visible = true;

        if (workbook is not null)
        {
            dynamic wb = workbook;
            dynamic windows = wb.Windows;
            try
            {
                var windowCount = Convert.ToInt32(windows.Count);
                if (windowCount > 0)
                {
                    dynamic firstWindow = windows.Item(1);
                    try
                    {
                        firstWindow.Visible = true;
                        wb.Activate();
                    }
                    finally
                    {
                        ExcelCommandSupport.SafeReleaseComObject(firstWindow);
                    }
                }
            }
            finally
            {
                ExcelCommandSupport.SafeReleaseComObject(windows);
            }
        }

        var visible = Convert.ToBoolean(app.Visible);
        if (!visible)
        {
            throw new InvalidOperationException("Excel application remained hidden after visibility enforcement.");
        }
    }

    public void SetDisplayAlerts(bool enabled)
    {
        dynamic app = Application;
        app.DisplayAlerts = enabled;
    }

    public int GetOpenWorkbookCount()
    {
        dynamic app = Application;
        dynamic workbooks = app.Workbooks;
        try
        {
            return Convert.ToInt32(workbooks.Count);
        }
        finally
        {
            ExcelCommandSupport.SafeReleaseComObject(workbooks);
        }
    }

    public void Quit()
    {
        dynamic app = Application;
        app.Quit();
    }

    public void Release()
    {
        ExcelCommandSupport.SafeReleaseComObject(Application);
    }
}
