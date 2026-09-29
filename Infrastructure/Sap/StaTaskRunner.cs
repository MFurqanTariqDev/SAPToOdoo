namespace SAPToOdoo.Infrastructure.Sap;

/// <summary>
/// SAP Business One's DI API is a legacy COM component that must be called from
/// a Single-Threaded Apartment. ASP.NET Core's thread pool threads are MTA, so
/// every DI API call is dispatched onto a dedicated STA thread through this helper.
/// </summary>
internal static class StaTaskRunner
{
    public static Task<T> RunAsync<T>(Func<T> function)
    {
        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

        var thread = new Thread(() =>
        {
            try
            {
                tcs.SetResult(function());
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        })
        {
            IsBackground = true
        };

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        return tcs.Task;
    }
}
