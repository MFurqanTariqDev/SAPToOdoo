using System.Runtime.InteropServices;
using SAPToOdoo.Application.DTOs.Items;
using SAPToOdoo.Application.Interfaces;
using SAPToOdoo.Common;

namespace SAPToOdoo.Infrastructure.Sap;

public sealed class SapItemService : ISapItemService
{
    private readonly ISapConnectionService _connectionService;
    private readonly ILogger<SapItemService> _logger;

    public SapItemService(ISapConnectionService connectionService, ILogger<SapItemService> logger)
    {
        _connectionService = connectionService;
        _logger = logger;
    }

    public Task<ItemResponse?> GetByItemCodeAsync(string itemCode)
    {
        return StaTaskRunner.RunAsync<ItemResponse?>(() =>
        {
            using var handle = _connectionService.Connect();
            dynamic? item = null;
            try
            {
                item = handle.GetBusinessObject(SapObjectTypes.Items, "look up the item");

                bool found = item.GetByKey(itemCode);
                if (!found)
                {
                    return null;
                }

                object rawCreateDate = item.CreateDate;

                return new ItemResponse
                {
                    ItemCode = item.ItemCode,
                    ItemName = item.ItemName,
                    ItemsGroupCode = item.ItemsGroupCode,
                    CreateDate = rawCreateDate is DateTime createDate ? createDate : null,
                    DefaultWarehouse = item.DefaultWarehouse,
                    UnitOfMeasure = item.InventoryUOM
                };
            }
            catch (SapException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new SapException(ErrorCodes.SapOperationFailed,
                    $"Failed to retrieve item '{itemCode}' from SAP: {ex.Message}", innerException: ex);
            }
            finally
            {
                if (item is not null && Marshal.IsComObject(item))
                {
                    Marshal.FinalReleaseComObject(item);
                }
            }
        });
    }

    public Task<ItemResponse> CreateAsync(ItemRequest request)
    {
        return StaTaskRunner.RunAsync(() =>
        {
            using var handle = _connectionService.Connect();
            dynamic? item = null;
            try
            {
                item = handle.GetBusinessObject(SapObjectTypes.Items, "create item");

                item.ItemCode = request.ItemCode;
                item.ItemName = request.ItemName;
                item.ItemsGroupCode = request.ItemsGroupCode!.Value;

                int result = item.Add();
                if (result != 0)
                {
                    int errorCode = handle.Company.GetLastErrorCode();
                    string errorMessage = handle.Company.GetLastErrorDescription();
                    throw new SapException(ErrorCodes.SapOperationFailed, errorMessage, errorCode);
                }

                string newItemCode = handle.Company.GetNewObjectKey();

                _logger.LogInformation("SAP Item {ItemCode} created", newItemCode);

                return new ItemResponse
                {
                    ItemCode = newItemCode,
                    ItemName = request.ItemName,
                    ItemsGroupCode = request.ItemsGroupCode!.Value
                };
            }
            catch (SapException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new SapException(ErrorCodes.SapOperationFailed,
                    $"Failed to create item '{request.ItemCode}' in SAP: {ex.Message}", innerException: ex);
            }
            finally
            {
                if (item is not null && Marshal.IsComObject(item))
                {
                    Marshal.FinalReleaseComObject(item);
                }
            }
        });
    }

    public Task<IReadOnlyList<ItemResponse>> GetLatestAsync(
    int limit,
    string? warehouseCode = null)
    {
        limit = Math.Clamp(limit, 1, 100);

        // Parse multiple comma-separated warehouse codes.
        var warehouseCodes = string.IsNullOrWhiteSpace(warehouseCode)
            ? new List<string>()
            : warehouseCode
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => x.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

        if (warehouseCodes.Any(x => x.Length > 8))
        {
            throw new SapException(
                ErrorCodes.SapOperationFailed,
                "Each warehouse code must be at most 8 characters.");
        }

        // Escape single quotes in warehouse codes.
        string warehouseInSql = string.Join(", ",
            warehouseCodes.Select(x => $"'{x.Replace("'", "''")}'"));

        bool hasWarehouseFilter = warehouseCodes.Count > 0;

        return StaTaskRunner.RunAsync<IReadOnlyList<ItemResponse>>(() =>
        {
            using var handle = _connectionService.Connect();
            dynamic? recordSet = null;

            try
            {
                recordSet = handle.GetBusinessObjectFromCandidates(
                    SapObjectTypes.BoRecordsetCandidates,
                    "query latest items by warehouse");

                // When warehouse codes are supplied, select only items
                // that have a matching OITW record.
                string itemFilter = hasWarehouseFilter
                    ? $@"
WHERE EXISTS
(
    SELECT 1
    FROM OITW X
    WHERE X.ItemCode = OITM.ItemCode
      AND X.WhsCode IN ({warehouseInSql})
)"
                    : string.Empty;

                // INNER JOIN ensures each returned item has a warehouse record.
                // When a filter is supplied, include only those warehouses.
                string warehouseJoin = hasWarehouseFilter
                    ? $@"
INNER JOIN OITW W
    ON W.ItemCode = I.ItemCode
   AND W.WhsCode IN ({warehouseInSql})
INNER JOIN OWHS H
    ON H.WhsCode = W.WhsCode"
                    : @"
INNER JOIN OITW W
    ON W.ItemCode = I.ItemCode
INNER JOIN OWHS H
    ON H.WhsCode = W.WhsCode";

                string query = $@"
SELECT
    I.ItemCode,
    I.ItemName,
    I.ItmsGrpCod,
    I.CreateDate,
    I.DfltWH,
    I.InvntryUom,
    W.WhsCode,
    H.WhsName,
    W.OnHand
FROM
(
    SELECT TOP {limit}
        ItemCode,
        ItemName,
        ItmsGrpCod,
        CreateDate,
        DfltWH,
        InvntryUom
    FROM OITM
    {itemFilter}
    ORDER BY CreateDate DESC, ItemCode DESC
) I
{warehouseJoin}
ORDER BY
    I.CreateDate DESC,
    I.ItemCode DESC,
    W.WhsCode";

                recordSet.DoQuery(query);

                var items = new List<ItemResponse>();
                var byCode = new Dictionary<string, ItemResponse>(
                    StringComparer.OrdinalIgnoreCase);

                while (!(bool)recordSet.EoF)
                {
                    string itemCode =
                        Convert.ToString(
                            recordSet.Fields.Item("ItemCode").Value)
                        ?? string.Empty;

                    if (!byCode.TryGetValue(itemCode, out var item))
                    {
                        object rawCreateDate =
                            recordSet.Fields.Item("CreateDate").Value;

                        string? defaultWarehouse =
                            Convert.ToString(
                                recordSet.Fields.Item("DfltWH").Value);

                        string? uom =
                            Convert.ToString(
                                recordSet.Fields.Item("InvntryUom").Value);

                        item = new ItemResponse
                        {
                            ItemCode = itemCode,

                            ItemName =
                                Convert.ToString(
                                    recordSet.Fields.Item("ItemName").Value)
                                ?? string.Empty,

                            ItemsGroupCode = Convert.ToInt32(
                                recordSet.Fields.Item("ItmsGrpCod").Value),

                            CreateDate = rawCreateDate is DateTime createDate
                                ? createDate
                                : null,

                            DefaultWarehouse =
                                string.IsNullOrWhiteSpace(defaultWarehouse)
                                    ? null
                                    : defaultWarehouse,

                            UnitOfMeasure =
                                string.IsNullOrWhiteSpace(uom)
                                    ? null
                                    : uom
                        };

                        byCode[itemCode] = item;
                        items.Add(item);
                    }

                    string? resultWarehouseCode =
                        Convert.ToString(
                            recordSet.Fields.Item("WhsCode").Value);

                    if (!string.IsNullOrWhiteSpace(resultWarehouseCode))
                    {
                        object rawOnHand =
                            recordSet.Fields.Item("OnHand").Value;

                        item.Warehouses.Add(new WarehouseStockResponse
                        {
                            WarehouseCode = resultWarehouseCode,

                            WarehouseName =
                                Convert.ToString(
                                    recordSet.Fields.Item("WhsName").Value)
                                ?? string.Empty,

                            OnHand = rawOnHand is null || rawOnHand is DBNull
                                ? 0m
                                : Convert.ToDecimal(rawOnHand)
                        });
                    }

                    recordSet.MoveNext();
                }

                return (IReadOnlyList<ItemResponse>)items;
            }
            catch (SapException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new SapException(
                    ErrorCodes.SapOperationFailed,
                    $"Failed to retrieve latest items from SAP: {ex.Message}",
                    innerException: ex);
            }
            finally
            {
                if (recordSet is not null && Marshal.IsComObject(recordSet))
                {
                    Marshal.FinalReleaseComObject(recordSet);
                }
            }
        });
    }

    public Task<ItemResponse?> UpdateAsync(string itemCode, ItemUpdateRequest request)
    {
        return StaTaskRunner.RunAsync<ItemResponse?>(() =>
        {
            using var handle = _connectionService.Connect();
            dynamic? item = null;
            try
            {
                item = handle.GetBusinessObject(SapObjectTypes.Items, "look up the item");

                bool found = item.GetByKey(itemCode);
                if (!found)
                {
                    return null;
                }

                item.ItemName = request.ItemName;

                int result = item.Update();
                if (result != 0)
                {
                    int errorCode = handle.Company.GetLastErrorCode();
                    string errorMessage = handle.Company.GetLastErrorDescription();
                    throw new SapException(ErrorCodes.SapOperationFailed, errorMessage, errorCode);
                }

                _logger.LogInformation("SAP Item {ItemCode} updated", itemCode);

                return new ItemResponse
                {
                    ItemCode = itemCode,
                    ItemName = request.ItemName,
                    ItemsGroupCode = (int)item.ItemsGroupCode
                };
            }
            catch (SapException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new SapException(ErrorCodes.SapOperationFailed,
                    $"Failed to update item '{itemCode}' in SAP: {ex.Message}", innerException: ex);
            }
            finally
            {
                if (item is not null && Marshal.IsComObject(item))
                {
                    Marshal.FinalReleaseComObject(item);
                }
            }
        });
    }

    public Task<IReadOnlyList<ItemResponse>> GetByWarehouseAsync(IEnumerable<string> warehouseCodes, int limit = 20, int offset = 0)
    {
        limit = Math.Clamp(limit, 1, 100);
        offset = Math.Max(offset, 0);

        // Accept repeated parameters and comma-separated values, in any mix.
        var codes = (warehouseCodes ?? Enumerable.Empty<string>())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .SelectMany(x => x.Split(',', StringSplitOptions.RemoveEmptyEntries))
            .Select(x => x.Trim())
            .Where(x => x.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (codes.Count == 0)
        {
            throw new SapException(
                ErrorCodes.SapOperationFailed,
                "At least one warehouseCode is required.");
        }

        if (codes.Count > 50)
        {
            throw new SapException(
                ErrorCodes.SapOperationFailed,
                "At most 50 warehouse codes can be requested at once.");
        }

        if (codes.Any(x => x.Length > 8))
        {
            throw new SapException(
                ErrorCodes.SapOperationFailed,
                "Each warehouse code must be at most 8 characters.");
        }

        string warehouseInSql = string.Join(", ",
            codes.Select(x => $"'{x.Replace("'", "''")}'"));

        return StaTaskRunner.RunAsync<IReadOnlyList<ItemResponse>>(() =>
        {
            using var handle = _connectionService.Connect();
            dynamic? recordSet = null;

            try
            {
                recordSet = handle.GetBusinessObjectFromCandidates(
                    SapObjectTypes.BoRecordsetCandidates,
                    "query items by warehouse");

                // Step 1 (inner query): pick one page of items that have a stock
                //   record in the requested warehouse(s).
                // Step 2 (joins): attach only the requested warehouses to each item.
                string query = $@"
SELECT
    I.ItemCode,
    I.ItemName,
    I.ItmsGrpCod,
    I.CreateDate,
    I.DfltWH,
    I.InvntryUom,
    W.WhsCode,
    H.WhsName,
    W.OnHand
FROM
(
    SELECT
        O.ItemCode,
        O.ItemName,
        O.ItmsGrpCod,
        O.CreateDate,
        O.DfltWH,
        O.InvntryUom
    FROM OITM O WITH (NOLOCK)
    WHERE EXISTS
    (
        SELECT 1
        FROM OITW X WITH (NOLOCK)
        WHERE X.ItemCode = O.ItemCode
          AND X.WhsCode IN ({warehouseInSql})
    )
    ORDER BY O.CreateDate DESC, O.ItemCode DESC
    OFFSET {offset} ROWS FETCH NEXT {limit} ROWS ONLY
) I
INNER JOIN OITW W WITH (NOLOCK)
    ON W.ItemCode = I.ItemCode
   AND W.WhsCode IN ({warehouseInSql})
INNER JOIN OWHS H WITH (NOLOCK)
    ON H.WhsCode = W.WhsCode
ORDER BY
    I.CreateDate DESC,
    I.ItemCode DESC,
    W.WhsCode";

                recordSet.DoQuery(query);

                var items = new List<ItemResponse>();
                var byCode = new Dictionary<string, ItemResponse>(StringComparer.OrdinalIgnoreCase);

                while (!(bool)recordSet.EoF)
                {
                    string itemCode =
                        Convert.ToString(recordSet.Fields.Item("ItemCode").Value) ?? string.Empty;

                    if (!byCode.TryGetValue(itemCode, out var item))
                    {
                        object rawCreateDate = recordSet.Fields.Item("CreateDate").Value;
                        string? dfltWh = Convert.ToString(recordSet.Fields.Item("DfltWH").Value);
                        string? uom = Convert.ToString(recordSet.Fields.Item("InvntryUom").Value);

                        item = new ItemResponse
                        {
                            ItemCode = itemCode,
                            ItemName = Convert.ToString(recordSet.Fields.Item("ItemName").Value) ?? string.Empty,
                            ItemsGroupCode = Convert.ToInt32(recordSet.Fields.Item("ItmsGrpCod").Value),
                            CreateDate = rawCreateDate is DateTime createDate ? createDate : null,
                            DefaultWarehouse = string.IsNullOrWhiteSpace(dfltWh) ? null : dfltWh,
                            UnitOfMeasure = string.IsNullOrWhiteSpace(uom) ? null : uom
                        };

                        byCode[itemCode] = item;
                        items.Add(item);
                    }

                    object rawOnHand = recordSet.Fields.Item("OnHand").Value;

                    item.Warehouses.Add(new WarehouseStockResponse
                    {
                        WarehouseCode = Convert.ToString(recordSet.Fields.Item("WhsCode").Value) ?? string.Empty,
                        WarehouseName = Convert.ToString(recordSet.Fields.Item("WhsName").Value) ?? string.Empty,
                        OnHand = rawOnHand is null || rawOnHand is DBNull ? 0m : Convert.ToDecimal(rawOnHand)
                    });

                    recordSet.MoveNext();
                }

                return (IReadOnlyList<ItemResponse>)items;
            }
            catch (SapException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new SapException(
                    ErrorCodes.SapOperationFailed,
                    $"Failed to retrieve items by warehouse from SAP: {ex.Message}",
                    innerException: ex);
            }
            finally
            {
                if (recordSet is not null && Marshal.IsComObject(recordSet))
                {
                    Marshal.FinalReleaseComObject(recordSet);
                }
            }
        });
    }
}
