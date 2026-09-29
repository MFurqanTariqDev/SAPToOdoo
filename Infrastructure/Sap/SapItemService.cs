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
                item = handle.GetBusinessObject(SapObjectTypes.Items, "look up the item");

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

    public Task<IReadOnlyList<ItemResponse>> GetLatestAsync(int limit)
    {
        return StaTaskRunner.RunAsync<IReadOnlyList<ItemResponse>>(() =>
        {
            using var handle = _connectionService.Connect();
            dynamic? recordSet = null;
            try
            {
                recordSet = handle.GetBusinessObjectFromCandidates(SapObjectTypes.BoRecordsetCandidates, "query the latest items");

                // OITM has no CreateTime column (unlike some transactional tables), so
                // CreateDate is day-granularity only; ItemCode is a deterministic
                // tiebreaker for same-day rows, not a substitute ordering criterion.
                string query = $"SELECT TOP {limit} ItemCode, ItemName, ItmsGrpCod, CreateDate, DfltWH, InvntryUom " +
                                "FROM OITM ORDER BY CreateDate DESC, ItemCode DESC";

                recordSet.DoQuery(query);

                var results = new List<ItemResponse>();
                if (!(bool)recordSet.EoF)
                {
                    recordSet.MoveFirst();
                    while (!(bool)recordSet.EoF)
                    {
                        object rawCreateDate = recordSet.Fields.Item("CreateDate").Value;
                        object rawWarehouse = recordSet.Fields.Item("DfltWH").Value;
                        object rawUom = recordSet.Fields.Item("InvntryUom").Value;

                        results.Add(new ItemResponse
                        {
                            ItemCode = recordSet.Fields.Item("ItemCode").Value.ToString(),
                            ItemName = recordSet.Fields.Item("ItemName").Value.ToString(),
                            ItemsGroupCode = Convert.ToInt32(recordSet.Fields.Item("ItmsGrpCod").Value),
                            CreateDate = rawCreateDate is DateTime createDate ? createDate : null,
                            DefaultWarehouse = string.IsNullOrWhiteSpace(rawWarehouse?.ToString()) ? null : rawWarehouse.ToString(),
                            UnitOfMeasure = string.IsNullOrWhiteSpace(rawUom?.ToString()) ? null : rawUom.ToString()
                        });

                        recordSet.MoveNext();
                    }
                }

                return (IReadOnlyList<ItemResponse>)results;
            }
            catch (SapException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new SapException(ErrorCodes.SapOperationFailed,
                    $"Failed to retrieve latest items from SAP: {ex.Message}", innerException: ex);
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
}
