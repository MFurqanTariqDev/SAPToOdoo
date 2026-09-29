using SAPToOdoo.Application.DTOs.Items;

namespace SAPToOdoo.Application.Interfaces;

public interface ISapItemService
{
    Task<ItemResponse?> GetByItemCodeAsync(string itemCode);
    Task<ItemResponse> CreateAsync(ItemRequest request);
    Task<ItemResponse?> UpdateAsync(string itemCode, ItemUpdateRequest request);
    Task<IReadOnlyList<ItemResponse>> GetLatestAsync(int limit);
}
