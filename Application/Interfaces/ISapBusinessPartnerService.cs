using SAPToOdoo.Application.DTOs.BusinessPartners;

namespace SAPToOdoo.Application.Interfaces;

public interface ISapBusinessPartnerService
{
    Task<BusinessPartnerResponse?> GetByCardCodeAsync(string cardCode);
    Task<BusinessPartnerResponse> CreateAsync(BusinessPartnerRequest request);
    Task<BusinessPartnerResponse?> UpdateAsync(string cardCode, BusinessPartnerUpdateRequest request);
}
