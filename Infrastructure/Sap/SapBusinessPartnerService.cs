using System.Runtime.InteropServices;
using SAPToOdoo.Application.DTOs.BusinessPartners;
using SAPToOdoo.Application.Interfaces;
using SAPToOdoo.Common;

namespace SAPToOdoo.Infrastructure.Sap;

public sealed class SapBusinessPartnerService : ISapBusinessPartnerService
{
    private readonly ISapConnectionService _connectionService;
    private readonly ILogger<SapBusinessPartnerService> _logger;

    public SapBusinessPartnerService(ISapConnectionService connectionService, ILogger<SapBusinessPartnerService> logger)
    {
        _connectionService = connectionService;
        _logger = logger;
    }

    public Task<BusinessPartnerResponse?> GetByCardCodeAsync(string cardCode)
    {
        return StaTaskRunner.RunAsync<BusinessPartnerResponse?>(() =>
        {
            using var handle = _connectionService.Connect();
            dynamic? businessPartner = null;
            try
            {
                businessPartner = handle.GetBusinessObject(SapObjectTypes.BusinessPartners, "look up the business partner");

                bool found = businessPartner.GetByKey(cardCode);
                if (!found)
                {
                    return null;
                }

                return new BusinessPartnerResponse
                {
                    CardCode = businessPartner.CardCode,
                    CardName = businessPartner.CardName,
                    Phone = businessPartner.Phone1,
                    Email = businessPartner.EmailAddress
                };
            }
            catch (SapException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new SapException(ErrorCodes.SapOperationFailed,
                    $"Failed to retrieve Business Partner '{cardCode}' from SAP: {ex.Message}", innerException: ex);
            }
            finally
            {
                if (businessPartner is not null && Marshal.IsComObject(businessPartner))
                {
                    Marshal.FinalReleaseComObject(businessPartner);
                }
            }
        });
    }

    public Task<BusinessPartnerResponse?> UpdateAsync(string cardCode, BusinessPartnerUpdateRequest request)
    {
        return StaTaskRunner.RunAsync<BusinessPartnerResponse?>(() =>
        {
            using var handle = _connectionService.Connect();
            dynamic? businessPartner = null;
            try
            {
                businessPartner = handle.GetBusinessObject(SapObjectTypes.BusinessPartners, "look up the business partner");

                bool found = businessPartner.GetByKey(cardCode);
                if (!found)
                {
                    return null;
                }

                businessPartner.CardName = request.CardName;

                if (!string.IsNullOrWhiteSpace(request.Phone))
                {
                    businessPartner.Phone1 = request.Phone;
                }

                if (!string.IsNullOrWhiteSpace(request.Email))
                {
                    businessPartner.EmailAddress = request.Email;
                }

                int result = businessPartner.Update();
                if (result != 0)
                {
                    int errorCode = handle.Company.GetLastErrorCode();
                    string errorMessage = handle.Company.GetLastErrorDescription();
                    throw new SapException(ErrorCodes.SapOperationFailed, errorMessage, errorCode);
                }

                _logger.LogInformation("SAP Business Partner {CardCode} updated", cardCode);

                return new BusinessPartnerResponse
                {
                    CardCode = cardCode,
                    CardName = request.CardName,
                    Phone = request.Phone,
                    Email = request.Email
                };
            }
            catch (SapException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new SapException(ErrorCodes.SapOperationFailed,
                    $"Failed to update Business Partner '{cardCode}' in SAP: {ex.Message}", innerException: ex);
            }
            finally
            {
                if (businessPartner is not null && Marshal.IsComObject(businessPartner))
                {
                    Marshal.FinalReleaseComObject(businessPartner);
                }
            }
        });
    }

    public Task<BusinessPartnerResponse> CreateAsync(BusinessPartnerRequest request)
    {
        return StaTaskRunner.RunAsync(() =>
        {
            using var handle = _connectionService.Connect();
            dynamic? businessPartner = null;
            try
            {
                businessPartner = handle.GetBusinessObject(SapObjectTypes.BusinessPartners, "look up the business partner");

                businessPartner.CardCode = request.CardCode;
                businessPartner.CardName = request.CardName;
                businessPartner.CardType = SapCardTypes.Customer;

                if (!string.IsNullOrWhiteSpace(request.Phone))
                {
                    businessPartner.Phone1 = request.Phone;
                }

                if (!string.IsNullOrWhiteSpace(request.Email))
                {
                    businessPartner.EmailAddress = request.Email;
                }

                int result = businessPartner.Add();
                if (result != 0)
                {
                    int errorCode = handle.Company.GetLastErrorCode();
                    string errorMessage = handle.Company.GetLastErrorDescription();
                    throw new SapException(ErrorCodes.SapOperationFailed, errorMessage, errorCode);
                }

                string newCardCode = handle.Company.GetNewObjectKey();

                _logger.LogInformation("SAP Business Partner {CardCode} created", newCardCode);

                return new BusinessPartnerResponse
                {
                    CardCode = newCardCode,
                    CardName = request.CardName,
                    Phone = request.Phone,
                    Email = request.Email
                };
            }
            catch (SapException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new SapException(ErrorCodes.SapOperationFailed,
                    $"Failed to create Business Partner '{request.CardCode}' in SAP: {ex.Message}", innerException: ex);
            }
            finally
            {
                if (businessPartner is not null && Marshal.IsComObject(businessPartner))
                {
                    Marshal.FinalReleaseComObject(businessPartner);
                }
            }
        });
    }
}
