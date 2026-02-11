using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace CSharpFlexGrid
{
    public class JsonImportResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public BookingOrder BookingOrder { get; set; }
    }

    public static class JsonImportHelper
    {
        public static JsonImportResult ValidateAndParse(string jsonContent)
        {
            if (string.IsNullOrWhiteSpace(jsonContent))
            {
                return new JsonImportResult
                {
                    Success = false,
                    ErrorMessage = "JSON content is empty."
                };
            }

            BookingOrder order;
            try
            {
                order = JsonConvert.DeserializeObject<BookingOrder>(jsonContent);
            }
            catch (JsonException ex)
            {
                return new JsonImportResult
                {
                    Success = false,
                    ErrorMessage = $"Invalid JSON syntax: {ex.Message}"
                };
            }

            if (order == null)
            {
                return new JsonImportResult
                {
                    Success = false,
                    ErrorMessage = "Failed to parse JSON into a BookingOrder."
                };
            }

            // Validate required fields
            if (string.IsNullOrWhiteSpace(order.Agency))
            {
                return new JsonImportResult
                {
                    Success = false,
                    ErrorMessage = "Missing required field: 'agency'."
                };
            }

            if (string.IsNullOrWhiteSpace(order.Advertiser))
            {
                return new JsonImportResult
                {
                    Success = false,
                    ErrorMessage = "Missing required field: 'advertiser'."
                };
            }

            if (string.IsNullOrWhiteSpace(order.Product))
            {
                return new JsonImportResult
                {
                    Success = false,
                    ErrorMessage = "Missing required field: 'product'."
                };
            }

            // Validate campaign period
            if (order.CampaignPeriod == null)
            {
                return new JsonImportResult
                {
                    Success = false,
                    ErrorMessage = "Missing required field: 'campaign_period'."
                };
            }

            if (!DateTime.TryParse(order.CampaignPeriod.StartDate, out DateTime startDate))
            {
                return new JsonImportResult
                {
                    Success = false,
                    ErrorMessage = "Invalid or missing 'start_date' in campaign_period."
                };
            }

            if (!DateTime.TryParse(order.CampaignPeriod.EndDate, out DateTime endDate))
            {
                return new JsonImportResult
                {
                    Success = false,
                    ErrorMessage = "Invalid or missing 'end_date' in campaign_period."
                };
            }

            if (endDate < startDate)
            {
                return new JsonImportResult
                {
                    Success = false,
                    ErrorMessage = "Campaign 'end_date' must be on or after 'start_date'."
                };
            }

            // Validate spots list exists
            if (order.Spots == null)
            {
                order.Spots = new List<Spot>();
            }

            // Validate each spot
            for (int i = 0; i < order.Spots.Count; i++)
            {
                var spot = order.Spots[i];
                if (string.IsNullOrWhiteSpace(spot.ProgrammeName))
                {
                    return new JsonImportResult
                    {
                        Success = false,
                        ErrorMessage = $"Spot #{i + 1}: missing 'programme_name'."
                    };
                }

                if (string.IsNullOrWhiteSpace(spot.ProgrammeStartTime))
                {
                    return new JsonImportResult
                    {
                        Success = false,
                        ErrorMessage = $"Spot #{i + 1}: missing 'programme_start_time'."
                    };
                }

                if (spot.Dates == null)
                {
                    spot.Dates = new List<string>();
                }
            }

            return new JsonImportResult
            {
                Success = true,
                BookingOrder = order
            };
        }
    }
}
