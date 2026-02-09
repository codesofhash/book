using Newtonsoft.Json;
using System.Collections.Generic;

namespace CSharpFlexGrid
{
    public class CampaignPeriod
    {
        [JsonProperty("start_date")]
        public string StartDate { get; set; }

        [JsonProperty("end_date")]
        public string EndDate { get; set; }
    }

    public class Spot
    {
        [JsonProperty("programme_name")]
        public string ProgrammeName { get; set; }

        [JsonProperty("programme_start_time")]
        public string ProgrammeStartTime { get; set; }

        [JsonProperty("duration")]
        public string Duration { get; set; }

        [JsonProperty("dates")]
        public List<string> Dates { get; set; } = new List<string>();

        [JsonProperty("total_spots")]
        public int TotalSpots { get; set; }
    }

    public class BookingOrder
    {
        [JsonProperty("agency")]
        public string Agency { get; set; }

        [JsonProperty("advertiser")]
        public string Advertiser { get; set; }

        [JsonProperty("product")]
        public string Product { get; set; }

        [JsonProperty("company_name")]
        public string CompanyName { get; set; }

        [JsonProperty("campaign_period")]
        public CampaignPeriod CampaignPeriod { get; set; } = new CampaignPeriod();

        [JsonProperty("gross_cost")]
        public decimal GrossCost { get; set; }

        [JsonProperty("total_spots")]
        public int TotalSpots { get; set; }

        [JsonProperty("spots")]
        public List<Spot> Spots { get; set; } = new List<Spot>();
    }
}
