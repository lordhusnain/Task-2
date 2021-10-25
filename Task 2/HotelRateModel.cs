using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace Task_2
{
    public class HotelRateModel
    {
        [JsonProperty("hotel")]
        public Hotel Hotel { get; set; }
        [JsonProperty("hotelRates")]
        public List<HotelRate> HotelRates { get; set; }
    }

    public class Hotel
    {
        [JsonProperty("classification")]
        public int Classification { get; set; }
        [JsonProperty("hotelID")]
        public int HotelID { get; set; }
        [JsonProperty("name")]
        public string Name { get; set; }
        [JsonProperty("reviewscore")]
        public double Reviewscore { get; set; }
    }

    public class Price
    {
        [JsonProperty("currency")]
        public string Currency { get; set; }
        [JsonProperty("numericFloat")]
        public double NumericFloat { get; set; }
        [JsonProperty("numericInteger")]
        public int NumericInteger { get; set; }
    }

    public class RateTag
    {
        [JsonProperty("name")]
        public string Name { get; set; }
        [JsonProperty("shape")]
        public bool Shape { get; set; }
    }

    public class HotelRate
    {
        [JsonProperty("adults")]
        public int Adults { get; set; }
        [JsonProperty("los")]
        public int Los { get; set; }
        [JsonProperty("price")]
        public Price Price { get; set; }
        [JsonProperty("rateDescription")]
        public string RateDescription { get; set; }
        [JsonProperty("rateID")]
        public string RateID { get; set; }
        [JsonProperty("rateName")]
        public string RateName { get; set; }
        [JsonProperty("rateTags")]
        public List<RateTag> RateTags { get; set; }
        [JsonProperty("targetDay")]
        public DateTime TargetDay { get; set; }
    }
}
