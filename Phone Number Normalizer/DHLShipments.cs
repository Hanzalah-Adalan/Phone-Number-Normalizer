namespace DHL
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    public partial class DhlShipments
    {
        [JsonProperty("shipments")]
        public List<Shipment> Shipments { get; set; }

        [JsonProperty("possibleAdditionalShipmentsUrl")]
        public List<string> PossibleAdditionalShipmentsUrl { get; set; }
    }

    public partial class Shipment
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("service")]
        public string Service { get; set; }

        [JsonProperty("origin")]
        public Destination Origin { get; set; }

        [JsonProperty("destination")]
        public Destination Destination { get; set; }

        [JsonProperty("status")]
        public Status Status { get; set; }

        [JsonProperty("details")]
        public Details Details { get; set; }

        [JsonProperty("events")]
        public List<Status> Events { get; set; }
    }

    public partial class Destination
    {
        [JsonProperty("address")]
        public Address Address { get; set; }
    }

    public partial class Address
    {
        [JsonProperty("countryCode")]
        public string CountryCode { get; set; }

        [JsonProperty("postalCode")]
        public string PostalCode { get; set; }

        [JsonProperty("addressLocality")]
        public string AddressLocality { get; set; }
    }

    public partial class Details
    {
        [JsonProperty("product")]
        public Product Product { get; set; }

        [JsonProperty("references")]
        public List<Reference> References { get; set; }
    }

    public partial class Product
    {
        [JsonProperty("productName")]
        public string ProductName { get; set; }
    }

    public partial class Reference
    {
        [JsonProperty("number")]
        public string Number { get; set; }

        [JsonProperty("type")]
        public string Type { get; set; }
    }

    public partial class Status
    {
        [JsonProperty("timestamp")]
        public DateTimeOffset Timestamp { get; set; }

        [JsonProperty("location")]
        public Destination Location { get; set; }

        [JsonProperty("statusCode")]
        public string StatusCode { get; set; }

        [JsonProperty("status")]
        public string StatusStatus { get; set; }
    }

    public partial class DhlShipments
    {
        public static DhlShipments FromJson(string json) => JsonConvert.DeserializeObject<DhlShipments>(json, DHL.Converter.Settings);
    }

    public static class Serialize
    {
        public static string ToJson(this DhlShipments self) => JsonConvert.SerializeObject(self, DHL.Converter.Settings);
    }

    internal static class Converter
    {
        public static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            MetadataPropertyHandling = MetadataPropertyHandling.Ignore,
            DateParseHandling = DateParseHandling.None,
            Converters =
            {
                new IsoDateTimeConverter { DateTimeStyles = DateTimeStyles.AssumeUniversal }
            },
        };
    }
}
