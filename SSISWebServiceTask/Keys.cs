using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;

namespace SSISWCFTask100
{
    internal static class Keys
    {
        public const string SERVICE_URL = "ServiceUrl";
        public const string SERVICE = "Service";
        public const string WEBMETHOD = "WebMethod";
        public const string MAPPING_PARAMS = "MappingParams";
        public const string RETURNED_VALUE = "ReturnedValue";
        public const string IS_VALUE_RETURNED = "IsValueReturned";
    }


    [Serializable]
    public class MappingParam
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Value { get; set; }
    }

    [Serializable]
    public class MappingParams : List<MappingParam>
    {
    }
}
