using System;
using System.Collections.Generic;

namespace SSISWCFTask100
{
    internal static class Keys
    {
        public const string SERVICE_URL = "ServiceUrl";
        public const string SERVICE_CONTRACT = "ServiceContract";
        public const string OPERATION_CONTRACT = "OperationContract";
        public const string MAPPING_PARAMS = "MappingParams";
        public const string MAPPING_HEADERS = "MappingHeaders";
        public const string RETURNED_VALUE = "ReturnedValue";
        public const string IS_VALUE_RETURNED = "IsValueReturned";

        public const string TRUE = "True";
        public const string FALSE = "False";
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

    [Serializable]
    public class MappingHeaders : List<MappingParam>
    {
    }

    public class ComboBoxObjectComboItem
    {
        /// <summary>
        /// Gets or sets the value memeber.
        /// </summary>
        /// <value>The value memeber.</value>
        public object ValueMemeber { get; private set; }

        /// <summary>
        /// Gets or sets the display member.
        /// </summary>
        /// <value>The display member.</value>
        public object DisplayMember { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ComboBoxObjectComboItem"/> class.
        /// </summary>
        /// <param name="aBindingValue">A binding value.</param>
        /// <param name="aDisplayValue">A display value.</param>
        public ComboBoxObjectComboItem(object aBindingValue, object aDisplayValue)
        {
            ValueMemeber = aBindingValue;
            DisplayMember = aDisplayValue;
        }

        /// <summary>
        /// Returns a <see cref="System.String"/> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String"/> that represents this instance.
        /// </returns>
        public override String ToString()
        {
            return Convert.ToString(DisplayMember);
        }
    }
}
