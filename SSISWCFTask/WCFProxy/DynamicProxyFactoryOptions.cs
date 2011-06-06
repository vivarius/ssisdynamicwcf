using System.Text;

namespace SSISWCFTask100.WCFProxy
{
    public delegate string ProxyCodeModifier(string proxyCode);

    public class DynamicProxyFactoryOptions
    {
        #region FormatModeOptions enum

        public enum FormatModeOptions
        {
            Auto,
            XmlSerializer,
            DataContractSerializer
        }

        #endregion

        #region LanguageOptions enum

        public enum LanguageOptions
        {
            CS,
            VB
        } ;

        #endregion

        public DynamicProxyFactoryOptions()
        {
            Language = LanguageOptions.CS;
            FormatMode = FormatModeOptions.Auto;
            CodeModifier = null;
        }

        public LanguageOptions Language { get; set; }

        public FormatModeOptions FormatMode { get; set; }

        // The code modifier allows the user of the dynamic proxy factory to modify 
        // the generated proxy code before it is compiled and used. This is useful in 
        // situations where the generated proxy has to be modified manually for interop 
        // reason.
        public ProxyCodeModifier CodeModifier { get; set; }

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append("DynamicProxyFactoryOptions[");
            sb.Append("Language=" + Language);
            sb.Append(",FormatMode=" + FormatMode);
            sb.Append(",CodeModifier=" + CodeModifier);
            sb.Append("]");

            return sb.ToString();
        }
    }
}