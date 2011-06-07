using System;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Web.Services;
using System.Web.Services.Description;
using System.Web.Services.Discovery;
using System.Web.Services.Protocols;
using System.Xml;
using System.Xml.Serialization;
using Microsoft.CSharp;

namespace SSISWCFTask100
{
    [WebServiceBinding(Name = "WebServiceSoap", Namespace = "http://tempuri.org/")]
    public class WSDLHandler : SoapHttpClientProtocol
    {
        #region Propertis

        private string _wsdlContent = string.Empty;

        private readonly bool _useDefaultCredentialsSetExplicitly;
        private static readonly List<string> AssemblyReferences = new List<string> {
                                                                                       "System.Web.Services.dll",
                                                                                       "System.Xml.dll",
                                                                                       "System.Data.dll"
                                                                                   };


        public new string Url
        {
            get { return base.Url; }
            set
            {
                if ((((IsLocalFileSystemWebService(base.Url))
                      && (_useDefaultCredentialsSetExplicitly == false))
                     && (IsLocalFileSystemWebService(value) == false)))
                {
                    UseDefaultCredentials = false;
                }

                base.Url = value;
            }
        }

        public WebServiceMethods WebServiceMethods { get; set; }

        public Type ReturnedTypeOfTheSelectedMethod { get; set; }

        #endregion

        #region .ctor

        public WSDLHandler()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WSDLHandler"/> class.
        /// </summary>
        /// <param name="webServiceUri">The web service URI.</param>
        public WSDLHandler(Uri webServiceUri)
        {
            Url = webServiceUri.AbsoluteUri + "?wsdl";

            if ((IsLocalFileSystemWebService(Url)))
            {
                UseDefaultCredentials = true;
                _useDefaultCredentialsSetExplicitly = false;
            }
            else
            {
                _useDefaultCredentialsSetExplicitly = true;
            }

            _services = new List<string>();
            _availableTypes = new Dictionary<string, Type>();

            _webServiceAssembly = BuildAssembly(webServiceUri);

            Type[] types = _webServiceAssembly.GetExportedTypes();

            foreach (Type type in types)
            {
                _services.Add(type.FullName);
                _availableTypes.Add(type.FullName, type);
                GetServiceMethods(type.FullName);
            }
            GetWSDL();
            GetWebReferences();
        }

        #endregion

        #region WSDL

        public string GetWSDL()
        {
            using (var client = new WebClient())
            {
                _wsdlContent = client.DownloadString(Url);
                //client.DownloadFile(Url, @"C:\wsdl.xml");
            }

            return _wsdlContent;
        }

        /// <summary>
        /// Gets the web methods.
        /// </summary>
        /// <returns></returns>
        public WebServiceMethods GetWebMethods(MethodInfo[] methodInfos)
        {
            WebServiceMethods = new WebServiceMethods();

            foreach (var methodInfo in methodInfos)
            {
                WebServiceMethodParameters webServiceMethodParameters = new WebServiceMethodParameters();
                webServiceMethodParameters.AddRange(methodInfo.GetParameters().Select(parameterInfo => new WebServiceMethodParameter
                                                                                                           {
                                                                                                               Name = parameterInfo.Name,
                                                                                                               Type = parameterInfo.ParameterType.FullName
                                                                                                           }));

                WebServiceMethods.Add(new WebServiceMethod
                                          {
                                              Name = methodInfo.Name,
                                              ResultType = methodInfo.ReturnType.ToString(),
                                              WebServiceMethodParameters = webServiceMethodParameters
                                          });

                AssemblyReferences.Add((methodInfo.ReturnType).Module.Name);
                AssemblyReferences.AddRange(from parameters in methodInfo.GetParameters()
                                            select parameters.ParameterType.Module.ToString());
            }

            return WebServiceMethods;
        }

        #endregion

        #region Tools

        private readonly Dictionary<string, Type> _availableTypes;
        private readonly List<string> _services;
        private readonly Assembly _webServiceAssembly;

        /// <summary>
        /// Text description of the available services within this web service.
        /// </summary>
        /// <value>The available services.</value>
        public List<string> AvailableServices
        {
            get { return _services; }
        }

        /// <summary>
        /// Determines whether [is local file system web service] [the specified URL].
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <returns>
        /// 	<c>true</c> if [is local file system web service] [the specified URL]; otherwise, <c>false</c>.
        /// </returns>
        private static bool IsLocalFileSystemWebService(string url)
        {
            if (string.IsNullOrEmpty(url))
                return false;

            var wsUri = new Uri(url);
            return ((wsUri.Port >= 1024) &&
                    (string.Compare(wsUri.Host, "localHost", StringComparison.OrdinalIgnoreCase) == 0));
        }

        /// <summary>
        /// Invokes the remote method.
        /// </summary>
        /// <returns></returns>
        public object ExecuteWebMethod(string methodName, params object[] webMethodParam)
        {
            WebRequest webRequest = WebRequest.Create(Url);
            Stream getResponseStream = webRequest.GetResponse().GetResponseStream();
            ServiceDescription serviceDescription = ServiceDescription.Read(getResponseStream);


            string serviceDescriptionName = serviceDescription.Services[0].Name;
            var serviceDescriptionImporter = new ServiceDescriptionImporter
            {
                ProtocolName = "Soap12"
            };

            serviceDescriptionImporter.AddServiceDescription(serviceDescription, null, null);

            var codeNamespace = new CodeNamespace();
            var codeCompileUnit = new CodeCompileUnit();
            codeCompileUnit.Namespaces.Add(codeNamespace);

            ServiceDescriptionImportWarnings serviceDescriptionImportWarnings = serviceDescriptionImporter.Import(codeNamespace, codeCompileUnit);

            object returnedObject = null;

            if (serviceDescriptionImportWarnings == 0)
            {
                CodeDomProvider codeDomProvider = new CSharpCodeProvider();
                var AssemblyReferences = new List<string>
                    {
                        "System.Web.Services.dll",
                        "System.Xml.dll",
                        "System.Data.dll"
                    };
                var references = AssemblyReferences.Distinct().ToArray();

                var compilerParameters = new CompilerParameters(references);
                CompilerResults compilerResults = codeDomProvider.CreateCompiler().CompileAssemblyFromDom(compilerParameters, codeCompileUnit);

                object createInstance = compilerResults.CompiledAssembly.CreateInstance(serviceDescriptionName);
                Type getType = createInstance.GetType();
                MethodInfo methodInfo = getType.GetMethod("Invoke", BindingFlags.NonPublic | BindingFlags.Instance);

                object[] objParams = {
                                         methodName,
                                         webMethodParam
                                     };

                returnedObject = methodInfo.Invoke(createInstance, objParams);
            }

            return returnedObject;
        }


        /// <summary>
        /// Gets the service methods.
        /// </summary>
        /// <param name="serviceName">Name of the service.</param>
        /// <returns></returns>
        public List<string> GetServiceMethods(string serviceName)
        {
            var methods = new List<string>();

            if (!_availableTypes.ContainsKey(serviceName))
                throw new Exception("Service Not Available");

            Type type = _availableTypes[serviceName];

            methods.AddRange(type.GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.DeclaredOnly).Select(minfo => minfo.Name));

            WebServiceMethods = GetWebMethods(type.GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.DeclaredOnly));

            return methods;
        }

        /// <summary>
        /// Invokes the remote method.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="serviceName">Name of the service.</param>
        /// <param name="methodName">Name of the method.</param>
        /// <param name="args">The args.</param>
        /// <returns></returns>
        public object[] InvokeRemoteMethod<T>(string serviceName, string methodName, params object[] args)
        {
            object obj = _webServiceAssembly.CreateInstance(serviceName);

            Type type = obj.GetType();

            return new object[]
                       {
                           (T)type.InvokeMember(methodName, BindingFlags.InvokeMethod, null, obj, args)
                       };
        }

        /// <summary>
        /// Builds the service description.
        /// </summary>
        /// <param name="xmlreader">The xmlreader.</param>
        /// <returns></returns>
        private static ServiceDescriptionImporter BuildServiceDescription(XmlTextReader xmlreader)
        {
            if (!ServiceDescription.CanRead(xmlreader))
                throw new Exception("Invalid Web Service Description");

            ServiceDescription serviceDescription = ServiceDescription.Read(xmlreader);

            var descriptionImporter = new ServiceDescriptionImporter
            {
                ProtocolName = "Soap"
            };

            descriptionImporter.AddServiceDescription(serviceDescription, null, null);
            descriptionImporter.Style = ServiceDescriptionImportStyle.Client;
            descriptionImporter.CodeGenerationOptions = CodeGenerationOptions.GenerateProperties;

            return descriptionImporter;
        }

        public void GetWebReferences()
        {
            DiscoveryClientDocumentCollection wsdlCollection = new DiscoveryClientDocumentCollection();
            ServiceDescription description = ServiceDescription.Read(new StringReader(_wsdlContent));
            wsdlCollection.Add(Url, description);

            CodeDomProvider provider = CodeDomProvider.CreateProvider("CSharp");

            // Create a namespace and a unit for compilation.
            CodeCompileUnit unit = new CodeCompileUnit();
            CodeNamespace space = new CodeNamespace("eBay.Soap");
            unit.Namespaces.Add(space);

            // Create a web referernce using the WSDL collection.
            WebReference reference = new WebReference(wsdlCollection, space) {ProtocolName = "Soap12"};

            // Create a web reference collection.
            WebReferenceCollection references = new WebReferenceCollection();
            references.Add(reference);

            WebReferenceOptions options = new WebReferenceOptions
            {
                Style = ServiceDescriptionImportStyle.Client,
                CodeGenerationOptions = CodeGenerationOptions.GenerateNewAsync
            };

            StringCollection results = ServiceDescriptionImporter.GenerateWebReferences(references, provider, unit, options);
        }

        /// <summary>
        /// Creates the assembly.
        /// </summary>
        /// <param name="descriptionImporter">The description importer.</param>
        /// <returns></returns>
        private static Assembly CreateAssembly(ServiceDescriptionImporter descriptionImporter)
        {
            var codeNamespace = new CodeNamespace();
            var codeUnit = new CodeCompileUnit();

            codeUnit.Namespaces.Add(codeNamespace);

            ServiceDescriptionImportWarnings importWarnings = descriptionImporter.Import(codeNamespace, codeUnit);

            if (importWarnings != 0)
            {
                throw new Exception("Invalid WSDL");
            }

            CodeDomProvider compiler = CodeDomProvider.CreateProvider("CSharp");

            //var references = AssemblyReferences.Distinct().ToArray();

            //descriptionImporter.GenerateWebReferences()

            
            var AssemblyReferences = new List<string>
                    {
                        "System.Web.Services.dll",
                        "System.Xml.dll",
                        "System.Data.dll"
                    };
            var references = AssemblyReferences.Distinct().ToArray();
            var parameters = new CompilerParameters(references);

            CompilerResults results = compiler.CompileAssemblyFromDom(parameters, codeUnit);

            if (results.Errors.Cast<CompilerError>().Any())
            {
                var stringBuilder = new StringBuilder();

                foreach (var error in results.Errors)
                {
                    stringBuilder.Append(error.ToString());
                }

                throw new Exception(string.Format("Compilation Error Creating Assembly :: {0}", stringBuilder));
            }

            return results.CompiledAssembly;
        }

        /// <summary>
        /// Builds the assembly.
        /// </summary>
        /// <param name="webServiceUri">The web service URI.</param>
        /// <returns></returns>
        private static Assembly BuildAssembly(Uri webServiceUri)
        {
            if (String.IsNullOrEmpty(webServiceUri.ToString()))
                throw new Exception("Web Service Not Found");

            var xmlreader = new XmlTextReader(webServiceUri + "?wsdl");

            ServiceDescriptionImporter descriptionImporter = BuildServiceDescription(xmlreader);

            return CreateAssembly(descriptionImporter);
        }

        #endregion
    }

    #region Methods & Parameters

    public class WebServiceMethods : List<WebServiceMethod>
    {
    }

    public class WebServiceMethod
    {
        public WebServiceMethod()
        {
            WebServiceMethodParameters = new WebServiceMethodParameters();
        }

        public string Name { get; set; }
        public string ResultType { get; set; }

        public WebServiceMethodParameters WebServiceMethodParameters { get; set; }
    }

    public class WebServiceMethodParameters : List<WebServiceMethodParameter>
    {
    }

    public class WebServiceMethodParameter
    {
        public string Name { get; set; }
        public string Type { get; set; }
    }

    #endregion
}


///// <summary>
///// Gets the web methods. //OBSOLETE
///// </summary>
///// <returns></returns>
//public WebServiceMethods GetWebMethods()
//{
//    _wsdlContent = GetWSDL();

//    XDocument xDocument = XDocument.Parse(_wsdlContent);
//    XNamespace xNamespaceWSDL = "http://schemas.xmlsoap.org/wsdl/";
//    XNamespace xNamespace = "http://www.w3.org/2001/XMLSchema";

//    XElement schema = xDocument.Root
//        .Element(xNamespaceWSDL + "types")
//        .Element(xNamespace + "schema");

//    IEnumerable<XElement> elements = schema.Elements(xNamespace + "element");

//    Func<XElement, string> getName = el => el.Attribute("name").Value;
//    Func<XElement, string> getType = el => el.Attribute("type").Value;

//    var webServiceMethods = new WebServiceMethods();

//    foreach (WebServiceMethod Method in (from element in elements
//                                         let name = getName(element)
//                                         where name.EndsWith("Response")
//                                         select new WebServiceMethod
//                                         {
//                                             Name = getName(element).Replace("Response", string.Empty),
//                                             ResultType = getType(element.Descendants(xNamespace + "element").First())
//                                         }))
//    {
//        var webServiceMethod = new WebServiceMethod
//        {
//            Name = Method.Name,
//            ResultType = Method.ResultType.Split(':')[1]
//        };

//        XElement method = elements.Single(el => getName(el) == webServiceMethod.Name);

//        IEnumerable<WebServiceMethodParameter> parameters =
//            from par in method.Descendants(xNamespace + "element")
//            let name = getName(par)
//            where !name.EndsWith("Response")
//            select new WebServiceMethodParameter
//            {
//                Name = getName(par),
//                Type = getType(par).Split(':')[1]
//            };

//        webServiceMethod.WebServiceMethodParameters.AddRange(parameters);

//        webServiceMethods.Add(webServiceMethod);
//    }

//    return webServiceMethods;
//}

///// <summary>
///// Gets the WSDL.
///// </summary>
///// <returns></returns>
//public string GetWSDL()
//{
//    using (var client = new WebClient())
//    {
//        wsdlContent = client.DownloadString(Url);
//        //client.DownloadFile(Url, @"C:\wsdl.xml");
//    }
//    return wsdlContent;
//}