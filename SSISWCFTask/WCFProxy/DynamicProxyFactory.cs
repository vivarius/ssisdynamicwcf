using System;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Data;
using System.Data.Design;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Description;
using System.Text;
using System.Web.Services.Description;
using System.Web.Services.Discovery;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using WcfSamples.DynamicProxy;
using Binding = System.ServiceModel.Channels.Binding;
using ServiceDescription = System.Web.Services.Description.ServiceDescription;

namespace SSISWCFTask100.WCFProxy
{
    public class DynamicProxyFactory
    {
        internal const string DefaultNamespace = "http://tempuri.org/";
        private readonly DynamicProxyFactoryOptions _options;
        private readonly string _wsdlUri;
        private IEnumerable<Binding> _bindings;

        private CodeCompileUnit _codeCompileUnit;
        private CodeDomProvider _codeDomProvider;
        private IEnumerable<MetadataConversionError> _codegenWarnings;
        private IEnumerable<CompilerError> _compilerWarnings;
        private ServiceContractGenerator _contractGenerator;

        private IEnumerable<ContractDescription> _contracts;
        private ServiceEndpointCollection _endpoints;
        private IEnumerable<MetadataConversionError> _importWarnings;
        private Collection<MetadataSection> _metadataCollection;

        private Assembly _proxyAssembly;
        private string _proxyCode;

        public DynamicProxyFactory(string wsdlUri, DynamicProxyFactoryOptions options)
        {
            if (wsdlUri == null)
                throw new ArgumentNullException("wsdlUri");

            if (options == null)
                throw new ArgumentNullException("options");

            _wsdlUri = wsdlUri;
            _options = options;

            DownloadMetadata();
            ImportMetadata();
            CreateProxy();
            WriteCode();
            CompileProxy();
        }

        public DynamicProxyFactory(string wsdlUri)
            : this(wsdlUri, new DynamicProxyFactoryOptions())
        {
        }

        public IEnumerable<MetadataSection> Metadata
        {
            get { return _metadataCollection; }
        }

        public IEnumerable<Binding> Bindings
        {
            get { return _bindings; }
        }

        public IEnumerable<ContractDescription> Contracts
        {
            get { return _contracts; }
        }

        public IEnumerable<ServiceEndpoint> Endpoints
        {
            get { return _endpoints; }
        }

        public Assembly ProxyAssembly
        {
            get { return _proxyAssembly; }
        }

        public string ProxyCode
        {
            get { return _proxyCode; }
        }

        public IEnumerable<MetadataConversionError> MetadataImportWarnings
        {
            get { return _importWarnings; }
        }

        public IEnumerable<MetadataConversionError> CodeGenerationWarnings
        {
            get { return _codegenWarnings; }
        }

        public IEnumerable<CompilerError> CompilationWarnings
        {
            get { return _compilerWarnings; }
        }

        private void DownloadMetadata()
        {
            var epr = new EndpointAddress(_wsdlUri);

            var disco = new DiscoveryClientProtocol
                            {
                                AllowAutoRedirect = true,
                                UseDefaultCredentials = true
                            };

            disco.DiscoverAny(_wsdlUri);
            disco.ResolveAll();

            var results = new Collection<MetadataSection>();
            foreach (object document in disco.Documents.Values)
            {
                AddDocumentToResults(document, results);
            }
            _metadataCollection = results;
        }

        private void AddDocumentToResults(object document, Collection<MetadataSection> results)
        {
            var wsdl = document as ServiceDescription;
            var schema = document as XmlSchema;
            var xmlDoc = document as XmlElement;

            if (wsdl != null)
            {
                results.Add(MetadataSection.CreateFromServiceDescription(wsdl));
            }
            else if (schema != null)
            {
                results.Add(MetadataSection.CreateFromSchema(schema));
            }
            else if (xmlDoc != null && xmlDoc.LocalName == "Policy")
            {
                results.Add(MetadataSection.CreateFromPolicy(xmlDoc, null));
            }
            else
            {
                var mexDoc = new MetadataSection { Metadata = document };
                results.Add(mexDoc);
            }
        }


        private void ImportMetadata()
        {
            _codeCompileUnit = new CodeCompileUnit();
            CreateCodeDomProvider();

            var importer = new WsdlImporter(new MetadataSet(_metadataCollection));
            AddStateForDataContractSerializerImport(importer);
            AddStateForXmlSerializerImport(importer);

            _bindings = importer.ImportAllBindings();
            _contracts = importer.ImportAllContracts();
            _endpoints = importer.ImportAllEndpoints();
            _importWarnings = importer.Errors;

            bool success = true;
            if (_importWarnings != null)
            {
                success = _importWarnings.All(error => error.IsWarning);
            }

            if (!success)
            {
                var exception = new DynamicProxyException(Constants.ErrorMessages.ImportError)
                                    {
                                        MetadataImportErrors = _importWarnings
                                    };
                throw exception;
            }
        }

        private void AddStateForXmlSerializerImport(WsdlImporter importer)
        {
            var importOptions = new XmlSerializerImportOptions(_codeCompileUnit)
                                    {
                                        CodeProvider = _codeDomProvider,
                                        WebReferenceOptions = new WebReferenceOptions
                                                                  {
                                                                      CodeGenerationOptions = CodeGenerationOptions.GenerateProperties | CodeGenerationOptions.GenerateOrder
                                                                  }
                                    };

            importOptions.WebReferenceOptions.SchemaImporterExtensions.Add(typeof(TypedDataSetSchemaImporterExtension).AssemblyQualifiedName);
            importOptions.WebReferenceOptions.SchemaImporterExtensions.Add(typeof(DataSetSchemaImporterExtension).AssemblyQualifiedName);

            importer.State.Add(typeof(XmlSerializerImportOptions), importOptions);
        }

        private void AddStateForDataContractSerializerImport(WsdlImporter importer)
        {
            var xsdDataContractImporter = new XsdDataContractImporter(_codeCompileUnit)
                                              {
                                                  Options = new ImportOptions
                                                                {
                                                                    ImportXmlType = (_options.FormatMode == DynamicProxyFactoryOptions.FormatModeOptions.DataContractSerializer),
                                                                    CodeProvider = _codeDomProvider
                                                                }
                                              };

            importer.State.Add(typeof(XsdDataContractImporter), xsdDataContractImporter);

            foreach (var dcConverter in importer.WsdlImportExtensions.OfType<DataContractSerializerMessageContractImporter>())
            {
                dcConverter.Enabled = _options.FormatMode != DynamicProxyFactoryOptions.FormatModeOptions.XmlSerializer;
            }
        }

        private void CreateProxy()
        {
            CreateServiceContractGenerator();

            foreach (ContractDescription contract in _contracts)
            {
                _contractGenerator.GenerateServiceContractType(contract);
            }

            bool success = true;
            _codegenWarnings = _contractGenerator.Errors;
            if (_codegenWarnings != null)
            {
                success = _codegenWarnings.All(error => error.IsWarning);
            }

            if (!success)
            {
                var exception = new DynamicProxyException(Constants.ErrorMessages.CodeGenerationError)
                                    {
                                        CodeGenerationErrors = _codegenWarnings
                                    };
                throw exception;
            }
        }

        private void CompileProxy()
        {
            // reference the required assemblies with the correct path.
            var compilerParams = new CompilerParameters();

            AddAssemblyReference(typeof(ServiceContractAttribute).Assembly, compilerParams.ReferencedAssemblies);

            AddAssemblyReference(typeof(ServiceDescription).Assembly, compilerParams.ReferencedAssemblies);

            AddAssemblyReference(typeof(DataContractAttribute).Assembly, compilerParams.ReferencedAssemblies);

            AddAssemblyReference(typeof(XmlElement).Assembly, compilerParams.ReferencedAssemblies);

            AddAssemblyReference(typeof(Uri).Assembly, compilerParams.ReferencedAssemblies);

            AddAssemblyReference(typeof(DataSet).Assembly, compilerParams.ReferencedAssemblies);

            CompilerResults results = _codeDomProvider.CompileAssemblyFromSource(compilerParams, _proxyCode);

            if ((results.Errors != null) && (results.Errors.HasErrors))
            {
                var exception = new DynamicProxyException(Constants.ErrorMessages.CompilationError)
                                    {
                                        CompilationErrors = ToEnumerable(results.Errors)
                                    };

                throw exception;
            }

            _compilerWarnings = ToEnumerable(results.Errors);
            _proxyAssembly = Assembly.LoadFile(results.PathToAssembly);
        }

        private void WriteCode()
        {
            using (var writer = new StringWriter())
            {
                var codeGenOptions = new CodeGeneratorOptions
                                         {
                                             BracingStyle = "C"
                                         };

                _codeDomProvider.GenerateCodeFromCompileUnit(_codeCompileUnit, writer, codeGenOptions);
                writer.Flush();
                _proxyCode = writer.ToString();
            }

            // use the modified proxy code, if code modifier is set.
            if (_options.CodeModifier != null)
                _proxyCode = _options.CodeModifier(_proxyCode);
        }

        private void AddAssemblyReference(Assembly referencedAssembly, StringCollection refAssemblies)
        {
            if (referencedAssembly != null)
            {
                if (referencedAssembly.Location != null)
                {
                    string path = Path.GetFullPath(referencedAssembly.Location);
                    string name = Path.GetFileName(path);
                    if (!(refAssemblies.Contains(name) || refAssemblies.Contains(path)))
                    {
                        refAssemblies.Add(path);
                    }
                }
            }
        }

        public ServiceEndpoint GetEndpoint(string contractName)
        {
            return GetEndpoint(contractName, null);
        }

        public ServiceEndpoint GetEndpoint(string contractName, string contractNamespace)
        {
            ServiceEndpoint matchingEndpoint = Endpoints.FirstOrDefault(endpoint => ContractNameMatch(endpoint.Contract, contractName) && ContractNsMatch(endpoint.Contract, contractNamespace));

            if (matchingEndpoint == null)
                throw new ArgumentException(string.Format(Constants.ErrorMessages.EndpointNotFound, contractName, contractNamespace));

            return matchingEndpoint;
        }

        private bool ContractNameMatch(ContractDescription cDesc, string name)
        {
            return (string.Compare(cDesc.Name, name, true) == 0);
        }

        private bool ContractNsMatch(ContractDescription cDesc, string ns)
        {
            return ((ns == null) || (string.Compare(cDesc.Namespace, ns, true) == 0));
        }

        public DynamicProxy CreateProxy(string contractName)
        {
            return CreateProxy(contractName, null);
        }

        public DynamicProxy CreateProxy(string contractName, string contractNamespace)
        {
            ServiceEndpoint endpoint = GetEndpoint(contractName, contractNamespace);

            return CreateProxy(endpoint);
        }

        public DynamicProxy CreateProxy(ServiceEndpoint endpoint)
        {
            Type contractType = GetContractType(endpoint.Contract.Name, endpoint.Contract.Namespace);

            Type proxyType = GetProxyType(contractType);

            return new DynamicProxy(proxyType, endpoint.Binding, endpoint.Address);
        }

        private Type GetContractType(string contractName, string contractNamespace)
        {
            Type[] allTypes = _proxyAssembly.GetTypes();
            ServiceContractAttribute scAttr = null;
            Type contractType = null;
            XmlQualifiedName cName;

            foreach (Type type in allTypes)
            {
                // Is it an interface?
                if (!type.IsInterface) continue;

                // Is it marked with ServiceContract attribute?
                object[] attrs = type.GetCustomAttributes(typeof(ServiceContractAttribute), false);
                if ((attrs == null) || (attrs.Length == 0)) continue;

                // is it the required service contract?
                scAttr = (ServiceContractAttribute)attrs[0];
                cName = GetContractName(type, scAttr.Name, scAttr.Namespace);

                if (string.Compare(cName.Name, contractName, true) != 0)
                    continue;

                if (string.Compare(cName.Namespace, contractNamespace, true) != 0)
                    continue;

                contractType = type;
                break;
            }

            if (contractType == null)
                throw new ArgumentException(Constants.ErrorMessages.UnknownContract);

            return contractType;
        }

        internal static XmlQualifiedName GetContractName(Type contractType, string name, string ns)
        {
            if (String.IsNullOrEmpty(name))
            {
                name = contractType.Name;
            }

            ns = ns == null 
                    ? DefaultNamespace 
                    : Uri.EscapeUriString(ns);

            return new XmlQualifiedName(name, ns);
        }

        private Type GetProxyType(Type contractType)
        {
            Type clientBaseType = typeof(ClientBase<>).MakeGenericType(contractType);

            Type[] allTypes = ProxyAssembly.GetTypes();
            Type proxyType = allTypes.FirstOrDefault(type => type.IsClass && contractType.IsAssignableFrom(type) && type.IsSubclassOf(clientBaseType));

            if (proxyType == null)
                throw new DynamicProxyException(string.Format(Constants.ErrorMessages.ProxyTypeNotFound, contractType.FullName));

            return proxyType;
        }


        private void CreateCodeDomProvider()
        {
            _codeDomProvider = CodeDomProvider.CreateProvider(_options.Language.ToString());
        }

        private void CreateServiceContractGenerator()
        {
            _contractGenerator = new ServiceContractGenerator(_codeCompileUnit);
            _contractGenerator.Options |= ServiceContractGenerationOptions.ClientClass;
        }

        public static string ToString(IEnumerable<MetadataConversionError> importErrors)
        {
            if (importErrors != null)
            {
                var importErrStr = new StringBuilder();

                foreach (MetadataConversionError error in importErrors)
                {
                    if (error.IsWarning)
                        importErrStr.AppendLine("Warning : " + error.Message);
                    else
                        importErrStr.AppendLine("Error : " + error.Message);
                }

                return importErrStr.ToString();
            }

            return null;
        }

        public static string ToString(IEnumerable<CompilerError> compilerErrors)
        {
            if (compilerErrors != null)
            {
                var builder = new StringBuilder();
                foreach (CompilerError error in compilerErrors)
                    builder.AppendLine(error.ToString());

                return builder.ToString();
            }

            return null;
        }

        private static IEnumerable<CompilerError> ToEnumerable(CompilerErrorCollection collection)
        {
            if (collection == null)
                return null;

            return collection.Cast<CompilerError>().ToList();
        }
    }
}