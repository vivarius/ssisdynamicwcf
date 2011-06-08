using System;
using System.ComponentModel;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Description;
using System.Xml;
using Microsoft.SqlServer.Dts.Runtime;
using Microsoft.SqlServer.Dts.Runtime.Wrapper;
using SSISWCFTask100.WCFProxy;
using DTSExecResult = Microsoft.SqlServer.Dts.Runtime.DTSExecResult;
using DTSProductLevel = Microsoft.SqlServer.Dts.Runtime.DTSProductLevel;
using VariableDispenser = Microsoft.SqlServer.Dts.Runtime.VariableDispenser;

namespace SSISWCFTask100
{
    [DtsTask(
        DisplayName = "Dynamic WCF Task",
        UITypeName = "SSISWCFTask100.SSISWCFTaskUIInterface" +
        ",SSISWCFTask100," +
        "Version=1.0.0.25," +
        "Culture=Neutral," +
        "PublicKeyToken=dbcd46b65a9ba84f",
        IconResource = "SSISWCFTask100.Communication.ico",
        TaskContact = "cosmin.vlasiu@gmail.com",
        RequiredProductLevel = DTSProductLevel.None
        )]
    public class SSISWCFTask : Task, IDTSComponentPersist
    {
        #region Constructor
        public SSISWCFTask()
        {
        }

        #endregion

        #region Public Properties
        [Category("Component specific"), Description("WCF URL")]
        public string ServiceUrl { get; set; }
        [Category("Component specific"), Description("ServiceContract")]
        public string ServiceContract { get; set; }
        [Category("Component specific"), Description("OperationContract")]
        public string OperationContract { get; set; }
        [Category("Component specific"), Description("MappingParams")]
        public object MappingParams { get; set; }
        [Category("Component specific"), Description("Output Variable")]
        public string ReturnedValue { get; set; }
        [Category("Component specific"), Description("The method returns a value? (O/1)")]
        public string IsValueReturned { get; set; }
        #endregion

        #region Private Properties

        Variables _vars = null;

        #endregion

        #region Validate

        /// <summary>
        /// Validate local parameters
        /// </summary>
        public override DTSExecResult Validate(Connections connections, VariableDispenser variableDispenser, IDTSComponentEvents componentEvents, IDTSLogging log)
        {
            bool isBaseValid = true;

            if (base.Validate(connections, variableDispenser, componentEvents, log) != DTSExecResult.Success)
            {
                componentEvents.FireError(0, "SSISWCFTask", "Base validation failed", "", 0);
                isBaseValid = false;
            }

            if (string.IsNullOrEmpty(ServiceUrl))
            {
                componentEvents.FireError(0, "SSISWCFTask", "An URL is required.", "", 0);
                isBaseValid = false;
            }

            if (string.IsNullOrEmpty(ServiceContract))
            {
                componentEvents.FireError(0, "SSISWCFTask", "A service contract is required.", "", 0);
                isBaseValid = false;
            }

            if (string.IsNullOrEmpty(OperationContract))
            {
                componentEvents.FireError(0, "SSISWCFTask", "An Operation Contract name is required.", "", 0);
                isBaseValid = false;
            }

            return isBaseValid ? DTSExecResult.Success : DTSExecResult.Failure;
        }

        #endregion

        #region Execute

        /// <summary>
        /// This method is a run-time method executed dtsexec.exe
        /// </summary>
        /// <param name="connections"></param>
        /// <param name="variableDispenser"></param>
        /// <param name="componentEvents"></param>
        /// <param name="log"></param>
        /// <param name="transaction"></param>
        /// <returns></returns>
        public override DTSExecResult Execute(Connections connections, VariableDispenser variableDispenser, IDTSComponentEvents componentEvents, IDTSLogging log, object transaction)
        {
            bool refire = false;

            componentEvents.FireInformation(0,
                                            "SSISWCFTask",
                                            "Prepare variables",
                                            string.Empty,
                                            0,
                                            ref refire);

            GetNeededVariables(variableDispenser, componentEvents);

            try
            {
                componentEvents.FireInformation(0,
                                               "SSISWCFTask",
                                               string.Format("Initialize the WCF Service: {0}", EvaluateExpression(ServiceUrl, variableDispenser)),
                                               string.Empty,
                                               0,
                                               ref refire);
                var dynamicProxyFactory = new DynamicProxyFactory(EvaluateExpression(ServiceUrl, variableDispenser).ToString());

                int count = 0;
                foreach (ServiceEndpoint endpoint in dynamicProxyFactory.Endpoints.Where(endPoint => endPoint.Contract.Name == EvaluateExpression(ServiceContract, variableDispenser).ToString()))
                {
                    componentEvents.FireInformation(0, "SSISWCFTask", string.Format("Service Endpoint[{0}]", count++), string.Empty, 0, ref refire);
                    componentEvents.FireInformation(0, "SSISWCFTask", string.Format("Address: {0}", endpoint.Address), string.Empty, 0, ref refire);
                    componentEvents.FireInformation(0, "SSISWCFTask", string.Format("Contract: {0}", endpoint.Contract.Name), string.Empty, 0, ref refire);
                    componentEvents.FireInformation(0, "SSISWCFTask", string.Format("Binding: {0}", endpoint.Binding.Name), string.Empty, 0, ref refire);
                }

                componentEvents.FireInformation(0,
                                               "SSISWCFTask",
                                               string.Format("InvokeRemoteMethod: {0}=>{1}",
                                                             EvaluateExpression(ServiceContract, variableDispenser),
                                                             EvaluateExpression(OperationContract, variableDispenser)),
                                               string.Empty,
                                               0,
                                               ref refire);

                componentEvents.FireInformation(0, "SSISWCFTask", string.Format("Creating DynamicProxy to {0} ServiceContract", EvaluateExpression(ServiceContract, variableDispenser)), string.Empty, 0, ref refire);

                DynamicProxy dynamicProxy = dynamicProxyFactory.CreateProxy(EvaluateExpression(ServiceContract, variableDispenser).ToString());
                if (dynamicProxy == null)
                    throw new Exception("Cannot create the proxy");

                object result = null;

                componentEvents.FireInformation(0, "SSISWCFTask", "Creating the Client Channel", string.Empty, 0, ref refire);

                using (IClientChannel innerChannel = dynamicProxy.GetProperty("InnerChannel") as IClientChannel)
                {
                    if (innerChannel != null)
                    {
                        using (new OperationContextScope(innerChannel))
                        {
                            componentEvents.FireInformation(0, "SSISWCFTask", "Creating the Message Header", string.Empty, 0, ref refire);
                            OperationContext.Current.OutgoingMessageHeaders.Add(MessageHeader.CreateHeader("SSISDynamicWCF", "", Guid.NewGuid().ToString()));

                            componentEvents.FireInformation(0, "SSISWCFTask", "The used params:", string.Empty, 0, ref refire);
                            foreach (var param in ((MappingParams)MappingParams))
                            {
                                componentEvents.FireInformation(0, "SSISWCFTask",
                                                                string.Format("Name: {0} | Type: {1} | Value: {2}",
                                                                               param.Name,
                                                                               param.Type,
                                                                               EvaluateExpression(param.Value, variableDispenser)),
                                                                string.Empty, 0, ref refire);
                            }

                            componentEvents.FireInformation(0, "SSISWCFTask", string.Format("Call [{0}] OperationContract", EvaluateExpression(OperationContract, variableDispenser)), string.Empty, 0, ref refire);
                            result = dynamicProxy.CallMethod(EvaluateExpression(OperationContract, variableDispenser).ToString(),
                                                             new object[]
                                                                           {
                                                                                (from parameters in ((MappingParams)MappingParams)
                                                                                 select Convert.ChangeType(EvaluateExpression(parameters.Value, variableDispenser).ToString(), Type.GetType(parameters.Type))).ToArray()
                                                                           });
                        }
                    }
                    else
                    {
                        componentEvents.FireInformation(0, "SSISWCFTask", "Cannot create the Inner Channel", string.Empty, 0, ref refire);
                        throw new Exception("Cannot Create the Inner Channel");
                    }
                }

                if (result != null)
                {
                    if (IsValueReturned == Keys.TRUE)
                    {
                        componentEvents.FireInformation(0,
                                                        "SSISWCFTask",
                                                        string.Format("Get the Returned Value to: {0}", ReturnedValue),
                                                        string.Empty,
                                                        0,
                                                        ref refire);

                        string val = ReturnedValue.Split(new[] { "::" }, StringSplitOptions.RemoveEmptyEntries)[1].Trim();

                        componentEvents.FireInformation(0,
                                                        "SSISWCFTask",
                                                        string.Format("Get the Returned Value to {0} and convert to {1}",
                                                                      val.Substring(0, val.Length - 1),
                                                                      _vars[val.Substring(0, val.Length - 1)].DataType),
                                                        string.Empty,
                                                        0,
                                                        ref refire);

                        _vars[val.Substring(0, val.Length - 1)].Value = Convert.ChangeType(result, _vars[val.Substring(0, val.Length - 1)].DataType);

                        componentEvents.FireInformation(0,
                                                        "SSISWCFTask",
                                                        string.Format("The String Result is {0} ",
                                                                      _vars[val.Substring(0, val.Length - 1)].Value),
                                                        string.Empty,
                                                        0,
                                                        ref refire);
                    }
                    else
                    {
                        componentEvents.FireInformation(0,
                                                        "SSISWCFTask",
                                                        "Execution without return or no associated return variable",
                                                        string.Empty,
                                                        0,
                                                        ref refire);
                    }
                }
            }
            catch (Exception ex)
            {
                componentEvents.FireError(0,
                                          "SSISWCFTask",
                                          string.Format("Problem: {0}",
                                                        ex.Message + "\n" + ex.StackTrace),
                                          "",
                                          0);
            }
            finally
            {
                if (_vars.Locked)
                {
                    _vars.Unlock();
                }
            }

            return base.Execute(connections, variableDispenser, componentEvents, log, transaction);
        }

        #endregion

        #region Methods
        /// <summary>
        /// This method evaluate expressions like @([System::TaskName] + [System::TaskID]) or any other operation created using 
        /// ExpressionBuilder
        /// </summary>
        /// <param name="mappedParam"></param>
        /// <param name="variableDispenser"></param>
        /// <returns></returns>
        private static object EvaluateExpression(string mappedParam, VariableDispenser variableDispenser)
        {
            object variableObject = null;
            try
            {
                var expressionEvaluatorClass = new ExpressionEvaluatorClass
                {
                    Expression = mappedParam
                };

                expressionEvaluatorClass.Evaluate(DtsConvert.GetExtendedInterface(variableDispenser), out variableObject, false);
            }
            catch
            {
                variableObject = mappedParam;
            }
            return variableObject;
        }

        /// <summary>
        /// Gets the needed variables.
        /// </summary>
        /// <param name="variableDispenser">The variable dispenser.</param>
        /// <param name="componentEvents">The component events.</param>
        private void GetNeededVariables(VariableDispenser variableDispenser, IDTSComponentEvents componentEvents)
        {
            bool refire = false;

            try
            {
                var param = ServiceUrl;

                componentEvents.FireInformation(0, "SSISWCFTask", "ServiceUrl = " + ServiceUrl, string.Empty, 0, ref refire);

                if (param.Contains("@"))
                {
                    var regexStr = param.Split('@');

                    foreach (var nexSplitedVal in regexStr.Where(val => val.Trim().Length != 0).Select(strVal => strVal.Split(new[] { "::" }, StringSplitOptions.RemoveEmptyEntries)))
                    {
                        try
                        {
                            componentEvents.FireInformation(0, "SSISWCFTask", nexSplitedVal[1].Remove(nexSplitedVal[1].IndexOf(']')), string.Empty, 0, ref refire);
                            variableDispenser.LockForRead(nexSplitedVal[1].Remove(nexSplitedVal[1].IndexOf(']')));
                        }
                        catch (Exception exception)
                        {
                            throw new Exception(exception.Message);
                        }
                    }
                }
            }
            catch
            {
                //We will continue...
            }

            try
            {
                var param = ServiceContract;

                componentEvents.FireInformation(0, "SSISWCFTask", "ServiceContract = " + ServiceContract, string.Empty, 0, ref refire);

                if (param.Contains("@"))
                {
                    var regexStr = param.Split('@');

                    foreach (var nexSplitedVal in regexStr.Where(val => val.Trim().Length != 0).Select(strVal => strVal.Split(new[] { "::" }, StringSplitOptions.RemoveEmptyEntries)))
                    {
                        try
                        {
                            componentEvents.FireInformation(0, "SSISWCFTask", nexSplitedVal[1].Remove(nexSplitedVal[1].IndexOf(']')), string.Empty, 0, ref refire);
                            variableDispenser.LockForRead(nexSplitedVal[1].Remove(nexSplitedVal[1].IndexOf(']')));
                        }
                        catch (Exception exception)
                        {
                            throw new Exception(exception.Message);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                throw new Exception(exception.Message);
            }

            try
            {
                var param = OperationContract;

                componentEvents.FireInformation(0, "SSISWCFTask", "WebMethod = " + OperationContract, string.Empty, 0, ref refire);

                if (param.Contains("@"))
                {
                    var regexStr = param.Split('@');

                    foreach (var nexSplitedVal in regexStr.Where(val => val.Trim().Length != 0).Select(strVal => strVal.Split(new[] { "::" }, StringSplitOptions.RemoveEmptyEntries)))
                    {
                        try
                        {
                            componentEvents.FireInformation(0, "SSISWCFTask", nexSplitedVal[1].Remove(nexSplitedVal[1].IndexOf(']')), string.Empty, 0, ref refire);
                            variableDispenser.LockForRead(nexSplitedVal[1].Remove(nexSplitedVal[1].IndexOf(']')));
                        }
                        catch (Exception exception)
                        {
                            throw new Exception(exception.Message);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                throw new Exception(exception.Message);
            }

            try
            {

                if (!string.IsNullOrEmpty(ReturnedValue))
                {
                    var param = ReturnedValue;

                    componentEvents.FireInformation(0, "SSISWCFTask", "ReturnedValue = " + ReturnedValue,
                                                    string.Empty, 0, ref refire);

                    if (param.Contains("@"))
                    {
                        var regexStr = param.Split('@');

                        foreach (var nexSplitedVal in regexStr.Where(val => val.Trim().Length != 0).Select(strVal => strVal.Split(new[] { "::" }, StringSplitOptions.RemoveEmptyEntries)))
                        {
                            try
                            {
                                componentEvents.FireInformation(0, "SSISWCFTask",
                                                                nexSplitedVal[1].Remove(nexSplitedVal[1].IndexOf(']')),
                                                                string.Empty, 0, ref refire);
                                variableDispenser.LockForWrite(nexSplitedVal[1].Remove(nexSplitedVal[1].IndexOf(']')));
                            }
                            catch (Exception exception)
                            {
                                throw new Exception(exception.Message);
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                throw new Exception(exception.Message);
            }

            try
            {
                componentEvents.FireInformation(0, "SSISWCFTask", "MappingParams ", string.Empty, 0, ref refire);

                //Get variables for MappingParams
                foreach (var mappingParams in (MappingParams)MappingParams)
                {

                    try
                    {
                        if (mappingParams.Value.Contains("@"))
                        {
                            var regexStr = mappingParams.Value.Split('@');

                            foreach (var nexSplitedVal in
                                    regexStr.Where(val => val.Trim().Length != 0).Select(strVal => strVal.Split(new[] { "::" }, StringSplitOptions.RemoveEmptyEntries)))
                            {
                                try
                                {
                                    componentEvents.FireInformation(0, "SSISWCFTask", nexSplitedVal[1].Remove(nexSplitedVal[1].IndexOf(']')), string.Empty, 0, ref refire);
                                    variableDispenser.LockForRead(nexSplitedVal[1].Remove(nexSplitedVal[1].IndexOf(']')));
                                }
                                catch (Exception exception)
                                {
                                    throw new Exception(exception.Message);
                                }
                            }
                        }
                    }
                    catch (Exception exception)
                    {
                        throw new Exception(exception.Message);
                    }

                }
            }
            catch (Exception ex)
            {
                componentEvents.FireError(0, "SSISReportGeneratorTask", string.Format("Problem MappingParams: {0} {1}", ex.Message, ex.StackTrace), "", 0);
            }

            variableDispenser.GetVariables(ref _vars);
        }

        #endregion

        #region Implementation of IDTSComponentPersist

        /// <summary>
        /// Saves to XML.
        /// </summary>
        /// <param name="doc">The doc.</param>
        /// <param name="infoEvents">The info events.</param>
        void IDTSComponentPersist.SaveToXML(XmlDocument doc, IDTSInfoEvents infoEvents)
        {
            XmlElement taskElement = doc.CreateElement(string.Empty, "SSISWCFTask", string.Empty);

            XmlAttribute serviceUrl = doc.CreateAttribute(string.Empty, Keys.SERVICE_URL, string.Empty);
            serviceUrl.Value = ServiceUrl;

            XmlAttribute service = doc.CreateAttribute(string.Empty, Keys.SERVICE_CONTRACT, string.Empty);
            service.Value = ServiceContract;

            XmlAttribute webMethod = doc.CreateAttribute(string.Empty, Keys.OPERATION_CONTRACT, string.Empty);
            webMethod.Value = OperationContract;

            XmlAttribute mappingParams = doc.CreateAttribute(string.Empty, Keys.MAPPING_PARAMS, string.Empty);
            mappingParams.Value = Serializer.SerializeToXmlString(MappingParams);

            XmlAttribute returnedVariable = doc.CreateAttribute(string.Empty, Keys.RETURNED_VALUE, string.Empty);
            returnedVariable.Value = ReturnedValue;

            XmlAttribute isReturnedVariable = doc.CreateAttribute(string.Empty, Keys.IS_VALUE_RETURNED, string.Empty);
            isReturnedVariable.Value = IsValueReturned;

            taskElement.Attributes.Append(serviceUrl);
            taskElement.Attributes.Append(service);
            taskElement.Attributes.Append(webMethod);
            taskElement.Attributes.Append(mappingParams);
            taskElement.Attributes.Append(returnedVariable);
            taskElement.Attributes.Append(isReturnedVariable);

            doc.AppendChild(taskElement);
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="node">The node.</param>
        /// <param name="infoEvents">The info events.</param>
        void IDTSComponentPersist.LoadFromXML(XmlElement node, IDTSInfoEvents infoEvents)
        {
            if (node.Name != "SSISWCFTask")
            {
                throw new Exception("Wrong node name");
            }

            try
            {
                ServiceUrl = node.Attributes.GetNamedItem(Keys.SERVICE_URL).Value;
                ServiceContract = node.Attributes.GetNamedItem(Keys.SERVICE_CONTRACT).Value;
                OperationContract = node.Attributes.GetNamedItem(Keys.OPERATION_CONTRACT).Value;
                MappingParams = Serializer.DeSerializeFromXmlString(typeof(MappingParams), node.Attributes.GetNamedItem(Keys.MAPPING_PARAMS).Value);
                ReturnedValue = node.Attributes.GetNamedItem(Keys.RETURNED_VALUE).Value;
                IsValueReturned = node.Attributes.GetNamedItem(Keys.IS_VALUE_RETURNED).Value;
            }
            catch
            {
                throw new Exception("Unexpected task element when loading task.");
            }
        }

        #endregion
    }
}

