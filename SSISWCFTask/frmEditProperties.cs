using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.DataTransformationServices.Controls;
using Microsoft.SqlServer.Dts.Runtime;
using Microsoft.SqlServer.Dts.Runtime.Wrapper;
using SSISWCFTask100.WCFProxy;
using TaskHost = Microsoft.SqlServer.Dts.Runtime.TaskHost;
using Variable = Microsoft.SqlServer.Dts.Runtime.Variable;
using VariableDispenser = Microsoft.SqlServer.Dts.Runtime.VariableDispenser;

namespace SSISWCFTask100
{
    public partial class frmEditProperties : Form
    {
        #region Private Properties
        private readonly TaskHost _taskHost;
        private DynamicProxyFactory _dynamicProxyFactory;
        private string _withReturnValue = Keys.TRUE;
        #endregion

        #region Public Properties
        private Variables Variables
        {
            get { return _taskHost.Variables; }
        }

        #endregion

        #region .ctor
        public frmEditProperties(TaskHost taskHost, Connections connections)
        {
            InitializeComponent();

            grdParameters.DataError += grdParameters_DataError;

            _taskHost = taskHost;

            try
            {

                Cursor = Cursors.WaitCursor;

                cmbServices.Items.Clear();
                cmbMethods.Items.Clear();
                grdParameters.Rows.Clear();

                //Get URL's Service
                cmbURL.Items.AddRange(LoadVariables("System.String").ToArray());
                if (_taskHost.Properties[Keys.SERVICE_URL].GetValue(_taskHost) != null)
                    if (!string.IsNullOrEmpty(_taskHost.Properties[Keys.SERVICE_URL].GetValue(_taskHost).ToString()))
                    {
                        cmbURL.SelectedIndexChanged -= cmbURL_SelectedIndexChanged;
                        cmbServices.SelectedIndexChanged -= cmbServices_SelectedIndexChanged;
                        cmbMethods.SelectedIndexChanged -= cmbMethods_SelectedIndexChanged;


                        _dynamicProxyFactory = new DynamicProxyFactory(EvaluateExpression(_taskHost.Properties[Keys.SERVICE_URL].GetValue(_taskHost).ToString(), _taskHost.VariableDispenser).ToString());

                        cmbURL.Text = _taskHost.Properties[Keys.SERVICE_URL].GetValue(_taskHost).ToString();

                        //Get Services
                        foreach (var service in _dynamicProxyFactory.AvailableServices)
                        {
                            cmbServices.Items.Add(new ComboBoxObjectComboItem(service.Value, service.Key));
                        }


                        cmbServices.SelectedIndex = FindStringInComboBox(cmbServices, _taskHost.Properties[Keys.SERVICE_CONTRACT].GetValue(_taskHost).ToString(), -1);

                        //Get Operation Contracts by Service Contract

                        foreach (var method in ((WebServiceMethods)(((ComboBoxObjectComboItem)(cmbServices.SelectedItem)).ValueMemeber)))
                        {
                            cmbMethods.Items.Add(new ComboBoxObjectComboItem(method.WebServiceMethodParameters, method.Name));
                        }

                        cmbMethods.SelectedIndex = FindStringInComboBox(cmbMethods, _taskHost.Properties[Keys.OPERATION_CONTRACT].GetValue(_taskHost).ToString(), -1);

                        var webServiceMethods = from m in _dynamicProxyFactory.WebServiceMethods
                                                where m.Name == _taskHost.Properties[Keys.OPERATION_CONTRACT].GetValue(_taskHost).ToString()
                                                select new WebServiceMethod
                                                           {
                                                               Name = m.Name,
                                                               ResultType = m.ResultType
                                                           };

                        //Get returned variables
                        if (_taskHost.Properties[Keys.RETURNED_VALUE] != null)
                        {
                            if (!string.IsNullOrEmpty(_taskHost.Properties[Keys.RETURNED_VALUE].GetValue(_taskHost).ToString()))
                            {
                                cmbReturnVariable.Items.AddRange(LoadVariables((webServiceMethods.FirstOrDefault()).ResultType).ToArray());
                                cmbReturnVariable.SelectedIndex = FindStringInComboBox(cmbReturnVariable, _taskHost.Properties[Keys.RETURNED_VALUE].GetValue(_taskHost).ToString(), -1);
                            }
                        }

                        FillGridWithParams(_taskHost.Properties[Keys.MAPPING_PARAMS].GetValue(_taskHost) as MappingParams);

                        cmbURL.SelectedIndexChanged += cmbURL_SelectedIndexChanged;
                        cmbServices.SelectedIndexChanged += cmbServices_SelectedIndexChanged;
                        cmbMethods.SelectedIndexChanged += cmbMethods_SelectedIndexChanged;
                    }


            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
            finally
            {
                Cursor = Cursors.Arrow;
            }
        }

        void grdParameters_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }
        #endregion

        #region Methods

        /// <summary>
        /// Loads the variables.
        /// </summary>
        /// <param name="parameterInfo">The parameter info.</param>
        /// <returns></returns>
        private List<string> LoadVariables(string parameterInfo)
        {
            return Variables.Cast<Variable>().Where(variable => Type.GetTypeCode(Type.GetType(parameterInfo)) == variable.DataType).Select(variable => string.Format("@[{0}::{1}]", variable.Namespace, variable.Name)).ToList();
        }

        /// <summary>
        /// Loads the variables.
        /// </summary>
        /// <param name="parameterInfo">The parameter info.</param>
        /// <param name="selectedText">The selected text.</param>
        /// <returns></returns>
        private ComboBox LoadVariables(string parameterInfo, ref string selectedText)
        {
            var comboBox = new ComboBox();

            comboBox.Items.Add(string.Empty);

            foreach (Variable variable in Variables.Cast<Variable>().Where(variable => Type.GetTypeCode(Type.GetType(parameterInfo)) == variable.DataType))
            {
                comboBox.Items.Add(string.Format("@[{0}::{1}]", variable.Namespace, variable.Name));
            }

            return comboBox;
        }

        public int FindStringInComboBox(ComboBox comboBox, string searchTextItem, int startIndex)
        {
            if (startIndex >= comboBox.Items.Count)
                return -1;

            int indexPosition = comboBox.FindString(searchTextItem, startIndex);

            if (indexPosition <= startIndex)
                return -1;

            return comboBox.Items[indexPosition].ToString() == searchTextItem
                                    ? indexPosition
                                    : FindStringInComboBox(comboBox, searchTextItem, indexPosition);
        }

        /// <summary>
        /// This method evaluate expressions like @([System::TaskName] + [System::TaskID]) or any other operation created using
        /// ExpressionBuilder
        /// </summary>
        /// <param name="mappedParam">The mapped param.</param>
        /// <param name="variableDispenser">The variable dispenser.</param>
        /// <returns></returns>
        private static object EvaluateExpression(string mappedParam, VariableDispenser variableDispenser)
        {
            object variableObject;

            if (mappedParam.Contains("@"))
            {
                var expressionEvaluatorClass = new ExpressionEvaluatorClass
                                                   {
                                                       Expression = mappedParam
                                                   };

                expressionEvaluatorClass.Evaluate(DtsConvert.GetExtendedInterface(variableDispenser),
                                                  out variableObject,
                                                  false);
            }
            else
            {
                variableObject = mappedParam;
            }

            return variableObject;
        }

        #endregion

        #region Events

        private void btSave_Click(object sender, EventArgs e)
        {
            //Save the values
            _taskHost.Properties[Keys.SERVICE_URL].SetValue(_taskHost, cmbURL.Text);
            _taskHost.Properties[Keys.SERVICE_CONTRACT].SetValue(_taskHost, cmbServices.Text);
            _taskHost.Properties[Keys.OPERATION_CONTRACT].SetValue(_taskHost, cmbMethods.Text);

            var mappingParams = new MappingParams();

            mappingParams.AddRange(from DataGridViewRow mappingParam in grdParameters.Rows
                                   select new MappingParam
                                              {
                                                  Name = mappingParam.Cells[0].Value.ToString(),
                                                  Type = mappingParam.Cells[1].Value.ToString(),
                                                  Value = mappingParam.Cells[2].Value.ToString()
                                              });

            _taskHost.Properties[Keys.MAPPING_PARAMS].SetValue(_taskHost, mappingParams);
            _taskHost.Properties[Keys.RETURNED_VALUE].SetValue(_taskHost, cmbReturnVariable.Text);
            _taskHost.Properties[Keys.IS_VALUE_RETURNED].SetValue(_taskHost, _withReturnValue.ToString());
            DialogResult = DialogResult.OK;
            Close();
        }


        private void cmbURL_SelectedIndexChanged(object sender, EventArgs e)
        {
            _dynamicProxyFactory = new DynamicProxyFactory(EvaluateExpression(cmbURL.Text.Trim(), _taskHost.VariableDispenser).ToString());

            cmbServices.Items.Clear();
            cmbMethods.Items.Clear();
            grdParameters.Rows.Clear();

            foreach (var contract in _dynamicProxyFactory.Contracts)
            {
                foreach (KeyValuePair<string, WCFProxy.WebServiceMethods> service in _dynamicProxyFactory.AvailableServices)
                {
                    if (service.Key == contract.Name)
                    {
                        foreach (var serviceOperation in contract.Operations)
                        {
                            foreach (var val in service.Value)
                            {
                                if (serviceOperation.Name == val.Name)
                                {
                                    Console.WriteLine(service.Key);
                                    cmbServices.Items.Add(new ComboBoxObjectComboItem(service.Value, service.Key));

                                    /*foreach (var webServiceMethodParameters in val.WebServiceMethodParameters)
                                    {
                                        cmbServices.Items.Add(new ComboBoxObjectComboItem(webServiceMethodParameters, webServiceMethodParameters.Name));
                                    }*/
                                }
                            }
                        }
                    }
                }
            }
        }

        private void cmbServices_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbMethods.Items.Clear();
            grdParameters.Rows.Clear();

            foreach (var variable in ((WebServiceMethods)((ComboBoxObjectComboItem)(((ComboBox)sender).SelectedItem)).ValueMemeber))
            {
                cmbMethods.Items.Add(new ComboBoxObjectComboItem(variable.WebServiceMethodParameters, variable.Name));
            }
        }

        private void cmbMethods_SelectedIndexChanged(object sender, EventArgs e)
        {
            grdParameters.Rows.Clear();
            Cursor = Cursors.WaitCursor;

            FillGridWithParams();

            Cursor = Cursors.Arrow;
        }

        private void FillGridWithParams()
        {
            var webServiceMethodParameters = ((WebServiceMethodParameters)((ComboBoxObjectComboItem)((cmbMethods.SelectedItem))).ValueMemeber);

            foreach (var method in _dynamicProxyFactory.WebServiceMethods)
            {
                if (method.Name != cmbMethods.Text)
                    continue;

                string selectedText = string.Empty;

                webServiceMethodParameters = method.WebServiceMethodParameters;

                cmbReturnVariable.Items.Clear();

                cmbReturnVariable.Items.AddRange(LoadVariables(method.ResultType, ref selectedText).Items.Cast<string>().ToList().Where(s => s.Contains("User")).ToArray());
                cmbReturnVariable.SelectedIndex = FindStringInComboBox(cmbReturnVariable, selectedText, -1);

                if (method.ResultType == "System.Void")
                {
                    lbOutputValue.Visible = cmbReturnVariable.Visible = false;
                    _withReturnValue = Keys.FALSE;
                }
                else
                {
                    lbOutputValue.Visible = cmbReturnVariable.Visible = true;
                    _withReturnValue = Keys.TRUE;
                }

                break;
            }

            if (webServiceMethodParameters != null)
                foreach (var webServiceMethodparameter in webServiceMethodParameters)
                {

                    int index = grdParameters.Rows.Add();

                    DataGridViewRow row = grdParameters.Rows[index];

                    row.Cells["grdColParams"] = new DataGridViewTextBoxCell
                                                    {
                                                        Value = webServiceMethodparameter.Name,
                                                        Tag = webServiceMethodparameter.Name,
                                                    };

                    row.Cells["grdColDirection"] = new DataGridViewTextBoxCell
                                                       {
                                                           Value = webServiceMethodparameter.Type
                                                       };

                    row.Cells["grdColVars"] = LoadVariables(webServiceMethodparameter);
                    row.Cells["grdColExpression"] = new DataGridViewButtonCell();
                }
        }

        private void FillGridWithParams(MappingParams mappingParams)
        {
            if (mappingParams != null)
                foreach (var mappingParam in mappingParams)
                {

                    int index = grdParameters.Rows.Add();

                    DataGridViewRow row = grdParameters.Rows[index];

                    row.Cells["grdColParams"] = new DataGridViewTextBoxCell
                    {
                        Value = mappingParam.Name,
                        Tag = mappingParam.Name,
                    };

                    row.Cells["grdColDirection"] = new DataGridViewTextBoxCell
                    {
                        Value = mappingParam.Type
                    };

                    row.Cells["grdColVars"] = LoadVariables(mappingParam);
                    row.Cells["grdColExpression"] = new DataGridViewButtonCell();


                    if (_withReturnValue == Keys.FALSE)
                    {
                        lbOutputValue.Visible = cmbReturnVariable.Visible = false;
                        _withReturnValue = Keys.FALSE;
                    }
                    else
                    {
                        lbOutputValue.Visible = cmbReturnVariable.Visible = true;
                        _withReturnValue = Keys.TRUE;
                    }
                }
        }

        private void btGO_Click(object sender, EventArgs e)
        {
            _dynamicProxyFactory = new DynamicProxyFactory(EvaluateExpression(cmbURL.Text.Trim(), _taskHost.VariableDispenser).ToString());
        }

        private DataGridViewComboBoxCell LoadVariables(WebServiceMethodParameter parameterInfo)
        {
            var comboBoxCell = new DataGridViewComboBoxCell();

            foreach (Variable variable in Variables.Cast<Variable>().Where(variable => Type.GetTypeCode(Type.GetType(parameterInfo.Type)) == variable.DataType))
            {
                comboBoxCell.Items.Add(string.Format("@[{0}::{1}]", variable.Namespace, variable.Name));
            }

            return comboBoxCell;
        }

        private DataGridViewComboBoxCell LoadVariables(MappingParam parameterInfo)
        {
            var comboBoxCell = new DataGridViewComboBoxCell();

            foreach (Variable variable in Variables.Cast<Variable>().Where(variable => Type.GetTypeCode(Type.GetType(parameterInfo.Type)) == variable.DataType))
            {
                comboBoxCell.Items.Add(string.Format("@[{0}::{1}]", variable.Namespace, variable.Name));
            }

            comboBoxCell.Items.Add(parameterInfo.Value);
            comboBoxCell.Value = parameterInfo.Value;

            return comboBoxCell;
        }

        private void grdParameters_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 3:
                    {
                        using (var expressionBuilder = ExpressionBuilder.Instantiate(_taskHost.Variables, _taskHost.VariableDispenser, Type.GetType((grdParameters.Rows[e.RowIndex].Cells[1]).Value.ToString()), string.Empty))
                        {
                            if (expressionBuilder.ShowDialog() == DialogResult.OK)
                            {
                                ((DataGridViewComboBoxCell)grdParameters.Rows[e.RowIndex].Cells[e.ColumnIndex - 1]).Items.Add(expressionBuilder.Expression);
                                grdParameters.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value = expressionBuilder.Expression;
                            }
                        }
                    }

                    break;
            }
        }
        #endregion
    }
}
