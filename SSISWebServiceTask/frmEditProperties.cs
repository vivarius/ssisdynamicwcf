using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.DataTransformationServices.Controls;
using Microsoft.SqlServer.Dts.Runtime;
using Microsoft.SqlServer.Dts.Runtime.Wrapper;
using TaskHost = Microsoft.SqlServer.Dts.Runtime.TaskHost;
using Variable = Microsoft.SqlServer.Dts.Runtime.Variable;
using VariableDispenser = Microsoft.SqlServer.Dts.Runtime.VariableDispenser;

namespace SSISWebServiceTask100
{
    public partial class frmEditProperties : Form
    {
        #region Private Properties
        private readonly TaskHost _taskHost;
        private WSDLHandler _wsdlHandler;
        private int _withReturnValue = 1;
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
                if (_taskHost.Properties[NamedStringMembers.SERVICE_URL].GetValue(_taskHost) != null)
                    if (!string.IsNullOrEmpty(_taskHost.Properties[NamedStringMembers.SERVICE_URL].GetValue(_taskHost).ToString()))
                    {
                        cmbURL.SelectedIndexChanged -= cmbURL_SelectedIndexChanged;
                        cmbServices.SelectedIndexChanged -= cmbServices_SelectedIndexChanged;
                        cmbMethods.SelectedIndexChanged -= cmbMethods_SelectedIndexChanged;


                        _wsdlHandler = new WSDLHandler(new Uri(EvaluateExpression(_taskHost.Properties[NamedStringMembers.SERVICE_URL].GetValue(_taskHost).ToString(),
                                                                                  _taskHost.VariableDispenser).ToString()));

                        cmbURL.Text = _taskHost.Properties[NamedStringMembers.SERVICE_URL].GetValue(_taskHost).ToString();

                        //Get Services
                        cmbServices.Items.AddRange(_wsdlHandler.AvailableServices.ToArray());
                        cmbServices.SelectedIndex = FindStringInComboBox(cmbServices, _taskHost.Properties[NamedStringMembers.SERVICE].GetValue(_taskHost).ToString(), -1);

                        //Get Methods by service
                        cmbMethods.Items.AddRange(_wsdlHandler.GetServiceMethods(_taskHost.Properties[NamedStringMembers.SERVICE].GetValue(_taskHost).ToString()).ToArray());
                        cmbMethods.SelectedIndex = FindStringInComboBox(cmbMethods, _taskHost.Properties[NamedStringMembers.WEBMETHOD].GetValue(_taskHost).ToString(), -1);

                        var webServiceMethod = from m in _wsdlHandler.WebServiceMethods
                                               where m.Name == _taskHost.Properties[NamedStringMembers.WEBMETHOD].GetValue(_taskHost).ToString()
                                               select new WebServiceMethod
                                                          {
                                                              Name = m.Name,
                                                              ResultType = m.ResultType
                                                          };

                        //Get returned variables
                        if (_taskHost.Properties[NamedStringMembers.RETURNED_VALUE] != null)
                        {
                            if (!string.IsNullOrEmpty(_taskHost.Properties[NamedStringMembers.RETURNED_VALUE].GetValue(_taskHost).ToString()))
                            {
                                cmbReturnVariable.Items.AddRange(LoadVariables((webServiceMethod.FirstOrDefault()).ResultType).ToArray());
                                cmbReturnVariable.SelectedIndex = FindStringInComboBox(cmbReturnVariable, _taskHost.Properties[NamedStringMembers.RETURNED_VALUE].GetValue(_taskHost).ToString(), -1);
                            }
                        }

                        FillGridWithParams(_taskHost.Properties[NamedStringMembers.MAPPING_PARAMS].GetValue(_taskHost) as MappingParams);

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
            throw new NotImplementedException();
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
            return
                Variables.Cast<Variable>().Where(
                    variable => Type.GetTypeCode(Type.GetType(parameterInfo)) == variable.DataType).Select(
                        variable => string.Format("@[{0}::{1}]", variable.Namespace, variable.Name)).ToList();
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
            _taskHost.Properties[NamedStringMembers.SERVICE_URL].SetValue(_taskHost, cmbURL.Text);
            _taskHost.Properties[NamedStringMembers.SERVICE].SetValue(_taskHost, cmbServices.Text);
            _taskHost.Properties[NamedStringMembers.WEBMETHOD].SetValue(_taskHost, cmbMethods.Text);

            var mappingParams = new MappingParams();

            mappingParams.AddRange(from DataGridViewRow mappingParam in grdParameters.Rows
                                   select new MappingParam
                                              {
                                                  Name = mappingParam.Cells[0].Value.ToString(),
                                                  Type = mappingParam.Cells[1].Value.ToString(),
                                                  Value = mappingParam.Cells[2].Value.ToString()
                                              });

            _taskHost.Properties[NamedStringMembers.MAPPING_PARAMS].SetValue(_taskHost, mappingParams);
            _taskHost.Properties[NamedStringMembers.RETURNED_VALUE].SetValue(_taskHost, cmbReturnVariable.Text);
            _taskHost.Properties[NamedStringMembers.IS_VALUE_RETURNED].SetValue(_taskHost, _withReturnValue.ToString());
            DialogResult = DialogResult.OK;
            Close();
        }


        private void cmbURL_SelectedIndexChanged(object sender, EventArgs e)
        {
            _wsdlHandler = new WSDLHandler(new Uri(EvaluateExpression(cmbURL.Text.Trim(), _taskHost.VariableDispenser).ToString()));
            cmbServices.Items.Clear();
            cmbMethods.Items.Clear();
            grdParameters.Rows.Clear();
            cmbServices.Items.AddRange(_wsdlHandler.AvailableServices.ToArray());
        }

        private void cmbServices_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbMethods.Items.Clear();
            grdParameters.Rows.Clear();
            cmbMethods.Items.AddRange(_wsdlHandler.GetServiceMethods(cmbServices.Text).ToArray());
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
            WebServiceMethodParameters webServiceMethodParameters = new WebServiceMethodParameters();

            foreach (var method in _wsdlHandler.WebServiceMethods)
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
                    _withReturnValue = 0;
                }
                else
                {
                    lbOutputValue.Visible = cmbReturnVariable.Visible = true;
                    _withReturnValue = 1;
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


                    if (_withReturnValue == 0)
                    {
                        lbOutputValue.Visible = cmbReturnVariable.Visible = false;
                        _withReturnValue = 0;
                    }
                    else
                    {
                        lbOutputValue.Visible = cmbReturnVariable.Visible = true;
                        _withReturnValue = 1;
                    }
                }
        }

        private void btGO_Click(object sender, EventArgs e)
        {
            _wsdlHandler = new WSDLHandler(new Uri(EvaluateExpression(cmbURL.Text.Trim(), _taskHost.VariableDispenser).ToString()));
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
